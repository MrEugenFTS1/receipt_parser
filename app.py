# -*- coding: utf-8 -*-
"""
Чек-сканер: Streamlit-приложение.
Пользователь загружает фото чека -> программа распознаёт текст (Tesseract OCR, рус+англ),
находит номер смены и сумму чека в рублях и дописывает запись в Excel-файл (receipts.xlsx).
После обработки можно загрузить следующий чек.
"""

import os
import re
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, Side, PatternFill
from PIL import Image
import pytesseract
import streamlit as st

# ----------------------------------------------------------------------
# Настройки
# ----------------------------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent
EXCEL_PATH = BASE_DIR / "receipts.xlsx"

# Tesseract поставляется вместе с проектом в папке tesseract-env (устанавливается micromamba).
# Если папки нет — используется Tesseract из системного PATH.
_BUNDLED_TESSERACT = BASE_DIR / "tesseract-env" / "Library" / "bin" / "tesseract.exe"
if _BUNDLED_TESSERACT.exists():
    pytesseract.pytesseract.tesseract_cmd = str(_BUNDLED_TESSERACT)
    os.environ["TESSDATA_PREFIX"] = str(BASE_DIR / "tesseract-env" / "share" / "tessdata")

TESSERACT_CONFIG = "--oem 1 --psm 6 -l rus+eng"

EXCEL_HEADERS = ["№ п/п", "Дата обработки", "Номер смены", "Сумма чека, руб."]

# ----------------------------------------------------------------------
# Работа с Excel
# ----------------------------------------------------------------------

thin = Side(style="thin", color="999999")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
HEADER_FILL = PatternFill("solid", fgColor="D9E1F2")
HEADER_FONT = Font(bold=True)


def init_excel(path: Path) -> None:
    """Создаёт файл Excel с заголовками, если его ещё нет."""
    if path.exists():
        return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Чеки"
    for col, header in enumerate(EXCEL_HEADERS, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 16
    wb.save(path)

#Новое изменение добавлен комментарий
def append_record(path: Path, shift_number: str, amount: float) -> int:
    """Дописывает строку в Excel. Возвращает номер строки."""
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    row = ws.max_row + 1
    seq = row - 1  # порядковый номер записи
    values = [seq, datetime.now().strftime("%d.%m.%Y %H:%M"), shift_number, amount]
    for col, value in enumerate(values, start=1):
        cell = ws.cell(row=row, column=col, value=value)
        cell.border = BORDER
        if col == 4:
            cell.number_format = "#,##0.00"
            cell.alignment = Alignment(horizontal="right")
    wb.save(path)
    return seq


def read_records(path: Path) -> list[tuple]:
    """Читает все записи из Excel (для таблицы на странице)."""
    if not path.exists():
        return []
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    records = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if any(v is not None for v in row):
            records.append(row)
    return records


def delete_records(path: Path, excel_rows: list[int]) -> int:
    """
    Удаляет строки из Excel по номерам строк (1-based, строка 1 — заголовок).
    После удаления порядковые номера (№ п/п) пересчитываются.
    Возвращает количество удалённых строк.
    """
    if not excel_rows:
        return 0
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    # Удаляем снизу вверх, чтобы номера строк не сдвигались
    deleted = 0
    for row in sorted({r for r in excel_rows}, reverse=True):
        if 2 <= row <= ws.max_row:
            ws.delete_rows(row, 1)
            deleted += 1
    # Пересчитываем порядковые номера
    for seq, row in enumerate(range(2, ws.max_row + 1), start=1):
        ws.cell(row=row, column=1, value=seq)
    wb.save(path)
    return deleted


# ----------------------------------------------------------------------
# OCR и парсинг
# ----------------------------------------------------------------------

def extract_text(image: Image.Image) -> str:
    """Распознаёт текст на изображении."""
    # Чёрно-белая версия повышает точность распознавания чеков
    gray = image.convert("L")
    return pytesseract.image_to_string(gray, config=TESSERACT_CONFIG)


# Латинские буквы, визуально совпадающие с кириллическими, -> кириллица.
# Нужен для OCR: "СМЕНА" может быть распознана как "CMEHA", "СМ" как "CM" и т.д.
_LATIN_TO_CYRILLIC = str.maketrans("ABCEHKMOPTX", "АВСЕНКМОРТХ")

# Кириллические буквы, которые OCR принимает вместо цифр, -> цифры.
_LETTERS_TO_DIGITS = str.maketrans({"О": "0", "О": "0", "З": "3"})


def normalize_text(text: str) -> str:
    """Приводит OCR-текст к удобному для поиска виду."""
    # Приводим к единому регистру и заменяем латинские буквы-двойники на кириллицу
    return text.replace("ё", "е").upper().translate(_LATIN_TO_CYRILLIC)


def find_shift_number(normalized: str) -> str | None:
    """
    Ищет номер смены. Возможные варианты на чеках:
      СМЕНА 123, СМЕНЫ 123, СМЕНА: 123, СМЕНА№123, СМЕНА N 123, СМЕНА М 123
      (буква M/М/Н/N вместо №, до или после номера) и т.п.
    Буквы, похожие на цифры (О -> 0, З -> 3), в номере исправляются.
    """
    patterns = [
        # СМЕНА 123 / СМЕНА 123 N / СМЕНА N 123 / СМЕНА№123 / СМЕНА М123 / СМЕНА: 123
        r"СМЕН[ЫАЯИ]\s*[№#NMНМО]*\s*[.:\-]?\s*(\d{1,6})\s*[N№#MНМ]?",
        r"СМЕНА\s*[№#NMНМО]\s*[.:]?\s*(\d{1,6})",
        r"СМ\s*[№#NMНМО]\s*[.:]?\s*(\d{1,6})",
    ]
    for pattern in patterns:
        match = re.search(pattern, normalized)
        if match:
            number = match.group(1).translate(_LETTERS_TO_DIGITS)
            return str(int(number))  # убираем ведущие нули
    return None


def find_amount(normalized: str) -> float | None:
    """
    Ищет итоговую сумму чека. Приоритет: ИТОГ, ИТОГО, ВСЕГО.
    Суммы вида: ИТОГ: 1250.00 / ИТОГО=1250,50 / ВСЕГО 1 250 РУБ
    """
    keywords = ["ИТОГ", "ИТОГО", "ВСЕГО", "СУММА", "ОПЛАТА"]
    lines = normalized.splitlines()
    # Сначала ищем строки с ключевыми словами
    for keyword in keywords:
        for line in lines:
            if keyword in line:
                value = _parse_number_in_line(line)
                if value is not None:
                    return value
    # Ищем в самих строках ключевых слов соседние строки (сумма на следующей строке)
    for keyword in keywords:
        for i, line in enumerate(lines):
            if keyword in line and i + 1 < len(lines):
                value = _parse_number_in_line(lines[i + 1])
                if value is not None:
                    return value
    return None


def _parse_number_in_line(line: str) -> float | None:
    """Пытается извлечь число (возможно с пробелами-разделителями тысяч)."""
    candidates = [
        r"(\d{1,3}(?:\s\d{3})+[.,]\d{2})",   # 1 250.50 или 1 250,50
        r"(\d{1,3}(?:\s\d{3})+)",            # 1 250
        r"(\d+[.,]\d{1,2})",                 # 1250.50
        r"(\d{2,})",                         # 1250
    ]
    for pattern in candidates:
        match = re.search(pattern, line)
        if match:
            raw = match.group(1).replace(" ", "").replace(",", ".")
            try:
                value = float(raw)
                if 0 < value < 10_000_000:
                    return round(value, 2)
            except ValueError:
                continue
    return None


# ----------------------------------------------------------------------
# Интерфейс Streamlit
# ----------------------------------------------------------------------

st.set_page_config(page_title="Чек-сканер", page_icon="🧾", layout="centered")

st.title("🧾 Чек-сканер")
st.caption("Загрузите фото чека — программа найдёт номер смены и сумму и сохранит запись в Excel.")

init_excel(EXCEL_PATH)

uploaded_file = st.file_uploader(
    "Изображение чека (JPG, PNG)",
    type=["jpg", "jpeg", "png", "bmp", "webp"],
)

if uploaded_file is not None:
    image = Image.open(uploaded_file)
    st.image(image, caption="Загруженный чек", width="stretch")

    with st.spinner("Распознаю текст…"):
        try:
            raw_text = extract_text(image)
        except pytesseract.TesseractError as e:
            st.error(
                "Ошибка Tesseract OCR. Убедитесь, что Tesseract установлен и доступен, "
                "а русский язык (rus) добавлен. Подробнее — в README.md.\n\n"
                f"Детали: {e}"
            )
            st.stop()

    normalized = normalize_text(raw_text)

    with st.expander("Распознанный текст"):
        st.text(raw_text)

    shift_number = find_shift_number(normalized)
    amount = find_amount(normalized)

    col1, col2 = st.columns(2)
    with col1:
        st.metric("Номер смены", shift_number if shift_number else "не найден")
    with col2:
        st.metric(
            "Сумма чека",
            f"{amount:,.2f} ₽".replace(",", " ").replace(".", ",") if amount is not None else "не найдена",
        )

    if shift_number is None or amount is None:
        st.warning(
            "Не удалось найти номер смены и/или сумму автоматически. "
            "Проверьте распознанный текст выше — при необходимости можно ввести данные вручную."
        )

    with st.form("manual_fix"):
        st.subheader("Проверьте данные перед сохранением")
        shift_input = st.text_input(
            "Номер смены",
            value=shift_number or "",
            help="Поле заполнено автоматически, при необходимости исправьте",
        )
        amount_input = st.text_input(
            "Сумма чека, руб.",
            value=f"{amount:.2f}".replace(".", ",") if amount is not None else "",
            help="Например: 1250,50",
        )
        submitted = st.form_submit_button("💾 Сохранить в Excel")

        if submitted:
            if not shift_input.strip():
                st.error("Укажите номер смены.")
            else:
                try:
                    amount_value = float(amount_input.replace(" ", "").replace(",", "."))
                except ValueError:
                    st.error("Сумма чека указана некорректно. Пример: 1250,50")
                else:
                    seq = append_record(EXCEL_PATH, shift_input.strip(), amount_value)
                    st.success(f"Запись №{seq} сохранена в файл {EXCEL_PATH.name}!")
                    st.balloons()
                    st.info("Можно загружать следующий чек.")

st.divider()
st.subheader("Сохранённые записи")

# Сообщение после удаления переживает перезапуск страницы
flash = st.session_state.pop("records_flash", None)
if flash:
    st.success(flash)

records = read_records(EXCEL_PATH)
if records:
    import pandas as pd

    df = pd.DataFrame(records, columns=EXCEL_HEADERS)
    event = st.dataframe(
        df,
        width="stretch",
        hide_index=True,
        on_select="rerun",
        selection_mode="multi-row",
        key="records_table",
    )
    # Индексы выбранных строк DataFrame (0-based) -> номера строк Excel (строка 1 — заголовок)
    selected_rows = list(event.selection.rows)

    col_delete, col_hint = st.columns([1, 2])
    with col_delete:
        delete_clicked = st.button(
            "🗑 Удалить выбранную строку",
            disabled=not selected_rows,
            help="Выделите одну или несколько строк в таблице выше",
        )
    with col_hint:
        if selected_rows:
            st.caption(f"Выбрано строк: {len(selected_rows)}")
        else:
            st.caption("Кликните по строке таблицы, чтобы выбрать её для удаления")

    if delete_clicked and selected_rows:
        excel_rows = [i + 2 for i in selected_rows]
        deleted = delete_records(EXCEL_PATH, excel_rows)
        st.session_state["records_flash"] = f"Удалено записей: {deleted}."
        st.rerun()

    total = df["Сумма чека, руб."].sum()
    st.markdown(f"**Всего записей:** {len(records)}  |  **Общая сумма:** {total:,.2f} ₽".replace(",", " ").replace(".", ","))
    with open(EXCEL_PATH, "rb") as f:
        st.download_button(
            "⬇️ Скачать Excel-файл",
            data=f.read(),
            file_name="receipts.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("Записей пока нет. Загрузите первый чек.")
