# -*- coding: utf-8 -*-
"""Интеграционный тест: генерирует изображение чека с русским текстом и проверяет OCR."""
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))

from PIL import Image, ImageDraw, ImageFont
import app

# 1. Генерируем тестовое изображение чека
img = Image.new("RGB", (600, 700), "white")
draw = ImageDraw.Draw(img)
font = ImageFont.truetype("consola.ttf", 28)
lines = [
    "ООО РОМАШКА",
    "СМЕНА 123 КАССА 2",
    "Молоко 1л        89.90",
    "Хлеб             45.00",
    "ИТОГ: 134.90",
    "НАЛИЧНЫМИ       150.00",
    "СДАЧА            15.10",
]
y = 30
for line in lines:
    draw.text((40, y), line, fill="black", font=font)
    y += 50
img.save(str(BASE_DIR / "test_receipt.png"))

# 2. Распознаём через Tesseract (использует путь из app.py)
text = app.extract_text(img)
print("=== РАСПОЗНАННЫЙ ТЕКСТ ===")
print(text)
print("=== РЕЗУЛЬТАТ ===")
norm = app.normalize_text(text)
shift = app.find_shift_number(norm)
amount = app.find_amount(norm)
print("Номер смены:", shift)
print("Сумма:", amount)

assert shift == "123", f"Смена неверна: {shift}"
assert amount is not None and abs(amount - 134.90) < 0.01, f"Сумма неверна: {amount}"
print("OK: OCR и парсинг работают корректно")
