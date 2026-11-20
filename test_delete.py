# -*- coding: utf-8 -*-
"""Проверка удаления строк из Excel и пересчёта порядковых номеров."""
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))

import app

tmp_path = BASE_DIR / "test_delete.xlsx"
if tmp_path.exists():
    tmp_path.unlink()

app.init_excel(tmp_path)
app.append_record(tmp_path, "101", 100.00)
app.append_record(tmp_path, "102", 200.00)
app.append_record(tmp_path, "103", 300.00)

records = app.read_records(tmp_path)
assert len(records) == 3, records
assert [r[2] for r in records] == ["101", "102", "103"], records

# Удаляем вторую запись (строка Excel = 3)
deleted = app.delete_records(tmp_path, [3])
print("Удалено:", deleted)
assert deleted == 1

records = app.read_records(tmp_path)
print("Осталось:", records)
assert len(records) == 2, records
assert [r[2] for r in records] == ["101", "103"], records
# Порядковые номера пересчитаны
assert [r[0] for r in records] == [1, 2], records

# Удаление нескольких строк сразу
app.append_record(tmp_path, "104", 400.00)
app.append_record(tmp_path, "105", 500.00)
deleted = app.delete_records(tmp_path, [2, 4])  # 101 и 104
records = app.read_records(tmp_path)
print("После удаления нескольких:", records)
assert deleted == 2
assert [r[2] for r in records] == ["103", "105"], records
assert [r[0] for r in records] == [1, 2], records

# Удаление несуществующей строки ничего не ломает
assert app.delete_records(tmp_path, [99]) == 0
assert app.delete_records(tmp_path, []) == 0

tmp_path.unlink()
print("OK: удаление строк работает корректно")
