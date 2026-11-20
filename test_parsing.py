# -*- coding: utf-8 -*-
"""Быстрая проверка функций парсинга из app.py без GUI."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from app import find_shift_number, find_amount, normalize_text

tests_text = """
         ООО РОМАШКА
СМЕНА 123   КАССА 2
Молоко 1л         89.90
Хлеб              45.00
ИТОГ: 134.90
НАЛИЧНЫМИ 150.00
СДАЧА 15.10
"""

norm = normalize_text(tests_text)
shift = find_shift_number(norm)
amount = find_amount(norm)
print("Номер смены:", shift)
print("Сумма:", amount)
assert shift == "123", shift
assert amount == 134.90, amount

tests_text2 = "СМЕНА№456\nВСЕГО 12 450,00 РУБ\n"
norm2 = normalize_text(tests_text2)
shift2 = find_shift_number(norm2)
amount2 = find_amount(norm2)
print("Номер смены 2:", shift2)
print("Сумма 2:", amount2)
assert shift2 == "456", shift2
assert amount2 == 12450.00, amount2

# Вариант, когда OCR выдаёт букву N после номера: "Смена 123 N"
tests_text3 = "СМЕНА 123 N\nИТОГО 500,00\n"
norm3 = normalize_text(tests_text3)
shift3 = find_shift_number(norm3)
amount3 = find_amount(norm3)
print("Номер смены 3:", shift3)
print("Сумма 3:", amount3)
assert shift3 == "123", shift3
assert amount3 == 500.00, amount3

# Вариант "СМЕНА N 456"
norm4 = normalize_text("СМЕНА N 456 ИТОГ 2000")
shift4 = find_shift_number(norm4)
assert shift4 == "456", shift4

# Вариант "СМЕНА M123" (№ ошибочно распознан как M)
norm5 = normalize_text("СМЕНА M123 ИТОГ 999,99")
shift5 = find_shift_number(norm5)
amount5 = find_amount(norm5)
print("Номер смены 5:", shift5)
print("Сумма 5:", amount5)
assert shift5 == "123", shift5
assert amount5 == 999.99, amount5

# Вариант "СМЕНА M 321" и "СМЕНА 654 M"
norm6 = normalize_text("СМЕНА M 321\nИТОГ 100.00")
assert find_shift_number(norm6) == "321"
norm7 = normalize_text("СМЕНА 654 M\nИТОГ 100.00")
assert find_shift_number(norm7) == "654"

# Кириллическая М (выглядит как латинская M): "Смена М 123"
norm8 = normalize_text("Смена М 123\nИТОГ 300,00")
shift8 = find_shift_number(norm8)
print("Номер смены 8:", shift8)
assert shift8 == "123", shift8

# Латинские буквы-двойники: "CMEHA 777" вместо "СМЕНА 777"
norm9 = normalize_text("CMEHA 777\nИТОГ 50.00")
shift9 = find_shift_number(norm9)
print("Номер смены 9:", shift9)
assert shift9 == "777", shift9

# Буква О вместо нуля в номере: "СМЕНА О45" -> 45
norm10 = normalize_text("СМЕНА О45\nИТОГ 10.00")
shift10 = find_shift_number(norm10)
print("Номер смены 10:", shift10)
assert shift10 == "45", shift10

# Кириллическая Н вместо N: "СМЕНА Н 555"
norm11 = normalize_text("СМЕНА Н 555\nИТОГ 10.00")
assert find_shift_number(norm11) == "555"

print("OK: парсинг работает корректно")
