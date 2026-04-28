"""Проверка выходных файлов"""
from openpyxl import load_workbook
import os

dirs = os.listdir('output_departments')
first_dir = dirs[0]
file_path = os.path.join('output_departments', first_dir, first_dir + '.xlsx')

wb = load_workbook(file_path)
ws = wb.active

print(f'Файл: {first_dir}.xlsx')
print(f'Строк: {ws.max_row}, Колонок: {ws.max_column}')
print('\nПервые 12 строк (все колонки):')

for i in range(1, 13):
    row_data = []
    for j in range(1, 9):
        val = ws.cell(i, j).value
        row_data.append(str(val)[:25] if val else '')
    print(f'Строка {i:2d}: {row_data}')
