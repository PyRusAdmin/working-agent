"""
Скрипт разбивки штатной расстановки по подразделениям.

Использование:
    python split_by_department.py [путь_к_файлу]

Если путь не указан, ищет файл your_file_modified.xlsx в текущей папке.
Результат: папка output_departments/ с подпапками по каждому подразделению.
"""
import sys
import os
import re
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import copy

# ─────────────────────────── Настройки ───────────────────────────
HEADER_ROWS = 9          # строки 0..8 — шапка отчёта (индексы pandas)
OUTPUT_DIR  = "output_departments"


def sanitize_name(name: str) -> str:
    """Убирает из имени символы, недопустимые в имени папки/файла."""
    name = name.strip()
    name = re.sub(r'[\\/:*?"<>|]', '_', name)
    return name[:80]   # ограничение длины


def find_department_boundaries(df: pd.DataFrame):
    """
    Возвращает список (row_index, dept_name) для каждого
    первого вхождения подразделения-заголовка.

    Признак строки-заголовка подразделения:
      - col 0: непустой текст, нет '/', нет ',', не 'ВАКАНСИЯ'
      - col 5: числовое значение (Запланировано)
    """
    seen = set()
    boundaries = []

    skip_values = {
        'Подразделение', 'Позиция', 'Сотрудник, Состояние',
        'nan', 'Штатная расстановка', 'Организация', 'Дата отчета'
    }

    for i, row in df.iterrows():
        val0 = str(row[0]).strip() if pd.notna(row[0]) else ''
        val5 = row[5]

        if not val0 or val0 in skip_values:
            continue
        if '/' in val0 or ',' in val0 or val0 == 'ВАКАНСИЯ':
            continue
        if pd.isna(val5):
            continue
        try:
            float(val5)
        except (ValueError, TypeError):
            continue

        if val0 not in seen:
            seen.add(val0)
            boundaries.append((i, val0))

    return boundaries


def copy_row_style(src_row, dst_row):
    """Копирует стили ячеек из src_row в dst_row (openpyxl rows)."""
    for src_cell, dst_cell in zip(src_row, dst_row):
        dst_cell.value = src_cell.value
        if src_cell.has_style:
            dst_cell.font      = copy.copy(src_cell.font)
            dst_cell.fill      = copy.copy(src_cell.fill)
            dst_cell.border    = copy.copy(src_cell.border)
            dst_cell.alignment = copy.copy(src_cell.alignment)
            dst_cell.number_format = src_cell.number_format


def save_department_xlsx(src_wb, src_ws, row_indices: list, dept_name: str, out_path: str):
    """
    Создаёт новый xlsx-файл из строк src_ws с номерами row_indices
    (1-based, как в openpyxl).
    Сохраняет ширину столбцов.
    """
    from openpyxl import Workbook

    wb_new = Workbook()
    ws_new = wb_new.active
    ws_new.title = "Штатная расстановка"

    # Копируем ширину столбцов
    for col_letter, col_dim in src_ws.column_dimensions.items():
        ws_new.column_dimensions[col_letter].width = col_dim.width

    # Копируем строки
    dst_row_num = 1
    for src_row_num in row_indices:
        src_row = src_ws[src_row_num]
        dst_row = ws_new[dst_row_num]
        copy_row_style(src_row, dst_row)
        dst_row_num += 1

    wb_new.save(out_path)


def main():
    # ── Путь к файлу ──
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
    else:
        input_file = "output/your_file_modified.xlsx"

    if not os.path.exists(input_file):
        print(f"[ОШИБКА] Файл не найден: {input_file}")
        sys.exit(1)

    print(f"Читаю файл: {input_file}")

    # ── Загрузка через pandas для анализа структуры ──
    df = pd.read_excel(input_file, header=None)
    total_rows = len(df)

    boundaries = find_department_boundaries(df)
    print(f"Найдено подразделений: {len(boundaries)}")

    # ── Загрузка через openpyxl для сохранения стилей ──
    wb_src = load_workbook(input_file)
    ws_src = wb_src.active

    # Строки шапки: pandas 0..HEADER_ROWS-1  →  openpyxl 1..HEADER_ROWS
    header_openpyxl = list(range(1, HEADER_ROWS + 1))

    # ── Создание выходной папки ──
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # ── Разбивка по подразделениям ──
    for idx, (pd_start, dept_name) in enumerate(boundaries):

        # Конец секции = начало следующего подразделения или конец файла
        if idx + 1 < len(boundaries):
            pd_end = boundaries[idx + 1][0]   # не включая
        else:
            pd_end = total_rows

        # Строки секции в терминах pandas → openpyxl (+1 смещение, +HEADER_ROWS не нужен т.к. df без пропусков)
        # pandas индекс i соответствует openpyxl строке i+1
        dept_openpyxl = list(range(pd_start + 1, pd_end + 1))

        if not dept_openpyxl:
            continue

        row_indices = header_openpyxl + dept_openpyxl

        # Считаем кол-во вакансий для информации
        dept_slice = df.iloc[pd_start:pd_end]
        vacancy_count = (dept_slice[0] == 'ВАКАНСИЯ').sum()

        # Создаём папку подразделения
        folder_name = sanitize_name(dept_name)
        dept_dir = os.path.join(OUTPUT_DIR, folder_name)
        os.makedirs(dept_dir, exist_ok=True)

        # Имя файла
        out_file = os.path.join(dept_dir, f"{folder_name}.xlsx")

        save_department_xlsx(wb_src, ws_src, row_indices, dept_name, out_file)

        rows_in_section = len(dept_openpyxl)
        print(f"  [{idx+1:02d}/{len(boundaries)}] {dept_name[:55]:<55} "
              f"строк: {rows_in_section:4d}  вакансий: {vacancy_count}")

    print(f"\nГотово! Файлы сохранены в папку: {OUTPUT_DIR}/")
    print(f"Всего файлов: {len(boundaries)}")


if __name__ == "__main__":
    main()
