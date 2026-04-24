import openpyxl
from loguru import logger
from openpyxl.styles import Font
from openpyxl.utils import range_boundaries
from openpyxl.cell.cell import MergedCell

file_path = 'input/Штатная_расстановка.xlsx'
output_path = 'output/your_file_modified.xlsx'


def removes_grouping():
    """Снимает группировку со всех строк"""
    logger.warning("Снимаю группировку со всех строк")

    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    # Проходим по каждой строке и сбрасываем уровень группировки
    for row in range(1, ws.max_row + 1):
        ws.row_dimensions[row].outline_level = 0
        ws.row_dimensions[row].collapsed = False

    # Также можно сбросить общий уровень сворачивания листа
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.sheet_properties.outlinePr.summaryRight = False

    wb.save(output_path)
    logger.info("Группировка со всех строк успешно снята.")
   
def removes_merging_of_cells():
    """Снимает объединение ячеек со всех строк"""            
    
    logger.warning("Снимает объединение ячеек со всех строк")
    
    column_letter = "G"
    
    wb = openpyxl.load_workbook(output_path)
    ws = wb.active

    # Получаем все объединённые диапазоны на листе
    merged_ranges = list(ws.merged_cells.ranges)

    for merged_range in merged_ranges:
        # Получаем границы диапазона
        min_col, min_row, max_col, max_row = range_boundaries(str(merged_range))

        # Преобразуем букву столбца в номер
        column_num = openpyxl.utils.column_index_from_string(column_letter)

        # Проверяем, пересекается ли объединённый диапазон с нужным столбцом
        if min_col <= column_num <= max_col:
            # Разъединяем ячейки
            ws.unmerge_cells(range_string=str(merged_range))
            logger.info(f"Разъединён диапазон: {merged_range}")

    wb.save(output_path)
    logger.info(f"Объединённые ячейки в столбце {column_letter} разъединены. Файл сохранён как {output_path}")



def puts_vacancies_in_the_staffing_table():
    """Ставит обозначение о вакансии в штатной расстановке"""

    logger.warning("Ставит обозначение о вакансии в штатной расстановке")
    
    column_letter = "A"
    
    wb = openpyxl.load_workbook(output_path)
    ws = wb.active
    
    for cell in ws[column_letter]:
        if cell.value == ",":
            logger.info(f"Вакансия в столбце {column_letter}, строка {cell.row} найдена")
            
            # Если это объединённая ячейка — разъединяем
            if isinstance(cell, MergedCell):
                for merged_range in ws.merged_cells.ranges:
                    if (merged_range.min_row <= cell.row <= merged_range.max_row and
                        merged_range.min_col <= cell.column <= merged_range.max_col):
                        ws.unmerge_cells(str(merged_range))
                        logger.debug(f"  Разъединил: {merged_range}")
                        break
            
            # Заменяем запятую на "ВАКАНСИЯ" в ячейке A
            ws.cell(row=cell.row, column=1, value="ВАКАНСИЯ").font = Font(color="FF0000")
    
    wb.save(output_path)
    logger.info("Изменения сохранены в файл")

if __name__ == '__main__':
    removes_grouping() # снимает группировку со всех строк
    removes_merging_of_cells() # снимает объединение ячеек со всех строк
    puts_vacancies_in_the_staffing_table() # считываем все строки 