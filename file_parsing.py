import openpyxl



def removes_grouping():
    """Снимает группировку со всех строк"""             
    # Загружаем рабочую книгу и выбираем активный лист
    wb = openpyxl.load_workbook('input/Штатная_расстановка.xlsx')
    ws = wb.active

    # Снимаем группировку со всех строк (устанавливаем уровень 0)
    ws.row_dimensions.group(1, ws.max_row + 1, outline_level=0)

    # Сохраняем изменения
    wb.save('output/your_file_modified.xlsx')
    
    

if __name__ == '__main__':
    removes_grouping() # снимает группировку со всех строк