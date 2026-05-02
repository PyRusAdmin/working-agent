from openpyxl import load_workbook

input_file = "input/История изменений оплаты труда.xlsx"


def parsing_file(input_file):
    wb = load_workbook(input_file, data_only=True, read_only=False)
    ws = wb["Лист_1"]

    groups = []
    current_group = None

    for row_num in range(8, ws.max_row + 1):
        level = ws.row_dimensions[row_num].outlineLevel
        # print(level)
        value_a = ws.cell(row_num, 1).value
        print(value_a)

        # Строка группы: например "Оплата по окладу по дням"
        if level == 1 and value_a:
            current_group = {
                "name": value_a,
                "start_row": row_num,
                "rows": [],
            }
            groups.append(current_group)
            continue

        if current_group is not None:
            current_group["rows"].append(row_num)

    print(groups)
    return groups


def parsing_start_row_8(input_file, start_row, end_row):
    wb = load_workbook(input_file, data_only=True, read_only=False)
    ws = wb["Лист_1"]

    for row_num in range(start_row, end_row):
        value_a = ws.cell(row_num, 1).value
        print(value_a)
        value_b = ws.cell(row_num, 16).value
        print(value_b)


if __name__ == "__main__":
    # Определяем группы для парсинга в файле
    groups = parsing_file(input_file)

    for group in groups:
        print(group["name"], group["start_row"], len(group["rows"]))

    parsing_start_row_8(input_file=input_file, start_row=1074, end_row=2060)
