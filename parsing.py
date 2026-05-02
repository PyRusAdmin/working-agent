from openpyxl import load_workbook

input_file = "input/История изменений оплаты труда.xlsx"
target_file = "input/28.04.2026 (2).xlsx"
output_file = "output/28.04.2026 (2)_updated.xlsx"


def parsing_file(input_file):
    wb = load_workbook(input_file, data_only=True, read_only=False)
    ws = wb["Лист_1"]

    groups = []
    current_group = None

    for row_num in range(8, ws.max_row + 1):
        level = ws.row_dimensions[row_num].outlineLevel
        value_a = ws.cell(row_num, 1).value

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

    return groups


def parsing_start_row_8(input_file, start_row, end_row):
    wb = load_workbook(input_file, data_only=True, read_only=False)
    ws = wb["Лист_1"]

    data_file = []

    for row_num in range(start_row, end_row + 1):
        value_a = ws.cell(row_num, 1).value
        value_b = ws.cell(row_num, 16).value
        value_c = ws.cell(row_num, 19).value

        if value_a:
            data_file.append([value_a, value_b, value_c])

    return data_file


def normalize_name(value):
    if value is None:
        return ""

    name = str(value).split(",", maxsplit=1)[0]
    return " ".join(name.split()).lower()


def update_target_file(data_file, target_file, output_file=None):
    wb = load_workbook(target_file)
    ws = wb.active

    data_by_name = {}
    for full_name, value_l, value_m in data_file:
        normalized_name = normalize_name(full_name)
        if normalized_name:
            data_by_name[normalized_name] = [value_l, value_m]

    updated_rows = []

    for row_num in range(1, ws.max_row + 1):
        full_name = ws.cell(row_num, 1).value
        normalized_name = normalize_name(full_name)

        if normalized_name in data_by_name:
            value_l, value_m = data_by_name[normalized_name]
            ws.cell(row_num, 12).value = value_l
            ws.cell(row_num, 13).value = value_m
            updated_rows.append(row_num)

    save_path = output_file or target_file
    wb.save(save_path)

    return updated_rows


if __name__ == "__main__":
    # Определяем группы для парсинга в файле
    groups = parsing_file(input_file)

    for group in groups:
        print(group["name"], group["start_row"], len(group["rows"]))

    data_file = parsing_start_row_8(input_file=input_file, start_row=1074, end_row=2060)
    updated_rows = update_target_file(
        data_file=data_file,
        target_file=target_file,
        output_file=output_file,
    )

    print(f"Обновлено строк: {len(updated_rows)}")
    print(f"Первые обновленные строки: {updated_rows[:20]}")
    print(f"Файл сохранен: {output_file}")
