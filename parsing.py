from openpyxl import Workbook, load_workbook

input_file = "input/История изменений оплаты труда.xlsx"
target_file = "input/28.04.2026 (2).xlsx"
output_file = "output/28.04.2026 (2)_updated.xlsx"
missed_intensity_report_file = "output/missed_intensity_report.xlsx"
missed_disinfection_report_file = "output/missed_disinfection_report.xlsx"


def parsing_file(input_file):
    """
    Парсинг файла на наличие строк с группировкой
    """
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
    """
    Парсинг файла в определенном диапазоне строк
    :param input_file:
    :param start_row:
    :param end_row:
    :return:
    """
    wb = load_workbook(input_file, data_only=True, read_only=False)
    ws = wb["Лист_1"]

    data_file = []

    for row_num in range(start_row, end_row + 1):
        level = ws.row_dimensions[row_num].outlineLevel
        value_a = ws.cell(row_num, 1).value
        value_b = ws.cell(row_num, 16).value
        value_c = ws.cell(row_num, 19).value

        if level == 2 and value_a:
            data_file.append([row_num, value_a, value_b, value_c])

    return data_file


def normalize_name(value):
    if value is None:
        return ""

    name = str(value).split(",", maxsplit=1)[0]
    return " ".join(name.split()).lower()


def update_target_file(column_1, column_2, data_file, target_file, output_file=None):
    """
    За использование в работе дезинфицирующих средств, а также работникам, занятым уборкой туалетов
    :param data_file:
    :param target_file:
    :param output_file:
    :param column_1: Строка для записи процента
    :param column_2: Строка для записи суммы
    """
    wb = load_workbook(target_file)
    ws = wb.active

    data_by_name = aggregate_data_by_name(data_file)

    updated_rows = []
    updated_names = set()

    for row_num in range(1, ws.max_row + 1):
        full_name = ws.cell(row_num, 1).value
        normalized_name = normalize_name(full_name)

        if normalized_name in data_by_name and normalized_name not in updated_names:
            value_r, value_s = data_by_name[normalized_name]
            ws.cell(row_num, column_1).value = value_r
            ws.cell(row_num, column_2).value = value_s
            updated_rows.append(row_num)
            updated_names.add(normalized_name)

    save_path = output_file or target_file
    wb.save(save_path)


def aggregate_data_by_name(data_file):
    data_by_name = {}

    for _, full_name, value_l, value_m in data_file:
        normalized_name = normalize_name(full_name)
        if not normalized_name:
            continue

        if normalized_name not in data_by_name:
            data_by_name[normalized_name] = [0, 0]

        data_by_name[normalized_name][0] += to_number(value_l)
        data_by_name[normalized_name][1] += to_number(value_m)

    return data_by_name


def to_number(value):
    if isinstance(value, int | float):
        return value

    if value is None:
        return 0

    try:
        return float(str(value).replace(" ", "").replace(",", "."))
    except ValueError:
        return 0


def make_match_report(data_file, target_file, report_file):
    wb = load_workbook(target_file, read_only=True, data_only=True)
    ws = wb.active

    target_names = {}
    for row_num, row in enumerate(ws.iter_rows(values_only=True), start=1):
        normalized_name = normalize_name(row[0])
        if normalized_name:
            target_names.setdefault(normalized_name, []).append(row_num)

    aggregated_data = aggregate_data_by_name(data_file)
    source_names = {}
    for source_row, full_name, value_l, value_m in data_file:
        normalized_name = normalize_name(full_name)
        if normalized_name:
            source_names.setdefault(normalized_name, []).append(
                [source_row, full_name, value_l, value_m]
            )

    missed = []
    matched_total = 0
    missed_total = 0
    source_total = 0
    aggregated_matched_total = 0
    aggregated_missed_total = 0

    for source_row, full_name, value_l, value_m in data_file:
        amount = to_number(value_m)
        source_total += amount
        normalized_name = normalize_name(full_name)

        if normalized_name in target_names:
            matched_total += amount
        else:
            missed_total += amount
            missed.append([source_row, full_name, value_l, value_m])

    for normalized_name, (_, amount) in aggregated_data.items():
        if normalized_name in target_names:
            aggregated_matched_total += amount
        else:
            aggregated_missed_total += amount

    duplicate_source_names = []
    for rows in source_names.values():
        if len(rows) > 1:
            total_amount = sum(to_number(row[3]) for row in rows)
            last_amount = to_number(rows[-1][3])
            source_rows = ", ".join(str(row[0]) for row in rows)
            duplicate_source_names.append(
                [
                    rows[0][1],
                    len(rows),
                    total_amount,
                    last_amount,
                    total_amount - last_amount,
                    source_rows,
                ]
            )
    duplicate_target_names = [
        [name, len(rows), ", ".join(str(row) for row in rows[:20])]
        for name, rows in target_names.items()
        if len(rows) > 1
    ]

    report_wb = Workbook()
    ws_missed = report_wb.active
    ws_missed.title = "Не найдены"
    ws_missed.append(["Строка источника", "ФИО", "Процент", "Сумма"])
    for row in missed:
        ws_missed.append(row)

    ws_source_duplicates = report_wb.create_sheet("Дубли в источнике")
    ws_source_duplicates.append(
        [
            "ФИО",
            "Количество",
            "Сумма всех строк",
            "Последняя сумма",
            "Потеря при перезаписи",
            "Строки источника",
        ]
    )
    for row in duplicate_source_names:
        ws_source_duplicates.append(row)

    ws_target_duplicates = report_wb.create_sheet("Дубли в цели")
    ws_target_duplicates.append(["ФИО нормализованное", "Количество", "Первые строки"])
    for row in duplicate_target_names:
        ws_target_duplicates.append(row)

    ws_summary = report_wb.create_sheet("Итог")
    ws_summary.append(["Показатель", "Значение"])
    ws_summary.append(["Строк в источнике", len(data_file)])
    ws_summary.append(["Уникальных ФИО в источнике", len(source_names)])
    ws_summary.append(["Уникальных ФИО в цели", len(target_names)])
    ws_summary.append(["Сумма источника", source_total])
    ws_summary.append(["Сумма найденных в цели", matched_total])
    ws_summary.append(["Сумма не найденных в цели", missed_total])
    ws_summary.append(["Сумма найденных после объединения дублей", aggregated_matched_total])
    ws_summary.append(["Сумма не найденных после объединения дублей", aggregated_missed_total])
    ws_summary.append(["Не найдено строк источника", len(missed)])
    ws_summary.append(["ФИО с дублями в источнике", len(duplicate_source_names)])
    ws_summary.append(["ФИО с дублями в цели", len(duplicate_target_names)])

    report_wb.save(report_file)

    return {
        "source_total": source_total,
        "matched_total": matched_total,
        "missed_total": missed_total,
        "aggregated_matched_total": aggregated_matched_total,
        "aggregated_missed_total": aggregated_missed_total,
        "missed_count": len(missed),
        "duplicate_source_count": len(duplicate_source_names),
        "duplicate_target_count": len(duplicate_target_names),
    }


if __name__ == "__main__":
    # Определяем группы для парсинга в файле
    groups = parsing_file(input_file)

    for group in groups:
        print(group["name"], group["start_row"], len(group["rows"]))

    """За интенсивность труда работников"""

    # Парсим файл с данными (Имя, Процент, Сумма)
    data_file_1 = parsing_start_row_8(input_file=input_file, start_row=1074, end_row=2059)
    report_1 = make_match_report(
        data_file=data_file_1,
        target_file=target_file,
        report_file=missed_intensity_report_file,
    )
    update_target_file(
        column_1=12,
        column_2=13,
        data_file=data_file_1,
        target_file=target_file,
        output_file=output_file,
    )

    """За использование в работе дезинфицирующих средств, а также работникам, занятым уборкой туалетов"""

    # Парсим файл с данными (Имя, Процент, Сумма)
    data_file = parsing_start_row_8(input_file=input_file, start_row=2075, end_row=2181)
    report_2 = make_match_report(
        data_file=data_file,
        target_file=target_file,
        report_file=missed_disinfection_report_file,
    )
    update_target_file(
        column_1=18,
        column_2=19,
        data_file=data_file,
        target_file=output_file,
        output_file=output_file,
    )
