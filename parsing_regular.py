from bs4 import BeautifulSoup
import json

"""
Парсинг штатной расстановки
"""

# ==========================================
# Читаем HTML файл (штатной расстановки)
# ==========================================
with open("input/Штатная_расстановка.html", "r", encoding="utf-8-sig") as f:
    html = f.read()

# ==========================================
# Создаем BeautifulSoup
# ==========================================
soup = BeautifulSoup(html, "html.parser")

# ==========================================
# Список сотрудников и должностей
# ==========================================
employees = []
jobs_info = {}

# Текущее подразделение
current_department = None

current_job_title = None

# ==========================================
# Проходим по всем строкам
# ==========================================
for row in soup.find_all("tr"):

    # Получаем все ячейки
    cells = row.find_all("td")

    # Пропускаем пустые строки
    if not cells:
        continue

    # Получаем текст ячеек
    values = [
        cell.get_text(" ", strip=True)
        for cell in cells
    ]

    # ==========================================
    # ИЩЕМ ПОДРАЗДЕЛЕНИЕ
    # ==========================================

    first_td = cells[0]

    td_class = first_td.get("class", [])

    if "R9C0" in td_class and first_td.get("colspan") == "2":
        current_department = first_td.get_text(" ", strip=True)

        print(f"\n📂 Подразделение: {current_department}")

        continue

    # ==========================================
    # ИЩЕМ ДОЛЖНОСТЬ
    # ==========================================
    if "R10C0" in td_class and first_td.get("colspan") == "2":
        job_title_full = first_td.get_text(" ", strip=True)
        # Если должность содержит подразделение, например "Аппаратчик /Отдел/", отрежем его
        if " /" in job_title_full:
            current_job_title = job_title_full.split(" /")[0]
        else:
            current_job_title = job_title_full

        planned = values[1] if len(values) > 1 else "0"
        diff = values[2] if len(values) > 2 else "0"
        actual = values[3] if len(values) > 3 else "0"

        if current_department not in jobs_info:
            jobs_info[current_department] = {}
        jobs_info[current_department][current_job_title] = {
            "planned": planned,
            "diff": diff,
            "actual": actual
        }

        continue

    # ==========================================
    # ИЩЕМ СОТРУДНИКОВ
    # ==========================================

    if "R11C0" in td_class:
        fio_and_status = values[0]
        fio_parts = fio_and_status.split(",", 1)
        fio = fio_parts[0].strip()
        status = fio_parts[1].strip() if len(fio_parts) > 1 else ""

        if not fio:
            status = "Вакансия"

        employee = {
            "Подразделение": current_department,
            "ФИО": fio,
            "Табельный номер": "",
            "Должность": current_job_title,
            "Оклад": values[1] if len(values) > 1 else "",
            "Количество ставок": values[4] if len(values) > 4 else "",
            "Дата приема": "",
            "Испытательный срок": "",
            "Состояние": status,
        }
        employees.append(employee)

        print(f"   └── {employee['ФИО']}", flush=True)
        # Отладочный вывод
        print(f"DEBUG: Добавлен сотрудник: {employee}", flush=True)

# ==========================================
# Вывод результата
# ==========================================

# Вывод списка сотрудников отключен из-за проблем с кодировкой
# print("\n" + "=" * 120)
# 
# for employee in employees:
#     print(
#         f"{employee['Подразделение']} | "
#         f"{employee['ФИО']} | "
#         f"{employee['Должность']}"
#     )

# ==========================================
# Сохраняем JSON
# ==========================================

with open(
        "employees.json",
        "w",
        encoding="utf-8"
) as f:
    json.dump(
        employees,
        f,
        ensure_ascii=False,
        indent=4
    )

print("\nJSON сохранен: employees.json", flush=True)

# ==========================================
# ПОДСЧЕТ РЕЗУЛЬТАТОВ (ГРУППИРОВКА ПО ПРОФЕССИЯМ)
# ==========================================
from collections import defaultdict

# Вывод статистики отключен из-за проблем с кодировкой
# print("\n" + "=" * 80)
# print("ДЕТАЛЬНАЯ СТАТИСТИКА ПО ПОДРАЗДЕЛЕНИЯМ И ДОЛЖНОСТЯМ")
# print("=" * 80)

# Строим дерево: Подразделение -> Должность -> Состояние -> Количество ставок
stat_tree = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))


def parse_stakes(val):
    if not val:
        return 0.0
    try:
        return float(val.replace(",", ".").replace(" ", ""))
    except ValueError:
        return 0.0


for emp in employees:
    dept = emp["Подразделение"]
    position = emp["Должность"]
    status = emp["Состояние"]
    stakes = parse_stakes(emp.get("Количество ставок", "1"))

    stat_tree[dept][position][status] += stakes

# Выводим сгруппированные данные
for dept, positions in stat_tree.items():
    print(f"\n📂 {dept}")

    for position, statuses in positions.items():
        # Считаем общее количество ставок на этой должности в отделе
        total_pos_count = sum(statuses.values())

        # Формируем красивую строку со статусами
        status_details = []
        for status, count in statuses.items():
            # Убираем лишние нули после запятой, если число целое
            count_str = f"{count:g}" if count % 1 != 0 else f"{int(count)}"

            display_status = status.lower()
            if display_status == "трудовой договор приостановлен":
                display_status = "трудовой договор приостановлен (мобилизованный)"

            if status.lower() == "вакансия":
                if count_str.endswith("1") and not count_str.endswith("11"):
                    status_name = "вакансия"
                elif count_str[-1] in "234" and not (len(count_str) > 1 and count_str[-2:] in ("12", "13", "14")):
                    status_name = "вакансии"
                else:
                    status_name = "вакансий"
                status_details.append(f"{count_str} {status_name}")
            else:
                status_details.append(f"{count_str} {display_status}")

        status_str = ", ".join(status_details)

        total_pos_str = f"{total_pos_count:g}" if total_pos_count % 1 != 0 else f"{int(total_pos_count)}"

        # Склоняем слово "ставка/ставки" в зависимости от количества
        if total_pos_str.endswith("1") and not total_pos_str.endswith("11"):
            pos_word = "ставка"
        elif total_pos_str[-1] in "234" and not (len(total_pos_str) > 1 and total_pos_str[-2:] in ("12", "13", "14")):
            pos_word = "ставки"
        else:
            pos_word = "ставок"

        # Получаем данные из штатки по этой должности
        job_info = jobs_info.get(dept, {}).get(position, {})
        planned_stakes = parse_stakes(job_info.get("planned", "0"))
        diff_stakes = parse_stakes(job_info.get("diff", "0"))

        overdraft_str = ""
        if diff_stakes < 0:
            overdraft_val = abs(diff_stakes)
            overdraft_val_str = f"{overdraft_val:g}" if overdraft_val % 1 != 0 else f"{int(overdraft_val)}"
            planned_str = f"{planned_stakes:g}" if planned_stakes % 1 != 0 else f"{int(planned_stakes)}"
            overdraft_str = f". Перерасход: {overdraft_val_str} ставка (по штату: {planned_str})"

        print(f"   └── 🛠️ {position}: {total_pos_str} {pos_word} (из них: {status_str}){overdraft_str}")

# ==========================================
# ВЫВОД В ВИДЕ ТАБЛИЦЫ (PRETTYTABLE)
# ==========================================
from prettytable import PrettyTable

print("\n" + "=" * 120)
print("ИТОГОВАЯ ТАБЛИЦА ШТАТНОЙ РАССТАНОВКИ")
print("=" * 120)

output_lines = []
output_lines.append("ИТОГОВАЯ ТАБЛИЦА ШТАТНОЙ РАССТАНОВКИ")
output_lines.append("=" * 120)

for dept, positions in stat_tree.items():
    table = PrettyTable()
    table.field_names = ["Должность", "По штату", "Факт (ставок)", "Вакансии", "Перерасход", "Детализация (факт)"]
    table.align = "l"
    table.max_width["Должность"] = 40
    table.max_width["Детализация (факт)"] = 45
    
    for position, statuses in positions.items():
        total_pos_count = sum(statuses.values())
        
        # Получаем данные из штатки по этой должности
        job_info = jobs_info.get(dept, {}).get(position, {})
        planned_stakes = parse_stakes(job_info.get("planned", "0"))
        diff_stakes = parse_stakes(job_info.get("diff", "0"))
        
        overdraft_val = abs(diff_stakes) if diff_stakes < 0 else 0
        
        # Считаем количество вакансий
        vacancies = 0.0
        for status, count in statuses.items():
            if status.lower() == "вакансия":
                vacancies += count
                
        # Форматирование чисел
        planned_str = f"{planned_stakes:g}" if planned_stakes % 1 != 0 else f"{int(planned_stakes)}"
        total_pos_str = f"{total_pos_count:g}" if total_pos_count % 1 != 0 else f"{int(total_pos_count)}"
        vacancies_str = f"{vacancies:g}" if vacancies % 1 != 0 else f"{int(vacancies)}"
        overdraft_str = f"{overdraft_val:g}" if overdraft_val % 1 != 0 else f"{int(overdraft_val)}"
        
        # Формирование строки статусов
        status_details = []
        for status, count in statuses.items():
            if status.lower() == "вакансия":
                continue  # Пропускаем вакансии, так как для них есть отдельная колонка
                
            count_str = f"{count:g}" if count % 1 != 0 else f"{int(count)}"
            display_status = status.lower()
            if display_status == "трудовой договор приостановлен":
                display_status = "приост. т/д (мобилиз.)"
            status_details.append(f"{count_str} {display_status}")
            
        status_str = ", ".join(status_details)
        if not status_str:
            status_str = "-"
            
        table.add_row([
            position,
            planned_str,
            total_pos_str,
            vacancies_str,
            overdraft_str,
            status_str
        ])
    
    dept_header = f"\n📁 ПОДРАЗДЕЛЕНИЕ: {dept}"
    print(dept_header)
    print(table)
    
    output_lines.append(dept_header)
    output_lines.append(str(table))

# ==========================================
# СОХРАНЕНИЕ ТАБЛИЦЫ В ФАЙЛ
# ==========================================
with open("report_table.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(output_lines) + "\n")

print("\nТаблица сохранена в файл: report_table.txt", flush=True)
