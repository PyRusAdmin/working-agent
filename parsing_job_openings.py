from bs4 import BeautifulSoup
import json

# ==========================================
# Читаем HTML файл
# ==========================================
with open("employees.html", "r", encoding="utf-8") as f:
    html = f.read()

# ==========================================
# Создаем BeautifulSoup
# ==========================================
soup = BeautifulSoup(html, "html.parser")

# ==========================================
# Список сотрудников
# ==========================================
employees = []

# Текущее подразделение
current_department = None

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

    # Подразделение имеет:
    # CLASS="R12C0"
    # COLSPAN=8
    if ("R12C0" in td_class and first_td.get("colspan") == "8"):
        current_department = first_td.get_text(" ", strip=True)

        print(f"\n📂 Подразделение: {current_department}")

        continue

    # ==========================================
    # ИЩЕМ СОТРУДНИКОВ
    # ==========================================

    # У сотрудника минимум 8 колонок
    if len(values) >= 8:

        tab_number = values[1]

        # Проверяем что это сотрудник
        if tab_number.isdigit():
            employee = {
                "Подразделение": current_department,
                "ФИО": values[0],
                "Табельный номер": values[1],
                "Должность": values[2],
                "Оклад": values[3],
                "Количество ставок": values[4],
                "Дата приема": values[5],
                "Испытательный срок": values[6],
                "Состояние": values[7],
            }

            employees.append(employee)

            print(f"   └── {employee['ФИО']}")

# ==========================================
# Вывод результата
# ==========================================

print("\n" + "=" * 120)

for employee in employees:
    print(
        f"{employee['Подразделение']} | "
        f"{employee['ФИО']} | "
        f"{employee['Должность']} | "
        f"{employee['Состояние']}"
    )

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

print("\n✅ JSON сохранен: employees.json")


# ==========================================
# ПОДСЧЕТ РЕЗУЛЬТАТОВ (ГРУППИРОВКА ПО ПРОФЕССИЯМ)
# ==========================================
from collections import defaultdict

print("\n" + "=" * 80)
print("📊 ДЕТАЛЬНАЯ СТАТИСТИКА ПО ПОДРАЗДЕЛЕНИЯМ И ДОЛЖНОСТЯМ")
print("=" * 80)

# Строим дерево: Подразделение -> Должность -> Состояние -> Количество
stat_tree = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))

for emp in employees:
    dept = emp["Подразделение"]
    position = emp["Должность"]
    status = emp["Состояние"]

    stat_tree[dept][position][status] += 1

# Выводим сгруппированные данные
for dept, positions in stat_tree.items():
    print(f"\n📂 {dept}")

    for position, statuses in positions.items():
        # Считаем общее количество сотрудников на этой должности в отделе
        total_pos_count = sum(statuses.values())

        # Формируем красивую строку со статусами
        status_details = []
        for status, count in statuses.items():
            status_details.append(f"{count} {status.lower()}")

        status_str = ", ".join(status_details)

        # Склоняем слово "должность/должности" в зависимости от количества
        if total_pos_count == 1:
            pos_word = "должность"
        elif 2 <= total_pos_count <= 4:
            pos_word = "должности"
        else:
            pos_word = "должностей"

        print(f"   └── 🛠️ {position}: {total_pos_count} {pos_word} (из них: {status_str})")

print("\n" + "=" * 80 + "\n")
