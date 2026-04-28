"""
build_staffing_report.py
========================
Объединяет штатное расписание (ШР) и зарплатную ведомость в один Excel-файл.
Показывает: занятые должности (ФИО + таб.№), вакансии, сводку и несопоставленных сотрудников.

Использование:
    python build_staffing_report.py --sr ШР.xlsx --payroll Ведомость.xlsx --out Результат.xlsx

Зависимости:
    pip install pandas openpyxl
"""

import argparse
import math
import sys
from pathlib import Path

import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from loguru import logger

# Папка для хранение логов
logger.add("log/log.log")

# ─────────────────────────────────────────────
#  Константы — цвета (ARGB без #)
# ─────────────────────────────────────────────
CLR = {
    "header": "FF1F3864",  # тёмно-синий — заголовки таблиц
    "dept": "FF2E75B6",  # синий — строка отдела
    "subdept": "FFD6E4F0",  # светло-синий — строка подотдела
    "occ": "FFE2EFDA",  # светло-зелёный — занято
    "vac": "FFFCE4D6",  # светло-красный — вакансия
    "white": "FFFFFFFF",
    "yellow": "FFFFF2CC",  # жёлтый — непостоянный период
    "grey": "FFF2F2F2",
}


# ─────────────────────────────────────────────
#  Вспомогательные функции стилей
# ─────────────────────────────────────────────


def _fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", start_color=hex_color, end_color=hex_color)


def _font(
    bold: bool = False, sz: int = 10, color: str = "FF000000", white: bool = False
) -> Font:
    return Font(
        name="Arial",
        bold=bold,
        size=sz,
        color="FFFFFFFF" if white else color,
    )


def _center() -> Alignment:
    return Alignment(horizontal="center", vertical="center", wrap_text=True)


def _left() -> Alignment:
    return Alignment(horizontal="left", vertical="center", wrap_text=True)


def _border() -> Border:
    side = Side(style="thin", color="FFBFBFBF")
    return Border(left=side, right=side, top=side, bottom=side)


def _style(cell, fill_clr, font=None, align=None):
    """Быстро применяет fill + font + alignment + border к ячейке."""
    cell.fill = _fill(fill_clr)
    if font:
        cell.font = font
    if align:
        cell.alignment = align
    cell.border = _border()


# ─────────────────────────────────────────────
#  Чистка строк — убираем неразрывные пробелы
# ─────────────────────────────────────────────


def _clean(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).replace("\xa0", " ").replace("\u00a0", " ").strip()


def _safe_pay(salary, rate) -> int:
    """Возвращает оклад: приоритет salary, иначе rate*162. Если ничего — 0."""
    for v in (salary, rate):
        s = _clean(v)
        if s and s not in ("nan", "0"):
            try:
                f = float(s)
                if f > 0:
                    return int(f) if v is salary else int(f * 162)
            except ValueError:
                pass
    return 0


# ─────────────────────────────────────────────
#  Парсинг ШР
# ─────────────────────────────────────────────


def load_staffing_schedule(path: str) -> pd.DataFrame:
    """
    Читает файл штатного расписания.
    Возвращает DataFrame с колонками:
        Отдел, Подотдел, Подразделение, Должность_ШР, Оклад, ШтатЕд, Период
    """
    raw = pd.read_excel(path, sheet_name=0, header=6)
    # Columns: 0=Подразд1, 1-3=unnamed, 4=Подразд2, 5=Категория, 6=ШтатЕд, 7=unnamed, 8=Оклад
    raw.columns = [
        "Подразд1",
        "x1",
        "x2",
        "x3",
        "Подразд2",
        "Категория",
        "ШтатЕд",
        "x7",
        "Оклад",
    ]

    # Filter rows with valid ШтатЕд
    valid = raw[raw["ШтатЕд"].notna() & (raw["ШтатЕд"] > 0)].copy()

    rows = []
    for _, row in valid.iterrows():
        job = _clean(row["Подразд1"])
        units = row["ШтатЕд"]
        pay = _safe_pay(row["Оклад"], 0)

        # Skip empty job titles
        if not job or job == ",":
            continue

        rows.append(
            {
                "Отдел": "",
                "Подотдел": "",
                "Подразделение": "",
                "Должность_ШР": job,
                "Оклад": pay,
                "ШтатЕд": float(units),
                "Период": "постоянно",
            }
        )

    df = pd.DataFrame(rows)
    print(
        f"[ШР] Загружено позиций: {len(df)},  штатных единиц: {df['ШтатЕд'].sum():.2f}"
    )
    return df


# ─────────────────────────────────────────────
#  Парсинг ведомости
# ─────────────────────────────────────────────


def load_payroll(path: str) -> pd.DataFrame:
    """
    Читает зарплатную ведомость.
    Возвращает DataFrame с колонками:
        Категория, Подразд3, Профессия, ТабН, ФИО
    Исключает внутренних совместителей (таб. № оканчивается на 'С').
    """
    raw = pd.read_excel(path, sheet_name=0, header=3)
    raw = raw.iloc[2:].reset_index(drop=True)  # строки с колонками-номерами пропускаем
    raw.columns = [
        "№",
        "Категория",
        "Подразд1",
        "Подразд2",
        "Подразд3",
        "Профессия",
        "ТабН",
        "ФИО",
    ] + [f"c{i}" for i in range(raw.shape[1] - 8)]

    emp = raw[raw["ФИО"].notna() & (raw["ФИО"].astype(str).str.strip() != "nan")].copy()
    emp = emp[~emp["ТабН"].astype(str).str.endswith("С")].copy()
    emp = emp[["Категория", "Подразд3", "Профессия", "ТабН", "ФИО"]].copy()

    for col in ("Профессия", "Подразд3", "ФИО"):
        emp[col] = emp[col].apply(_clean)

    print(f"[Ведомость] Загружено сотрудников: {len(emp)}")
    return emp.reset_index(drop=True)


# ─────────────────────────────────────────────
#  Сопоставление
# ─────────────────────────────────────────────


def match(sr_df: pd.DataFrame, emp: pd.DataFrame):
    """
    Для каждой позиции ШР находит сотрудников из ведомости.
    Возвращает:
        out_df      — основная таблица (строка на каждую штатную единицу)
        unmatched   — сотрудники без позиции в ШР
    """
    output = []
    used_idx = set()

    for _, pos in sr_df.iterrows():
        dept = pos["Подразделение"]
        job = pos["Должность_ШР"]
        units = pos["ШтатЕд"]
        pay = pos["Оклад"]
        period = pos["Период"]
        m_dept = pos["Отдел"]
        subdept = pos["Подотдел"]

        mask = (
            (emp["Подразд3"] == dept)
            & (emp["Профессия"] == job)
            & (~emp.index.isin(used_idx))
        )
        matched = emp[mask]
        slots = math.ceil(units)

        for i in range(slots):
            if i < len(matched):
                e = matched.iloc[i]
                used_idx.add(matched.index[i])
                output.append(
                    {
                        "Отдел": m_dept,
                        "Подотдел": subdept,
                        "Должность": job,
                        "Штатных ед.": units,
                        "Оклад, руб.": pay,
                        "Период": period,
                        "Таб.№": str(e["ТабН"]),
                        "ФИО": e["ФИО"],
                        "Статус": "Занято",
                    }
                )
            else:
                output.append(
                    {
                        "Отдел": m_dept,
                        "Подотдел": subdept,
                        "Должность": job,
                        "Штатных ед.": units,
                        "Оклад, руб.": pay,
                        "Период": period,
                        "Таб.№": "",
                        "ФИО": "ВАКАНСИЯ",
                        "Статус": "Вакансия",
                    }
                )

    out_df = pd.DataFrame(output)
    unmatched = emp[~emp.index.isin(used_idx)].copy()

    occ = (out_df["Статус"] == "Занято").sum()
    vac = (out_df["Статус"] == "Вакансия").sum()
    print(
        f"[Сопоставление] Занято: {occ},  Вакансий: {vac},  Без позиции в ШР: {len(unmatched)}"
    )
    return out_df, unmatched


# ─────────────────────────────────────────────
#  Запись Excel — лист 1: основная таблица
# ─────────────────────────────────────────────


def _write_main_sheet(ws, out_df: pd.DataFrame):
    col_widths = [42, 46, 13, 14, 22, 14, 38, 13]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    headers = [
        "Отдел / Подразделение",
        "Должность / профессия",
        "Штат.\nед.",
        "Оклад,\nруб.",
        "Период действия",
        "Таб. №",
        "ФИО",
        "Статус",
    ]
    ws.row_dimensions[1].height = 38
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=ci, value=h)
        _style(c, CLR["header"], _font(bold=True, white=True), _center())
    ws.freeze_panes = "A2"

    rn = 2
    prev_dept = prev_subdept = None

    for _, r in out_df.iterrows():
        dept = r["Отдел"]
        subdept = r["Подотдел"]
        job = r["Должность"]
        units = r["Штатных ед."]
        pay = r["Оклад, руб."] or ""
        period = r["Период"]
        tabn = r["Таб.№"]
        fio = r["ФИО"]
        status = r["Статус"]
        occ = status == "Занято"

        # ── Строка отдела ──────────────────────────────────────
        if dept != prev_dept:
            ws.row_dimensions[rn].height = 22
            c = ws.cell(row=rn, column=1, value=dept)
            _style(c, CLR["dept"], _font(bold=True, sz=11, white=True), _left())
            ws.merge_cells(start_row=rn, start_column=1, end_row=rn, end_column=8)
            rn += 1
            prev_dept = dept
            prev_subdept = None

        # ── Строка подотдела ───────────────────────────────────
        if subdept and subdept != prev_subdept:
            ws.row_dimensions[rn].height = 18
            c = ws.cell(row=rn, column=1, value=f"  {subdept}")
            _style(c, CLR["subdept"], _font(bold=True, color="FF1F3864"), _left())
            ws.merge_cells(start_row=rn, start_column=1, end_row=rn, end_column=8)
            rn += 1
            prev_subdept = subdept

        # ── Строка данных ──────────────────────────────────────
        ws.row_dimensions[rn].height = 15
        rf = CLR["occ"] if occ else CLR["vac"]
        label = subdept if subdept else dept

        specs = [
            # (col, value, fill_clr, font, align)
            (1, label, CLR["white"], _font(sz=9, color="FF595959"), _left()),
            (2, job, rf, _font(sz=10), _left()),
            (3, units, rf, _font(sz=10), _center()),
            (4, pay, rf, _font(sz=10), _center()),
            (
                5,
                period,
                CLR["yellow"] if period != "постоянно" else rf,
                _font(sz=9),
                _center(),
            ),
            (6, tabn, rf, _font(sz=10), _center()),
            (
                7,
                fio,
                rf,
                _font(bold=not occ, sz=10, color="FF376923" if occ else "FFBF1F1F"),
                _left(),
            ),
            (
                8,
                status,
                rf,
                _font(bold=True, sz=10, color="FF376923" if occ else "FFBF1F1F"),
                _center(),
            ),
        ]
        for ci, val, fc, fn, al in specs:
            c = ws.cell(row=rn, column=ci, value=val)
            _style(c, fc, fn, al)
        rn += 1

    ws.auto_filter.ref = f"A1:{get_column_letter(8)}{rn - 1}"
    return rn - 1


# ─────────────────────────────────────────────
#  Запись Excel — лист 2: сводка
# ─────────────────────────────────────────────


def _write_summary_sheet(ws, out_df: pd.DataFrame, sr_df: pd.DataFrame):
    for col, w in zip("ABCDE", [52, 12, 12, 12, 14]):
        ws.column_dimensions[col].width = w

    hdrs = ["Подразделение", "Штат. ед.", "Занято", "Вакансий", "% укомпл."]
    ws.row_dimensions[1].height = 30
    for ci, h in enumerate(hdrs, 1):
        c = ws.cell(row=1, column=ci, value=h)
        _style(c, CLR["header"], _font(bold=True, white=True), _center())
    ws.freeze_panes = "A2"

    dept_units = sr_df.groupby("Отдел")["ШтатЕд"].sum().to_dict()
    summary = (
        out_df.groupby("Отдел")
        .agg(
            Занято=("Статус", lambda x: (x == "Занято").sum()),
            Вакансий=("Статус", lambda x: (x == "Вакансия").sum()),
        )
        .reset_index()
        .sort_values("Отдел")
    )

    rn = 2
    for _, row in summary.iterrows():
        dept = row["Отдел"]
        штат = dept_units.get(dept, 0)
        заня = row["Занято"]
        вак = row["Вакансий"]
        pct = round(заня / штат * 100, 1) if штат > 0 else 0

        rf = CLR["occ"] if pct >= 80 else CLR["yellow"] if pct >= 50 else CLR["vac"]
        ws.row_dimensions[rn].height = 16
        for ci, val in enumerate([dept, штат, заня, вак, f"{pct} %"], 1):
            c = ws.cell(row=rn, column=ci, value=val)
            _style(
                c,
                CLR["white"] if ci == 1 else rf,
                _font(sz=10),
                _left() if ci == 1 else _center(),
            )
        rn += 1

    # Итого
    total_u = sum(dept_units.values())
    total_z = (out_df["Статус"] == "Занято").sum()
    total_v = (out_df["Статус"] == "Вакансия").sum()
    total_p = round(total_z / total_u * 100, 1) if total_u else 0
    for ci, val in enumerate(
        ["ИТОГО", round(total_u, 2), total_z, total_v, f"{total_p} %"], 1
    ):
        c = ws.cell(row=rn, column=ci, value=val)
        _style(
            c,
            CLR["dept"],
            _font(bold=True, sz=11, white=True),
            _left() if ci == 1 else _center(),
        )

    ws.auto_filter.ref = f"A1:E{rn - 1}"


# ─────────────────────────────────────────────
#  Запись Excel — лист 3: несопоставленные
# ─────────────────────────────────────────────


def _write_unmatched_sheet(ws, unmatched: pd.DataFrame):
    for col, w in zip("ABCD", [45, 55, 14, 38]):
        ws.column_dimensions[col].width = w

    # Информационная строка
    ws.merge_cells("A1:D1")
    c = ws.cell(
        row=1,
        column=1,
        value=(
            "Сотрудники из ведомости, чья должность или подразделение "
            "не найдены в штатном расписании — требуют ручной проверки"
        ),
    )
    _style(
        c,
        CLR["yellow"],
        Font(name="Arial", italic=True, size=10, color="FF7F6000"),
        _left(),
    )
    ws.row_dimensions[1].height = 18

    hdrs = ["Подразделение (ведомость)", "Должность (ведомость)", "Таб. №", "ФИО"]
    ws.row_dimensions[2].height = 28
    for ci, h in enumerate(hdrs, 1):
        c = ws.cell(row=2, column=ci, value=h)
        _style(c, CLR["header"], _font(bold=True, white=True), _center())

    rn = 3
    for _, e in unmatched.sort_values(["Подразд3", "Профессия"]).iterrows():
        ws.row_dimensions[rn].height = 16
        for ci, val in enumerate(
            [e["Подразд3"], e["Профессия"], str(e["ТабН"]), e["ФИО"]], 1
        ):
            c = ws.cell(row=rn, column=ci, value=val)
            _style(
                c, CLR["vac"], _font(sz=10), _left() if ci in (1, 2, 4) else _center()
            )
        rn += 1

    ws.auto_filter.ref = f"A2:D{rn - 1}"


# ─────────────────────────────────────────────
#  Главная функция
# ─────────────────────────────────────────────


def build_report(sr_path: str, payroll_path: str, out_path: str):
    print("=" * 55)
    print("  Загрузка данных")
    print("=" * 55)
    sr_df = load_staffing_schedule(sr_path)
    emp_df = load_payroll(payroll_path)

    print("\n" + "=" * 55)
    print("  Сопоставление")
    print("=" * 55)
    out_df, unmatched = match(sr_df, emp_df)

    print("\n" + "=" * 55)
    print("  Формирование Excel")
    print("=" * 55)

    wb = openpyxl.Workbook()

    ws1 = wb.active
    ws1.title = "Штатное расписание"
    total_rows = _write_main_sheet(ws1, out_df)
    print(f"  Лист 1 «Штатное расписание»: {total_rows} строк")

    ws2 = wb.create_sheet("Сводка")
    _write_summary_sheet(ws2, out_df, sr_df)
    print(f"  Лист 2 «Сводка»: {out_df['Отдел'].nunique()} подразделений")

    ws3 = wb.create_sheet("Не сопоставлены")
    _write_unmatched_sheet(ws3, unmatched)
    print(f"  Лист 3 «Не сопоставлены»: {len(unmatched)} сотрудников")

    wb.save(out_path)
    print(f"\n[OK] Файл сохранён: {out_path}")
    print("=" * 55)


# ─────────────────────────────────────────────
#  CLI
# ─────────────────────────────────────────────


def main():
    parser = argparse.ArgumentParser(
        description="Объединяет штатное расписание и ведомость в один Excel-отчёт."
    )
    parser.add_argument(
        "--sr",
        default="input/28.04.2026.xlsx",
        help="Путь к файлу штатного расписания (.xlsx)",
    )
    parser.add_argument(
        "--payroll",
        default="input/Март_2026.xlsx",
        help="Путь к файлу зарплатной ведомости (.xlsx)",
    )
    parser.add_argument(
        "--out",
        default="ШР_с_сотрудниками.xlsx",
        help="Путь к выходному файлу (.xlsx)",
    )
    args = parser.parse_args()

    for p in (args.sr, args.payroll):
        if not Path(p).exists():
            logger.error(f"[ОШИБКА] Файл не найден: {p}", file=sys.stderr)
            sys.exit(1)

    build_report(args.sr, args.payroll, args.out)


if __name__ == "__main__":
    main()
