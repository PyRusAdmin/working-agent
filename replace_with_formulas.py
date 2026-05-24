import openpyxl
from openpyxl.utils import get_column_letter

INPUT_FILE = "1. Донецк ШР 01.01.26_план.xlsx"
OUTPUT_FILE = "1. Донецк ШР 01.01.26_план_formulas.xlsx"
SHEET_NAME = "ШР_26"

DATA_START_ROW = 27

wb = openpyxl.load_workbook(INPUT_FILE)
ws = wb[SHEET_NAME]
max_row = ws.max_row


def get_val(row, col):
    return ws.cell(row=row, column=col).value


def is_rukovoditel_row(row):
    d = get_val(row, 4)
    return isinstance(d, str) and "Руководители" in d


def is_rabochie_row(row):
    d = get_val(row, 4)
    return isinstance(d, str) and d.strip() == "Рабочие"


def is_position_row(row):
    d = get_val(row, 4)
    if not isinstance(d, str) or d.strip() == "":
        return False
    if "Руководители" in d:
        return False
    if d.strip() == "Рабочие":
        return False
    return True


def is_subgroup_header(row):
    c = get_val(row, 3)
    d = get_val(row, 4)
    return isinstance(c, str) and c.strip() != "" and (d is None or (isinstance(d, str) and d.strip() == ""))


def is_division_header(row):
    b = get_val(row, 2)
    d = get_val(row, 4)
    return isinstance(b, str) and b.strip() != "" and (d is None or (isinstance(d, str) and d.strip() == ""))


def is_grand_total_row(row):
    a = get_val(row, 1)
    if not isinstance(a, str):
        return False
    a_upper = a.upper()
    return "ИТОГ" in a_upper or "ВСЕГО" in a_upper


def is_formula(val):
    return isinstance(val, str) and val.startswith("=")


def parse_subgroup_structure(start, end_limit):
    """Parse a sequence of sub-groups within a range.
    Returns list of sub-groups and the row where parsing stopped."""
    subgroups = []
    i = start
    while i <= end_limit:
        if is_subgroup_header(i):
            sg_header = i
            sg_ruk = None
            sg_ruk_pos = []
            sg_rab = None
            sg_rab_pos = []
            j = i + 1
            while j <= end_limit:
                if is_rukovoditel_row(j) and sg_ruk is None:
                    sg_ruk = j
                    j += 1
                    continue
                if is_rabochie_row(j) and sg_ruk is not None and sg_rab is None:
                    sg_rab = j
                    j += 1
                    continue
                if sg_ruk is not None and sg_rab is None and is_position_row(j):
                    sg_ruk_pos.append(j)
                    j += 1
                    continue
                if sg_rab is not None and is_position_row(j):
                    sg_rab_pos.append(j)
                    j += 1
                    continue
                break
            subgroups.append({
                "header": sg_header,
                "ruk": sg_ruk,
                "ruk_pos": sg_ruk_pos,
                "rab": sg_rab,
                "rab_pos": sg_rab_pos,
            })
            i = j
        elif is_rukovoditel_row(i):
            # Standalone rukovoditel row (not under a sub-group header)
            sg_ruk = i
            sg_ruk_pos = []
            j = i + 1
            while j <= end_limit:
                if is_rabochie_row(j):
                    sg_rab = j
                    j += 1
                    break
                if is_position_row(j):
                    sg_ruk_pos.append(j)
                    j += 1
                    continue
                break
            sg_rab = None
            sg_rab_pos = []
            if j <= end_limit and is_rabochie_row(j - 1):
                sg_rab = j - 1
                while j <= end_limit:
                    if is_position_row(j):
                        sg_rab_pos.append(j)
                        j += 1
                        continue
                    break
            subgroups.append({
                "header": None,
                "ruk": sg_ruk,
                "ruk_pos": sg_ruk_pos,
                "rab": sg_rab,
                "rab_pos": sg_rab_pos,
            })
            i = j
        else:
            break
    return subgroups, i


def process_subgroup(sg):
    changes = {"k": 0, "l": 0}
    ruk = sg["ruk"]
    ruk_pos = sg["ruk_pos"]
    rab = sg["rab"]
    rab_pos = sg["rab_pos"]
    header = sg["header"]

    if ruk is not None and ruk_pos:
        ps, pe = ruk_pos[0], ruk_pos[-1]
        ws.cell(row=ruk, column=11).value = f"=SUM(K{ps}:K{pe})"
        changes["k"] += 1
        ws.cell(row=ruk, column=12).value = f"=SUM(L{ps}:L{pe})"
        changes["l"] += 1

    if rab is not None and rab_pos:
        ps, pe = rab_pos[0], rab_pos[-1]
        ws.cell(row=rab, column=11).value = f"=SUM(K{ps}:K{pe})"
        changes["k"] += 1
        ws.cell(row=rab, column=12).value = f"=SUM(L{ps}:L{pe})"
        changes["l"] += 1

    if header is not None and ruk is not None:
        if rab is not None:
            if not is_formula(get_val(header, 11)):
                ws.cell(row=header, column=11).value = f"=K{ruk}+K{rab}"
                changes["k"] += 1
            if not is_formula(get_val(header, 12)):
                ws.cell(row=header, column=12).value = f"=L{ruk}+L{rab}"
                changes["l"] += 1
        else:
            if not is_formula(get_val(header, 11)):
                ws.cell(row=header, column=11).value = f"=K{ruk}"
                changes["k"] += 1
            if not is_formula(get_val(header, 12)):
                ws.cell(row=header, column=12).value = f"=L{ruk}"
                changes["l"] += 1

    for pr in ruk_pos:
        if not is_formula(get_val(pr, 12)):
            ws.cell(row=pr, column=12).value = f"=I{pr}*K{pr}"
            changes["l"] += 1

    for pr in rab_pos:
        if not is_formula(get_val(pr, 12)):
            ws.cell(row=pr, column=12).value = f"=I{pr}*K{pr}"
            changes["l"] += 1

    return changes


def identify_groups():
    groups = []
    i = DATA_START_ROW
    while i <= max_row:
        if is_grand_total_row(i):
            break
        if is_division_header(i):
            div_header = i
            ruk_row = None
            ruk_pos_rows = []
            rabochie_row = None
            rab_pos_rows = []
            subgroups = []
            j = i + 1
            while j <= max_row:
                if is_rukovoditel_row(j) and ruk_row is None:
                    ruk_row = j
                    j += 1
                    continue
                if is_rabochie_row(j) and ruk_row is not None and rabochie_row is None:
                    rabochie_row = j
                    j += 1
                    continue
                if ruk_row is not None and rabochie_row is None and is_position_row(j):
                    ruk_pos_rows.append(j)
                    j += 1
                    continue
                if ruk_row is not None and is_position_row(j):
                    rab_pos_rows.append(j)
                    j += 1
                    continue
                if ruk_row is not None and is_subgroup_header(j):
                    # Found sub-groups after the main structure
                    subgroups, j = parse_subgroup_structure(j, max_row)
                    continue
                break
            groups.append({
                "div_header": div_header,
                "ruk_row": ruk_row,
                "ruk_pos_rows": ruk_pos_rows,
                "rabochie_row": rabochie_row,
                "rab_pos_rows": rab_pos_rows,
                "subgroups": subgroups,
            })
            i = j
        else:
            i += 1
    return groups


def find_grand_total_row():
    for r in range(DATA_START_ROW, max_row + 1):
        if is_grand_total_row(r):
            return r
    return None


groups = identify_groups()
grand_total_row = find_grand_total_row()

print(f"Found {len(groups)} division groups")
print(f"Grand total row: {grand_total_row}")

sg_count = sum(len(g["subgroups"]) for g in groups)
print(f"Total sub-groups: {sg_count}")

for idx, g in enumerate(groups):
    dh = g["div_header"]
    rr = g["ruk_row"]
    rpr = g["ruk_pos_rows"]
    rb = g["rabochie_row"]
    rbp = g["rab_pos_rows"]
    sgs = g["subgroups"]
    dh_name = str(get_val(dh, 2))[:40]
    info = f"div_header={dh} ('{dh_name}'), ruk_row={rr}, ruk_pos={len(rpr)}"
    if rb:
        info += f", rabochie_row={rb}, rab_pos={len(rbp)}"
    if sgs:
        info += f", subgroups={len(sgs)}"
    print(f"  Group {idx}: {info}")

changes_k = 0
changes_l = 0

for g in groups:
    ruk_row = g["ruk_row"]
    ruk_pos_rows = g["ruk_pos_rows"]
    rabochie_row = g["rabochie_row"]
    rab_pos_rows = g["rab_pos_rows"]
    div_header = g["div_header"]
    subgroups = g["subgroups"]

    # --- Rukovoditel row ---
    if ruk_row is not None and ruk_pos_rows:
        ps, pe = ruk_pos_rows[0], ruk_pos_rows[-1]
        ws.cell(row=ruk_row, column=11).value = f"=SUM(K{ps}:K{pe})"
        changes_k += 1
        ws.cell(row=ruk_row, column=12).value = f"=SUM(L{ps}:L{pe})"
        changes_l += 1

    # --- Rabochie row ---
    if rabochie_row is not None:
        # Find the last position row (either in rab_pos_rows or in sub-groups)
        all_rab_positions = list(rab_pos_rows)
        for sg in subgroups:
            all_rab_positions.extend(sg.get("ruk_pos", []))
            all_rab_positions.extend(sg.get("rab_pos", []))

        if all_rab_positions:
            ps, pe = min(all_rab_positions), max(all_rab_positions)
            ws.cell(row=rabochie_row, column=11).value = f"=SUM(K{ps}:K{pe})"
            changes_k += 1
            ws.cell(row=rabochie_row, column=12).value = f"=SUM(L{ps}:L{pe})"
            changes_l += 1

    # --- Sub-groups ---
    for sg in subgroups:
        sg_changes = process_subgroup(sg)
        changes_k += sg_changes["k"]
        changes_l += sg_changes["l"]

    # --- Division header ---
    if ruk_row is not None:
        existing_k = get_val(div_header, 11)
        existing_l = get_val(div_header, 12)
        if rabochie_row is not None:
            if not is_formula(existing_k):
                ws.cell(row=div_header, column=11).value = f"=K{ruk_row}+K{rabochie_row}"
                changes_k += 1
            if not is_formula(existing_l):
                ws.cell(row=div_header, column=12).value = f"=L{ruk_row}+L{rabochie_row}"
                changes_l += 1
        else:
            if not is_formula(existing_k):
                ws.cell(row=div_header, column=11).value = f"=K{ruk_row}"
                changes_k += 1
            if not is_formula(existing_l):
                ws.cell(row=div_header, column=12).value = f"=L{ruk_row}"
                changes_l += 1

    # --- Position rows under Rukovoditel ---
    for pr in ruk_pos_rows:
        if not is_formula(get_val(pr, 12)):
            ws.cell(row=pr, column=12).value = f"=I{pr}*K{pr}"
            changes_l += 1

    # --- Position rows under Rabochie ---
    for pr in rab_pos_rows:
        if not is_formula(get_val(pr, 12)):
            ws.cell(row=pr, column=12).value = f"=I{pr}*K{pr}"
            changes_l += 1

# --- Grand total row ---
if grand_total_row is not None:
    div_headers = [g["div_header"] for g in groups if g["div_header"] is not None]
    if div_headers:
        if not is_formula(get_val(grand_total_row, 10)):
            j_parts = "+".join([f"J{dh}" for dh in div_headers])
            ws.cell(row=grand_total_row, column=10).value = f"={j_parts}"
        if not is_formula(get_val(grand_total_row, 12)):
            l_parts = "+".join([f"L{dh}" for dh in div_headers])
            ws.cell(row=grand_total_row, column=12).value = f"={l_parts}"
            changes_l += 1

    ruk_rows = [g["ruk_row"] for g in groups if g["ruk_row"] is not None]
    if ruk_rows:
        if not is_formula(get_val(grand_total_row, 11)):
            k_parts = "+".join([f"K{rr}" for rr in ruk_rows])
            ws.cell(row=grand_total_row, column=11).value = f"={k_parts}"
            changes_k += 1

print(f"\nChanges made: K={changes_k}, L={changes_l}")

wb.save(OUTPUT_FILE)
print(f"\nSaved to: {OUTPUT_FILE}")
