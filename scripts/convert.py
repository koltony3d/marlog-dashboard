#!/usr/bin/env python3
"""
Convert weekly xlsx files → data/data.json for the annual logistics dashboard.

Usage:
  python scripts/convert.py                    # scan weeks/ folder
  python scripts/convert.py path/to/week.xlsx  # single file
  python scripts/convert.py path/to/folder/    # custom folder

Output: data/data.json
"""

import sys, json, os, re
from pathlib import Path
import pandas as pd

ROOT = Path(__file__).parent.parent
OUT_FILE = ROOT / "data" / "data.json"

DAY_NAMES = ["ראשון", "שני", "שלישי", "רביעי", "חמישי"]

# ── helpers ──────────────────────────────────────────────────────────────────

def to_int(v):
    try:
        f = float(str(v).strip().replace(',',''))
        return int(f) if f == int(f) else 0
    except:
        return 0

def is_driver_header(cell):
    """Detect "ערן  |  74898202" or "74898202  ערן" style cells."""
    s = str(cell).strip()
    return '|' in s or bool(re.search(r'\d{7,8}', s))

def parse_driver_from_cell(cell):
    """Return (name, license) from a driver header cell."""
    s = str(cell).strip()
    if '|' in s:
        parts = [p.strip() for p in s.split('|')]
        # figure out which part is name vs license
        for i, p in enumerate(parts):
            if re.fullmatch(r'\d{7,9}', p):
                lic = p
                name = parts[1-i] if len(parts) == 2 else parts[0]
                return name, lic
    # fallback: split by whitespace, find number
    tokens = s.split()
    for t in tokens:
        if re.fullmatch(r'\d{7,9}', t):
            name = ' '.join(tok for tok in tokens if tok != t)
            return name, t
    return s, ''

# ── NEW FORMAT (שבוע_13_2026.xlsx) ───────────────────────────────────────────

def parse_new_format(xl, week_num, year):
    """Parse the modern xlsx format with separate day sheets."""
    deliveries = []

    # Read branch map (for area info)
    branch_map = {}
    try:
        df_br = pd.read_excel(xl, sheet_name='סניפים', header=None)
        for _, row in df_br.iloc[2:].iterrows():
            num = to_int(row.iloc[0])
            if num > 0:
                branch_map[num] = {'name': str(row.iloc[1]).strip(), 'area': str(row.iloc[2]).strip() if not pd.isna(row.iloc[2]) else ''}
    except: pass

    for day_name in DAY_NAMES:
        # Handle both "ראשון" and "ראשון " (trailing space)
        sheet = next((s for s in xl.sheet_names if s.strip() == day_name), None)
        if not sheet:
            continue
        df = pd.read_excel(xl, sheet_name=sheet, header=None)

        # Extract date from title row 0: "סידור שבועי  |  יום א  |  29.03.26"
        title = str(df.iloc[0, 0]) if not pd.isna(df.iloc[0, 0]) else ''
        date_match = re.search(r'(\d{2}[./]\d{2}[./]\d{2,4})', title)
        date_str = date_match.group(1) if date_match else ''

        # Figure out month/year from date
        month = year_actual = 0
        if date_str:
            parts = re.split(r'[./]', date_str)
            if len(parts) == 3:
                month = int(parts[1])
                yr = int(parts[2])
                year_actual = 2000 + yr if yr < 100 else yr

        # ── Left table (branch list): determine column offsets ──
        # Find header row (has "מספר" and "שם סניף")
        header_row = None
        for i, row in df.iterrows():
            row_vals = [str(v) for v in row.values if not pd.isna(v)]
            if 'מספר' in row_vals and 'שם סניף' in row_vals:
                header_row = i
                break
        if header_row is None:
            header_row = 3  # fallback

        # Find column indices
        hrow = df.iloc[header_row]
        col_num = next((c for c,v in enumerate(hrow) if str(v).strip()=='מספר'), 1)
        col_name = next((c for c,v in enumerate(hrow) if str(v).strip()=='שם סניף'), 2)
        col_yavesh = next((c for c,v in enumerate(hrow) if str(v).strip()=='יבש'), 3)
        col_batzek = next((c for c,v in enumerate(hrow) if str(v).strip()=='בצק'), 4)
        col_cartons = next((c for c,v in enumerate(hrow) if str(v).strip()=='קרטונים'), 5)
        col_mush = next((c for c,v in enumerate(hrow) if str(v).strip()=='פטריות'), 6)

        # Build branch quantity map from left table
        branch_qty = {}
        for _, row in df.iloc[header_row+1:].iterrows():
            num = to_int(row.iloc[col_num])
            if num <= 0: continue
            branch_qty[num] = {
                'יבש': to_int(row.iloc[col_yavesh]),
                'בצק': to_int(row.iloc[col_batzek]),
                'קרטונים': to_int(row.iloc[col_cartons]),
                'פטריות': to_int(row.iloc[col_mush]),
            }

        # ── Right tables: driver assignments ──
        # Scan columns 9+ for driver headers and branch assignments
        # Structure: col A_name | col A_yavesh | col A_batzek | NaN | col B_name | ...
        # Find which columns are "assignment columns" by looking for driver headers
        assignment_cols = []  # list of (name_col, yavesh_col, batzek_col)
        ncols = len(df.columns)
        for c in range(col_mush+2, ncols-2):
            for i in range(header_row, min(header_row+15, len(df))):
                cell = df.iloc[i, c]
                if not pd.isna(cell) and is_driver_header(cell):
                    # Check for quantity columns to the right
                    assignment_cols.append((c, c+1, c+2))
                    break

        # Deduplicate
        seen = set()
        unique_assign_cols = []
        for ac in assignment_cols:
            if ac[0] not in seen:
                seen.add(ac[0])
                unique_assign_cols.append(ac)

        # Parse driver assignments
        driver_branch_map = {}  # branch_num -> (driver_name, driver_lic)
        for col_n, col_y, col_b in unique_assign_cols:
            current_driver = None
            for i in range(header_row, len(df)):
                cell = df.iloc[i, col_n] if col_n < ncols else None
                if cell is None or pd.isna(cell):
                    continue
                cell_s = str(cell).strip()
                if is_driver_header(cell):
                    current_driver = parse_driver_from_cell(cell)
                elif current_driver and cell_s and cell_s != 'סה"כ':
                    # It's a branch name — try to match by stripping number prefix
                    clean = re.sub(r'^\d+-', '', cell_s).strip()
                    # Find matching branch number
                    matched_num = None
                    for bnum, bdata in branch_map.items():
                        if bdata['name'] == clean or bdata['name'] == cell_s:
                            matched_num = bnum
                            break
                    if matched_num is None:
                        # Try partial match
                        for bnum, bdata in branch_map.items():
                            if clean in bdata['name'] or bdata['name'] in clean:
                                matched_num = bnum
                                break
                    if matched_num:
                        driver_branch_map[matched_num] = current_driver

        # Build delivery records
        for bnum, qty in branch_qty.items():
            total = qty['יבש'] + qty['בצק'] + qty['קרטונים'] + qty['פטריות']
            if total == 0:
                continue  # skip empty rows
            binfo = branch_map.get(bnum, {'name': f'סניף {bnum}', 'area': ''})
            drv = driver_branch_map.get(bnum, ('', ''))
            deliveries.append({
                'week': week_num, 'year': year_actual or year,
                'month': month, 'day': day_name, 'date': date_str,
                'branch_num': bnum, 'branch_name': binfo['name'], 'area': binfo['area'],
                'driver_lic': drv[1], 'driver_name': drv[0],
                'יבש': qty['יבש'], 'בצק': qty['בצק'],
                'קרטונים': qty['קרטונים'], 'פטריות': qty['פטריות'],
                'החזרות': 0
            })

    return deliveries

# ── OLD FORMAT (שבוע 1 - 2026 (1).xlsx) ────────────────────────────────────

def parse_old_format(xl, week_num, year):
    """Parse the legacy format with different column structure."""
    deliveries = []

    for day_name in DAY_NAMES:
        sheet = next((s for s in xl.sheet_names if s.strip() == day_name), None)
        if not sheet:
            continue
        df = pd.read_excel(xl, sheet_name=sheet, header=None)

        # Old format: row 3 = dates header, row 4 = col labels (פטריות, מספר, ...)
        # Col 4=מספר, col 5=שם+num, col 6=יבש, col 7=בצק, col 3=פטריות, col 8=קרטונים
        date_str = ''
        try:
            date_match = re.search(r'(\d{2}[./]\d{2}[./]\d{2,4})', str(df.iloc[3, 5]))
            if date_match: date_str = date_match.group(1)
        except: pass

        month = year_actual = 0
        if date_str:
            parts = re.split(r'[./]', date_str)
            if len(parts) == 3:
                month = int(parts[1])
                yr = int(parts[2])
                year_actual = 2000 + yr if yr < 100 else yr

        # Find driver columns: scan for cells containing 8-digit numbers
        driver_cols = {}  # col_idx -> (name, license)
        for c in range(9, len(df.columns)):
            for i in range(3, min(8, len(df))):
                cell = df.iloc[i, c]
                if not pd.isna(cell) and is_driver_header(str(cell)):
                    driver_cols[c] = parse_driver_from_cell(str(cell))
                    break

        # Parse branch data rows (starting ~row 5)
        for i in range(5, len(df)):
            row = df.iloc[i]
            bnum = to_int(row.iloc[4])
            if bnum <= 0: continue

            bname_raw = str(row.iloc[5]).strip() if not pd.isna(row.iloc[5]) else ''
            bname = re.sub(r'^\d+-', '', bname_raw).strip()
            yavesh = to_int(row.iloc[6])
            batzek = to_int(row.iloc[7])
            mushrooms = to_int(row.iloc[3])
            cartons = to_int(row.iloc[8]) if len(row) > 8 else 0

            if yavesh + batzek + cartons + mushrooms == 0:
                continue

            # Find which driver column contains this branch
            drv_name = drv_lic = ''
            for c, drv in driver_cols.items():
                cell = df.iloc[i, c] if c < len(row) else None
                if cell is not None and not pd.isna(cell):
                    cell_s = str(cell).strip()
                    if cell_s and not is_driver_header(cell_s) and cell_s != 'סה"כ':
                        drv_name, drv_lic = drv
                        break

            deliveries.append({
                'week': week_num, 'year': year_actual or year,
                'month': month, 'day': day_name, 'date': date_str,
                'branch_num': bnum, 'branch_name': bname, 'area': '',
                'driver_lic': drv_lic, 'driver_name': drv_name,
                'יבש': yavesh, 'בצק': batzek,
                'קרטונים': cartons, 'פטריות': mushrooms,
                'החזרות': 0
            })

    return deliveries

# ── DETECT FORMAT & EXTRACT WEEK INFO ────────────────────────────────────────

def extract_week_year(filename, xl):
    """Try to get week number and year from filename or dashboard sheet."""
    stem = Path(filename).stem
    m = re.search(r'(\d{1,2}).*?(\d{4})', stem)
    week_num = int(m.group(1)) if m else 0
    year = int(m.group(2)) if m else 2026

    # Try dashboard sheet for more reliable data
    try:
        df = pd.read_excel(xl, sheet_name='דשבורד', header=None)
        for r in range(3):
            cell = str(df.iloc[r, 0])
            wm = re.search(r'שבוע\s+(\d+)', cell)
            ym = re.search(r'\b(20\d{2})\b', cell)
            if wm: week_num = int(wm.group(1))
            if ym: year = int(ym.group(1))
    except: pass

    return week_num, year

def is_new_format(xl):
    return 'דשבורד' in xl.sheet_names and 'סניפים' in xl.sheet_names

# ── MAIN ─────────────────────────────────────────────────────────────────────

def process_file(path):
    print(f"  ← {Path(path).name}")
    try:
        xl = pd.ExcelFile(path)
        week_num, year = extract_week_year(path, xl)
        if is_new_format(xl):
            return parse_new_format(xl, week_num, year)
        else:
            return parse_old_format(xl, week_num, year)
    except Exception as e:
        print(f"    ⚠ שגיאה: {e}")
        return []

def main():
    args = sys.argv[1:]
    if args:
        target = Path(args[0])
        if target.is_file():
            files = [target]
        else:
            files = sorted(target.glob('*.xlsx'))
    else:
        files = sorted((ROOT / 'weeks').glob('*.xlsx'))

    if not files:
        print("לא נמצאו קבצי xlsx. שים קבצים בתיקיית weeks/")
        sys.exit(1)

    print(f"\nמעבד {len(files)} קבצים...\n")
    all_deliveries = []
    for f in files:
        recs = process_file(f)
        all_deliveries.extend(recs)
        print(f"    ✓ {len(recs)} רשומות")

    # Remove duplicates (same week+day+branch)
    seen = set()
    unique = []
    for d in all_deliveries:
        key = (d['week'], d['year'], d['day'], d['branch_num'])
        if key not in seen:
            seen.add(key)
            unique.append(d)

    unique.sort(key=lambda d: (d['year'], d['week'], DAY_NAMES.index(d['day']) if d['day'] in DAY_NAMES else 99))

    OUT_FILE.parent.mkdir(exist_ok=True)
    with open(OUT_FILE, 'w', encoding='utf-8') as f:
        json.dump(unique, f, ensure_ascii=False, indent=2)

    print(f"\n✅ נשמר ל: {OUT_FILE}")
    print(f"   {len(unique)} רשומות ייחודיות מ-{len(files)} קבצים\n")

if __name__ == '__main__':
    main()
