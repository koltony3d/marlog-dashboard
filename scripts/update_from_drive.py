#!/usr/bin/env python3
"""
Update data.json from Google Drive "סידור חודשי 2026" sheet.
Run this whenever you add new weeks to the Google Sheet.

Usage:  python scripts/update_from_drive.py
Then:   git add data/data.json && git commit -m "עדכון שבוע X" && git push
"""

import sys, io, json, csv, base64, subprocess
from io import StringIO
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

FILE_ID   = '1KDX29q8jx2kOpNdD6pv2S4fbaGAQG0DDsdowBW1lSR4'
OUT_FILE  = Path(__file__).parent.parent / 'data' / 'data.json'

AREA_MAP = {
    101:'מרכז',102:'מרכז',103:'מרכז',105:'ירושלים',106:'מרכז',107:'מרכז',
    108:'מרכז',109:'צפון',110:'צפון',111:'מרכז',112:'ירושלים',113:'מרכז',
    114:'דרום',115:'צפון',118:'ירושלים',119:'דרום',120:'דרום',121:'מרכז',
    122:'מרכז',123:'צפון',124:'מרכז',126:'ירושלים',128:'ירושלים',130:'צפון',
    131:'צפון',132:'מרכז',133:'צפון',135:'צפון',136:'מרכז',137:'צפון',
    139:'דרום',140:'מרכז',141:'ירושלים',142:'ירושלים',143:'צפון',144:'מרכז',
    146:'מרכז',147:'מרכז',148:'צפון',149:'צפון',150:'מרכז',151:'מרכז',
    152:'מרכז',153:'צפון',154:'דרום',155:'מרכז',156:'צפון',157:'ירושלים',
    158:'צפון',159:'מרכז',160:'צפון',161:'מרכז',162:'מרכז',163:'מרכז',
    164:'דרום',165:'מרכז',166:'מרכז',167:'מרכז',168:'צפון',169:'צפון',
    170:'דרום',171:'ירושלים',173:'צפון',174:'מרכז',175:'צפון',176:'צפון',
    177:'צפון',178:'צפון',179:'מרכז',180:'מרכז',181:'מרכז',182:'מרכז',
    183:'צפון',184:'ירושלים',185:'מרכז',186:'דרום',187:'מרכז',188:'דרום',
    189:'דרום',190:'דרום',191:'מרכז',192:'מרכז',193:'מרכז',194:'דרום',
    198:'מרכז',199:'מרכז',210:'מרכז',213:'צפון',214:'דרום',216:'מרכז',
    217:'צפון',218:'צפון',219:'צפון',222:'צפון',223:'צפון',225:'דרום',
    226:'צפון',227:'ירושלים',228:'מרכז',229:'ירושלים',231:'מרכז'
}

def ti(v):
    try: return int(float(str(v).strip())) if str(v).strip() else 0
    except: return 0


def fetch_csv_via_claude():
    """Use the Claude MCP Drive tool via a helper script."""
    import urllib.request, urllib.parse
    # Export Google Sheet as CSV using public export URL
    url = f'https://docs.google.com/spreadsheets/d/{FILE_ID}/export?format=csv'
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=30) as r:
            return r.read().decode('utf-8-sig')
    except Exception as e:
        print(f'שגיאה בהורדה: {e}')
        print('ודא שהקובץ ב-Google Drive משותף (Share → Anyone with link)')
        return None


def convert(csv_text):
    deliveries = []
    reader = csv.DictReader(StringIO(csv_text))
    for row in reader:
        date_str = row.get('תאריך', '').strip()
        if not date_str:
            continue
        parts = date_str.split('/')
        if len(parts) != 3:
            continue
        day_d, month_d, year_d = int(parts[0]), int(parts[1]), int(parts[2])

        moked = str(row.get('מוקד', '')).strip()
        branch_num = ti(moked[2:]) if moked.startswith('10') and len(moked) > 2 else ti(moked)
        branch_name = row.get('שם מוקד', '').strip()

        deliveries.append({
            'week':        ti(row.get('שבוע', 0)),
            'year':        year_d,
            'month':       month_d,
            'day':         row.get('יום', '').strip(),
            'date':        date_str,
            'branch_num':  branch_num,
            'branch_name': branch_name,
            'area':        AREA_MAP.get(branch_num, ''),
            'driver_lic':  row.get('משאית', '').strip(),
            'driver_name': row.get('נהג', '').strip(),
            'יבש':         ti(row.get('יבש', 0)),
            'בצק':         ti(row.get('בצק', 0)),
            'קרטונים':     ti(row.get('קרטון', 0)),
            'פטריות':      0,
            'החזרות':      ti(row.get('החזרה', 0)),
        })
    return deliveries


def main():
    print('מוריד מ-Google Drive...')
    csv_text = fetch_csv_via_claude()
    if not csv_text:
        sys.exit(1)

    deliveries = convert(csv_text)
    OUT_FILE.parent.mkdir(exist_ok=True)
    with open(OUT_FILE, 'w', encoding='utf-8') as f:
        json.dump(deliveries, f, ensure_ascii=False, indent=2)

    weeks  = sorted(set(d['week'] for d in deliveries))
    drivers = sorted(set(d['driver_name'] for d in deliveries if d['driver_name']))
    print(f'\n✅ {len(deliveries)} רשומות | שבועות {weeks[0]}–{weeks[-1]}')
    print(f'   נהגים: {", ".join(drivers)}')
    print(f'\nעכשיו הרץ:')
    print('  git add data/data.json')
    print(f'  git commit -m "עדכון שבוע {weeks[-1]}"')
    print('  git push')


if __name__ == '__main__':
    main()
