import json, sys, io, csv, base64
from pathlib import Path
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

CACHED = r'C:\Users\kolto\.claude\projects\C--Users-kolto-AI--\eb442321-7b24-44b8-bbc1-1aeadde5eb51\tool-results\mcp-ea389a72-291f-4d5d-aecf-a2fcc9c1c114-download_file_content-1777702375016.txt'
OUT = Path(__file__).parent.parent / 'data' / 'data.json'

AREA_MAP = {
    101:'מרכז',102:'מרכז',103:'מרכז',105:'ירושלים',106:'מרכז',107:'מרכז',108:'מרכז',109:'צפון',110:'צפון',
    111:'מרכז',112:'ירושלים',113:'מרכז',114:'דרום',115:'צפון',118:'ירושלים',119:'דרום',120:'דרום',121:'מרכז',
    122:'מרכז',123:'צפון',124:'מרכז',126:'ירושלים',128:'ירושלים',130:'צפון',131:'צפון',132:'מרכז',133:'צפון',
    135:'צפון',136:'מרכז',137:'צפון',139:'דרום',140:'מרכז',141:'ירושלים',142:'ירושלים',143:'צפון',144:'מרכז',
    146:'מרכז',147:'מרכז',148:'צפון',149:'צפון',150:'מרכז',151:'מרכז',152:'מרכז',153:'צפון',154:'דרום',
    155:'מרכז',156:'צפון',157:'ירושלים',158:'צפון',159:'מרכז',160:'צפון',161:'מרכז',162:'מרכז',163:'מרכז',
    164:'דרום',165:'מרכז',166:'מרכז',167:'מרכז',168:'צפון',169:'צפון',170:'דרום',171:'ירושלים',173:'צפון',
    174:'מרכז',175:'צפון',176:'צפון',177:'צפון',178:'צפון',179:'מרכז',180:'מרכז',181:'מרכז',182:'מרכז',
    183:'צפון',184:'ירושלים',185:'מרכז',186:'דרום',187:'מרכז',188:'דרום',189:'דרום',190:'דרום',191:'מרכז',
    192:'מרכז',193:'מרכז',194:'דרום',198:'מרכז',199:'מרכז',210:'מרכז',213:'צפון',214:'דרום',216:'מרכז',
    217:'צפון',218:'צפון',219:'צפון',222:'צפון',223:'צפון',225:'דרום',226:'צפון',227:'ירושלים',228:'מרכז',
    229:'ירושלים',231:'מרכז'
}

def ti(v):
    try: return int(float(str(v).strip())) if str(v).strip() else 0
    except: return 0

j = json.load(open(CACHED, encoding='utf-8'))
content = j['content']
try:
    txt = base64.b64decode(content).decode('utf-8-sig')
except Exception:
    txt = content.lstrip('﻿')
rows = list(csv.reader(txt.splitlines()))
header = rows[0]
print('Header:', header)

idx = {h: i for i, h in enumerate(header)}

deliveries = []
for r in rows[1:]:
    if len(r) < len(header): r = r + [''] * (len(header) - len(r))
    date_str = r[idx['תאריך']].strip()
    if not date_str: continue
    p = date_str.split('/')
    if len(p) != 3: continue
    day_d, month_d, year_d = int(p[0]), int(p[1]), int(p[2])
    if year_d < 100: year_d += 2000

    moked = r[idx['מוקד']].strip()
    branch_num = ti(moked[2:]) if moked.startswith('10') and len(moked) > 2 else ti(moked)
    deliveries.append({
        'week':        ti(r[idx['שבוע']]),
        'year':        year_d,
        'month':       month_d,
        'day':         r[idx['יום']].strip(),
        'date':        date_str,
        'branch_num':  branch_num,
        'branch_name': r[idx['שם מוקד']].strip(),
        'area':        AREA_MAP.get(branch_num, ''),
        'driver_lic':  r[idx['משאית']].strip(),
        'driver_name': r[idx['נהג']].strip(),
        'יבש':         ti(r[idx['יבש']]),
        'בצק':         ti(r[idx['בצק']]),
        'קרטונים':     ti(r[idx['קרטון']]),
        'פטריות':      0,
        'החזרות':      ti(r[idx['החזרות']]),
    })

OUT.parent.mkdir(exist_ok=True)
with open(OUT, 'w', encoding='utf-8') as f:
    json.dump(deliveries, f, ensure_ascii=False, indent=2)

total_ret = sum(d['החזרות'] for d in deliveries)
weeks = sorted(set(d['week'] for d in deliveries))
drivers = {}
for d in deliveries:
    drivers[d['driver_name']] = drivers.get(d['driver_name'], 0) + d['החזרות']
print(f'\nנשמר: {len(deliveries)} רשומות, שבועות {weeks[0]}-{weeks[-1]}')
print(f'סה"כ החזרות: {total_ret}')
print(f'לפי נהג: {drivers}')
