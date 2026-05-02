import json, sys, io, csv, base64
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
p = r'C:\Users\kolto\.claude\projects\C--Users-kolto-AI--\eb442321-7b24-44b8-bbc1-1aeadde5eb51\tool-results\mcp-ea389a72-291f-4d5d-aecf-a2fcc9c1c114-download_file_content-1777702375016.txt'
j = json.load(open(p, encoding='utf-8'))
print('mimeType:', j.get('mimeType'), 'title:', j.get('title'))
content = j['content']
try:
    txt = base64.b64decode(content).decode('utf-8-sig')
except Exception:
    txt = content if not content.startswith('﻿') else content[1:]
rows = list(csv.reader(txt.splitlines()))
print('total rows:', len(rows))
print('HEADER:', rows[0])
print('cols:', len(rows[0]))
print('--- 5 rows with non-empty החזרות ---')
hi = rows[0].index('החזרות') if 'החזרות' in rows[0] else -1
print('idx החזרות:', hi)
shown = 0
for r in rows[1:]:
    if hi >= 0 and hi < len(r) and r[hi].strip() and r[hi].strip() != '0':
        print(r)
        shown += 1
        if shown >= 8: break
# aggregate per driver
from collections import Counter
totals = Counter()
ni = rows[0].index('נהג')
for r in rows[1:]:
    if hi < len(r) and r[hi].strip():
        try:
            totals[r[ni]] += int(float(r[hi]))
        except: pass
print('TOTALS:', dict(totals))
print('total rows w/ returns:', sum(1 for r in rows[1:] if hi<len(r) and r[hi].strip() and r[hi].strip()!='0'))
