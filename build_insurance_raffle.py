"""Builds raffle_insurance.html - a separate, self-contained raffle for the
"chance insurance" draws: every participant takes part in every one of the
prizes, weighted by the number of tickets they bought (the count of numbers
in the "tickets bought" column). A participant can win only once.

Runs side by side with the main raffle - it keeps its own browser storage.

usage: python build_insurance_raffle.py <participants.xlsx> [output.html] [num_prizes]
"""
import base64
import json
import sys

import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

if len(sys.argv) < 2:
    sys.exit(__doc__)

XLSX_PATH = sys.argv[1]
OUT_PATH = sys.argv[2] if len(sys.argv) > 2 else 'raffle_insurance.html'
NUM_PRIZES = int(sys.argv[3]) if len(sys.argv) > 3 else 20

HEADERS = {'first': 'שם פרטי', 'last': 'שם משפחה', 'pkg': 'חבילה', 'tickets': 'כרטיסים שנרכשו'}


def data_uri_file(path, mime):
    with open(path, 'rb') as f:
        return f'data:{mime};base64,' + base64.b64encode(f.read()).decode()


def read_participants(path):
    ws = openpyxl.load_workbook(path, data_only=True).active
    rows = list(ws.iter_rows(values_only=True))

    # Header row is not always the first one - find it by its titles
    header_idx = next(i for i, r in enumerate(rows)
                      if r and HEADERS['tickets'] in [str(c).strip() for c in r if c])
    header = [str(c).strip() if c else '' for c in rows[header_idx]]
    col = {k: header.index(v) for k, v in HEADERS.items()}

    participants = []
    next_ticket = 1
    for line_no, r in enumerate(rows[header_idx + 1:], start=header_idx + 2):
        if not any(r):
            continue
        tickets = [t.strip() for t in str(r[col['tickets']] or '').split(',') if t.strip()]
        if not all(t.isdigit() for t in tickets):
            sys.exit(f'Row {line_no}: unexpected value in tickets column: {r[col["tickets"]]!r}')
        if not tickets:
            print(f'Warning: row {line_no} has no tickets - skipped')
            continue
        name = ' '.join(str(r[col[k]]).strip() for k in ('first', 'last') if r[col[k]])
        participants.append({
            'pid': len(participants),
            'name': name,
            'pkg': str(r[col['pkg']] or '').strip(),
            'count': len(tickets),
            'first_ticket': next_ticket,
        })
        next_ticket += len(tickets)
    return participants


participants = read_participants(XLSX_PATH)
data = {
    'prizes': {str(i): f'הגרלת ביטוח סיכויים {i}' for i in range(1, NUM_PRIZES + 1)},
    'participants': participants,
}

with open('raffle_insurance_template.html', 'r', encoding='utf-8') as f:
    html = f.read()

logo = data_uri_file('shul_catalog_logo.png', 'image/png')
favicon = data_uri_file('favicon.ico', 'image/x-icon')

html = html.replace('href="favicon.ico"', f'href="{favicon}"')
html = html.replace("onerror=\"this.src='logo.png'\"", '')
html = html.replace("'shul_catalog_logo.png'", f"'{logo}'")
html = html.replace('"shul_catalog_logo.png"', f'"{logo}"')
html = html.replace('__NUM_PRIZES__', str(NUM_PRIZES))
html = html.replace('__RAFFLE_DATA_PLACEHOLDER__', json.dumps(data, ensure_ascii=False))

with open(OUT_PATH, 'w', encoding='utf-8-sig') as f:
    f.write(html)

total = sum(p['count'] for p in participants)
print(f'Built {OUT_PATH}: {NUM_PRIZES} prizes, {len(participants)} participants, {total} tickets per draw')
