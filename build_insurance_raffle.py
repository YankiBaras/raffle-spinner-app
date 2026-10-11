"""Builds raffle_insurance.html - a separate, self-contained raffle for the
"chance insurance" draws: every participant takes part in every one of the
prizes, weighted by the number of tickets they bought (the count of numbers
in the "tickets bought" column). A participant can win only once.

Runs side by side with the main raffle - it keeps its own browser storage.
The storage key includes a hash of the data, so a rebuilt file (other
participants or prizes) starts with a clean winners list instead of reading
winners recorded by an older build in the same browser.

usage: python build_insurance_raffle.py <participants.xlsx>
           [--out raffle_insurance.html] [--prizes 20] [--add "name=tickets" ...]

--add appends a participant who is missing from the Excel file (repeatable).
"""
import argparse
import base64
import hashlib
import json
import sys

import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

parser = argparse.ArgumentParser(usage=__doc__)
parser.add_argument('xlsx')
parser.add_argument('--out', default='raffle_insurance.html')
parser.add_argument('--prizes', type=int, default=20)
parser.add_argument('--add', action='append', default=[], metavar='NAME=TICKETS')
args = parser.parse_args()

XLSX_PATH = args.xlsx
OUT_PATH = args.out
NUM_PRIZES = args.prizes

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

for extra in args.add:
    name, sep, count = extra.rpartition('=')
    if not sep or not name.strip() or not count.strip().isdigit() or int(count) < 1:
        sys.exit(f'--add expects "name=tickets", got {extra!r}')
    last = participants[-1]
    participants.append({
        'pid': len(participants),
        'name': name.strip(),
        'pkg': '',
        'count': int(count),
        'first_ticket': last['first_ticket'] + last['count'],
    })

names = [p['name'] for p in participants]
dups = sorted({n for n in names if names.count(n) > 1})
if dups:
    sys.exit(f'Duplicate participant names: {dups}')
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
data_json = json.dumps(data, ensure_ascii=False)
data_version = hashlib.sha256(data_json.encode('utf-8')).hexdigest()[:10]

html = html.replace('__NUM_PRIZES__', str(NUM_PRIZES))
html = html.replace('__DATA_VERSION__', data_version)
html = html.replace('__RAFFLE_DATA_PLACEHOLDER__', data_json)

with open(OUT_PATH, 'w', encoding='utf-8-sig') as f:
    f.write(html)

total = sum(p['count'] for p in participants)
print(f'Built {OUT_PATH}: {NUM_PRIZES} prizes, {len(participants)} participants, '
      f'{total} tickets per draw (data version {data_version})')
for p in participants[len(participants) - len(args.add):] if args.add else []:
    print(f'  added: {p["name"]} - {p["count"]} tickets')
