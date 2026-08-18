# -*- coding: utf-8 -*-
import json
import sys

sys.stdout.reconfigure(encoding='utf-8')

with open('raffle_data.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

prizes_list = []
for code_str, item in data.items():
    code = int(code_str)
    name = item['name']
    tickets = item['tickets']
    donors = set(t['name'] for t in tickets)
    pkgs = {}
    for t in tickets:
        pkg = t.get('pkg', 'ללא')
        pkgs[pkg] = pkgs.get(pkg, 0) + 1
    
    prizes_list.append({
        'code': code,
        'name': name,
        'tickets_count': len(tickets),
        'donors_count': len(donors),
        'pkgs': pkgs
    })

prizes_list.sort(key=lambda x: x['tickets_count'], reverse=True)

with open('raffle_summary.txt', 'w', encoding='utf-8') as out:
    out.write(f"Total Raffles: {len(prizes_list)}\n")
    out.write(f"Total Tickets: {sum(x['tickets_count'] for x in prizes_list)}\n\n")
    out.write(f"{'דירוג':<5} | {'קוד':<4} | {'שם הפרס':<45} | {'כרטיסים':<8} | {'תורמים ייחודיים':<15}\n")
    out.write("-" * 85 + "\n")
    for rank, p in enumerate(prizes_list, 1):
        out.write(f"{rank:<5} | {p['code']:<4} | {p['name']:<45} | {p['tickets_count']:<8} | {p['donors_count']:<15}\n")

print("Generated raffle_summary.txt successfully!")
