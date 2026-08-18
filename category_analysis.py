# -*- coding: utf-8 -*-
import json
import pandas as pd
from collections import defaultdict

with open('raffle_data.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

# Total unique donors across all lotteries
all_donors = set()
donor_tickets_count = defaultdict(int)

for code_str, item in data.items():
    for t in item['tickets']:
        name = t['name']
        all_donors.add(name)
        donor_tickets_count[name] += 1

print(f"Total unique donors across entire campaign: {len(all_donors)}")
print(f"Top 5 donors by total tickets allocated:")
for name, cnt in sorted(donor_tickets_count.items(), key=lambda x: x[1], reverse=True)[:5]:
    print(f" - {name}: {cnt} tickets")

# Category mapping
categories = {
    1: 'נופש ואירוח',
    2: 'ריהוט ועיצוב הבית',
    3: 'שירותים וייעוץ',
    4: 'שירותים וייעוץ',
    5: 'נופש ואירוח',
    6: 'נופש ואירוח',
    7: 'נופש ואירוח',
    8: 'ריהוט ועיצוב הבית',
    9: 'ריהוט ועיצוב הבית',
    10: 'ריהוט ועיצוב הבית',
    11: 'כלי כסף ויודאיקה',
    12: 'כלי כסף ויודאיקה',
    13: 'ריהוט ועיצוב הבית',
    14: 'ריהוט ועיצוב הבית',
    15: 'צילום ומדיה',
    16: 'צילום ומדיה',
    17: 'תכשיטים ויופי',
    18: 'תכשיטים ויופי',
    19: 'תכשיטים ויופי',
    20: 'תכשיטים ויופי',
    21: 'ריהוט ועיצוב הבית',
    22: 'טקסטיל ונדוניה',
    23: 'ריהוט ועיצוב הבית',
    24: 'ריהוט ועיצוב הבית',
    25: 'עולם התורה וספרים',
    26: 'עולם התורה וספרים',
    27: 'שבת ויודאיקה',
    28: 'שבת ויודאיקה',
    29: 'שבת ויודאיקה',
    30: 'מוצרי חשמל',
    31: 'מוצרי חשמל',
    32: 'מוצרי חשמל',
    33: 'מוצרי חשמל',
    34: 'מוצרי חשמל וניקיון',
    35: 'מוצרי חשמל וניקיון',
    36: 'מוצרי מטבח וקולינריה',
    37: 'מוצרי מטבח וקולינריה',
    38: 'מוצרי מטבח וקולינריה',
    39: 'מוצרי מטבח וקולינריה',
    40: 'מוצרי מטבח וקולינריה',
    41: 'מוצרי מטבח וקולינריה',
    42: 'ביגוד והנעלה',
    43: 'אופטיקה וראייה',
    44: 'טיפוח ופנאי',
    45: 'ביגוד והנעלה',
    46: 'ביגוד והנעלה',
    47: 'טיפוח ופנאי',
    48: 'טיפוח ופנאי',
    49: 'תיקים ומזוודות',
    50: 'כושר ובריאות',
    51: 'צילום ומדיה',
    52: 'ביגוד והנעלה',
    53: 'משחקים ופנאי',
    54: 'מוצרי תינוקות',
    55: 'כלי עבודה ועשה זאת בעצמך',
    56: 'קולינריה ופנאי'
}

cat_stats = defaultdict(lambda: {'tickets': 0, 'raffles': 0})
for code_str, item in data.items():
    code = int(code_str)
    cat = categories.get(code, 'שונות')
    cat_stats[cat]['tickets'] += len(item['tickets'])
    cat_stats[cat]['raffles'] += 1

print("\n--- התפלגות לפי קטגוריות ---")
for cat, stats in sorted(cat_stats.items(), key=lambda x: x[1]['tickets'], reverse=True):
    print(f"{cat:<25}: {stats['tickets']:4d} כרטיסים ({stats['raffles']} הגרלות, ממוצע {stats['tickets']/stats['raffles']:.1f} להגרלה)")
