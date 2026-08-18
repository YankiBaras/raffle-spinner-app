# -*- coding: utf-8 -*-
import pandas as pd
import json

excel_path = 'כרטיסים_לפי_הגרלה.xlsx'
xl = pd.ExcelFile(excel_path)
print("Sheet names in כרטיסים_לפי_הגרלה.xlsx:", xl.sheet_names)

df_all = pd.read_excel(excel_path, sheet_name='כל_הכרטיסים')
print("Columns in כל_הכרטיסים:", df_all.columns.tolist())
print("Head:\n", df_all.head(3))

# Group by raffle
stats = []
for code, group in df_all.groupby('קוד הגרלה'):
    prize_name = group['שם הפרס'].iloc[0]
    total_tickets = len(group)
    unique_donors = group['שם המתפלל'].nunique()
    pkgs = group['חבילה'].value_counts().to_dict() if 'חבילה' in group.columns else {}
    stats.append({
        'קוד': code,
        'פרס': prize_name,
        'כרטיסים': total_tickets,
        'תורמים_ייחודיים': unique_donors,
        'חבילות': pkgs
    })

stats_df = pd.DataFrame(stats).sort_values(by='כרטיסים', ascending=False)
print("\nTop 20 Raffles:")
for idx, r in stats_df.head(20).iterrows():
    print(f"{r['קוד']:2d}. {r['פרס']:<40} | {r['כרטיסים']:3d} כרטיסים | {r['תורמים_ייחודיים']:2d} תורמים")

print("\nBottom 15 Raffles:")
for idx, r in stats_df.tail(15).iterrows():
    print(f"{r['קוד']:2d}. {r['פרס']:<40} | {r['כרטיסים']:3d} כרטיסים | {r['תורמים_ייחודיים']:2d} תורמים")
