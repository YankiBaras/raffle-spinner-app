# -*- coding: utf-8 -*-
import json
import pandas as pd

with open('raffle_data.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

summary = []
for code_str, info in data.items():
    code = int(code_str)
    name = info.get('name', '')
    tickets = info.get('tickets', [])
    summary.append({
        'קוד הגרלה': code,
        'שם הפרס': name,
        'מספר כרטיסים': len(tickets)
    })

df = pd.DataFrame(summary).sort_values(by='קוד הגרלה')
print("Total lotteries:", len(df))
print("Total tickets across all lotteries:", df['מספר כרטיסים'].sum())

print("\n--- כל ההגרלות ממוינות לפי כמות כרטיסים (מהגבוה לנמוך) ---")
df_sorted = df.sort_values(by='מספר כרטיסים', ascending=False)
for idx, row in df_sorted.iterrows():
    print(f"{row['קוד הגרלה']:2d}. {row['שם הפרס']:<45} : {row['מספר כרטיסים']} כרטיסים")

# Also check האקסל מוכן להגרלה
try:
    df_raw = pd.read_excel('האקסל מוכן להגרלה.xlsx')
    print("\nColumns in האקסל מוכן להגרלה.xlsx:", df_raw.columns.tolist())
    print("Shape:", df_raw.shape)
except Exception as e:
    print("Error reading האקסל מוכן להגרלה.xlsx:", e)
