import json
import sys

sys.stdout.reconfigure(encoding='utf-8')

print("Building latest raffle_spinner.html from template...")

with open('raffle_data.json', 'r', encoding='utf-8') as f:
    json_data = json.load(f)

json_str = json.dumps(json_data, ensure_ascii=False)

with open('raffle_spinner_template.html', 'r', encoding='utf-8') as f:
    template = f.read()

final_html = template.replace('__RAFFLE_DATA_PLACEHOLDER__', json_str)

with open('raffle_spinner.html', 'w', encoding='utf-8-sig') as f:
    f.write(final_html)

print("✓ Successfully built raffle_spinner.html!")
