import re
import sys
import pandas as pd
from pathlib import Path
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

# Set stdout encoding for Windows console
sys.stdout.reconfigure(encoding="utf-8")

INPUT_CSV = Path("האקסל מוכן להגרלה.csv")
OUT_XLSX = Path("כרטיסים_לפי_הגרלה.xlsx")
OUT_DOCX_LIST = Path("רשימת_כרטיסים_להגרלה.docx")
OUT_DOCX_WINNERS = Path("טבלת_פרסים_למילוי_זוכי_הגרלה.docx")

print("1. Reading CSV file...")
# Read CSV, auto-skipping top empty rows
with open(INPUT_CSV, "r", encoding="utf-8-sig") as f:
    lines = [line for line in f.readlines() if line.strip()]

# Find line with headers
header_idx = None
for i, line in enumerate(lines[:10]):
    if "שם פרטי" in line and "כרטיסים שנרכשו" in line:
        header_idx = i
        break

if header_idx is None:
    raise ValueError("Header line not found in CSV")

df = pd.read_csv(INPUT_CSV, encoding="utf-8-sig", skiprows=header_idx, dtype=str)
df.columns = df.columns.str.strip()

print(f"Loaded {len(df)} rows. Columns: {df.columns.tolist()}")

# Helper to split comma-separated codes
def parse_codes(text):
    if pd.isna(text) or not str(text).strip():
        return []
    return [int(c.strip()) for c in str(text).split(",") if c.strip().isdigit()]

# Process tickets
all_tickets = []

for idx, row in df.iterrows():
    first_name = str(row.get("שם פרטי", "")).strip() if pd.notna(row.get("שם פרטי")) else ""
    last_name = str(row.get("שם משפחה", "")).strip() if pd.notna(row.get("שם משפחה")) else ""
    full_name = f"{first_name} {last_name}".strip()
    
    phone = str(row.get("טלפון", "")).strip() if pd.notna(row.get("טלפון")) else ""
    if phone.endswith(".0"):
        phone = phone[:-2]
    if phone and len(phone) == 9 and not phone.startswith("0"):
        phone = "0" + phone
        
    pkg = str(row.get("חבילה", "")).strip() if pd.notna(row.get("חבילה")) else ""
    
    codes = parse_codes(row.get("כרטיסים שנרכשו"))
    for code in codes:
        all_tickets.append({
            "קוד הגרלה": code,
            "שם המתפלל": full_name,
            "טלפון": phone,
            "חבילה": pkg
        })

tickets_df = pd.DataFrame(all_tickets)

# Sort by prize code then buyer name
tickets_df = tickets_df.sort_values(by=["קוד הגרלה", "שם המתפלל"]).reset_index(drop=True)

# Assign sequential ticket number within each prize code
tickets_df["מספר כרטיס"] = tickets_df.groupby("קוד הגרלה").cumcount() + 1

# Reorder columns
tickets_df = tickets_df[["קוד הגרלה", "מספר כרטיס", "שם המתפלל", "טלפון", "חבילה"]]

total_tickets_count = len(tickets_df)
unique_codes = sorted(tickets_df["קוד הגרלה"].unique())
print(f"Processed {total_tickets_count} tickets across {len(unique_codes)} prize codes (from {min(unique_codes)} to {max(unique_codes)}).")

# ---------------------------------------------------------
# 2. Export Excel: כרטיסים_לפי_הגרלה.xlsx
# ---------------------------------------------------------
print("2. Generating Excel file...")
summary_df = tickets_df.groupby("קוד הגרלה").agg(
    סהכ_כרטיסים=("מספר כרטיס", "count"),
    קונים_ייחודיים=("שם המתפלל", "nunique")
).reset_index()
summary_df.columns = ["קוד הגרלה", "סה\"כ כרטיסים שנמכרו", "מספר קונים ייחודיים"]

with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as writer:
    # Summary sheet
    summary_df.to_excel(writer, sheet_name="סיכום_הגרלות", index=False)
    # All tickets sheet
    tickets_df.to_excel(writer, sheet_name="כל_הכרטיסים", index=False)
    # Sheet per prize code
    for code, group in tickets_df.groupby("קוד הגרלה"):
        sheet_title = f"הגרלה {code}"
        group.to_excel(writer, sheet_name=sheet_title, index=False)

print(f"✓ Excel saved: {OUT_XLSX}")

# ---------------------------------------------------------
# Helper XML functions for Word RTL & borders
# ---------------------------------------------------------
def set_rtl(paragraph_or_table):
    pPr = paragraph_or_table._element.get_or_add_pPr() if hasattr(paragraph_or_table, "_element") and paragraph_or_table._element.tag.endswith("p") else None
    if pPr is not None:
        bidi = parse_xml(r'<w:bidi %s/>' % nsdecls('w'))
        pPr.append(bidi)

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_table_rtl(table):
    tblPr = table._tbl.tblPr
    bidiVisual = parse_xml(r'<w:bidiVisual %s/>' % nsdecls('w'))
    tblPr.append(bidiVisual)

# ---------------------------------------------------------
# 3. Export Word List: רשימת_כרטיסים_להגרלה.docx
# ---------------------------------------------------------
print("3. Generating Word document of tickets...")
doc = Document()

# Configure default style
style = doc.styles["Normal"]
style.font.name = "Arial"
style._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
style.font.size = Pt(10)

title_p = doc.add_paragraph()
set_rtl(title_p)
title_p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
run = title_p.add_run(f"רשימת כרטיסי הגרלה - בית הכנסת (סה\"כ {total_tickets_count} כרטיסים)")
run.font.size = Pt(18)
run.bold = True
run.font.color.rgb = RGBColor(26, 54, 93)

sub_p = doc.add_paragraph()
set_rtl(sub_p)
sub_p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
run_sub = sub_p.add_run(f"פירוט הכרטיסים בחלוקה לפי 56 הגרלות")
run_sub.font.size = Pt(12)
run_sub.font.italic = True

first = True
for code in unique_codes:
    group = tickets_df[tickets_df["קוד הגרלה"] == code].sort_values(by="מספר כרטיס")
    
    if not first:
        doc.add_page_break()
    first = False
    
    # Heading per prize
    hp = doc.add_paragraph()
    set_rtl(hp)
    h_run = hp.add_run(f"הגרלה קוד {code} (סה\"כ {len(group)} כרטיסים)")
    h_run.font.size = Pt(14)
    h_run.bold = True
    h_run.font.color.rgb = RGBColor(43, 108, 176)
    
    # Create Table
    tbl = doc.add_table(rows=1, cols=4, style="Table Grid")
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_rtl(tbl)
    
    hdr_cells = tbl.rows[0].cells
    headers = ["מס' כרטיס", "שם המתפלל", "טלפון", "חבילה"]
    widths = [Inches(1.0), Inches(2.5), Inches(1.5), Inches(1.5)]
    
    for i, title in enumerate(headers):
        cell = hdr_cells[i]
        cell.width = widths[i]
        set_cell_background(cell, "2B6CB0")
        p = cell.paragraphs[0]
        set_rtl(p)
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        r = p.add_run(title)
        r.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)
        
    for idx, row in group.iterrows():
        row_cells = tbl.add_row().cells
        vals = [str(row["מספר כרטיס"]), row["שם המתפלל"], row["טלפון"], row["חבילה"]]
        
        # Alternating background
        bg_color = "F7FAFC" if (row["מספר כרטיס"] % 2 == 0) else "FFFFFF"
        
        for i, val in enumerate(vals):
            cell = row_cells[i]
            cell.width = widths[i]
            if bg_color != "FFFFFF":
                set_cell_background(cell, bg_color)
            p = cell.paragraphs[0]
            set_rtl(p)
            p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER if i in [0, 2] else WD_PARAGRAPH_ALIGNMENT.RIGHT
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.add_run(val)

doc.save(OUT_DOCX_LIST)
print(f"✓ Word List saved: {OUT_DOCX_LIST}")

# ---------------------------------------------------------
# 4. Export Word Winners Fill Sheet: טבלת_פרסים_למילוי_זוכי_הגרלה.docx
# ---------------------------------------------------------
print("4. Generating Word winners fill sheet...")
wdoc = Document()

style = wdoc.styles["Normal"]
style.font.name = "Arial"
style._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
style.font.size = Pt(11)

w_title = wdoc.add_paragraph()
set_rtl(w_title)
w_title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
w_run = w_title.add_run("טבלת מילוי תוצאות ההגרלה הגדולה")
w_run.font.size = Pt(18)
w_run.bold = True
w_run.font.color.rgb = RGBColor(26, 54, 93)

wtbl = wdoc.add_table(rows=1, cols=5, style="Table Grid")
wtbl.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_rtl(wtbl)

whdr = wtbl.rows[0].cells
w_headers = ["קוד הגרלה", "סה\"כ כרטיסים", "מספר כרטיס זוכה", "שם הזוכה", "טלפון הזוכה"]
w_widths = [Inches(1.0), Inches(1.2), Inches(1.3), Inches(2.2), Inches(1.5)]

for i, title in enumerate(w_headers):
    cell = whdr[i]
    cell.width = w_widths[i]
    set_cell_background(cell, "1A365D")
    p = cell.paragraphs[0]
    set_rtl(p)
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(title)
    r.bold = True
    r.font.color.rgb = RGBColor(255, 255, 255)

for code in unique_codes:
    t_count = len(tickets_df[tickets_df["קוד הגרלה"] == code])
    row_cells = wtbl.add_row().cells
    vals = [str(code), str(t_count), "", "", ""]
    
    bg_color = "EDF2F7" if (code % 2 == 0) else "FFFFFF"
    
    for i, val in enumerate(vals):
        cell = row_cells[i]
        cell.width = w_widths[i]
        if bg_color != "FFFFFF":
            set_cell_background(cell, bg_color)
        p = cell.paragraphs[0]
        set_rtl(p)
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        p.add_run(val)

wdoc.save(OUT_DOCX_WINNERS)
print(f"✓ Word Winners Fill Sheet saved: {OUT_DOCX_WINNERS}")
print("✓ All processing completed successfully!")
