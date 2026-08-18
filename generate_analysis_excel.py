# -*- coding: utf-8 -*-
import json
import xlsxwriter
import sys

sys.stdout.reconfigure(encoding='utf-8')

def create_full_analysis_excel(output_filename="ניתוח_כרטיסי_הגרלה_היושבת_בגנים.xlsx"):
    with open('raffle_data.json', 'r', encoding='utf-8') as f:
        raffle_data = json.load(f)

    # Category mapping
    categories_map = {
        1: 'נופש ואירוח',
        2: 'ריהוט ועיצוב הבית',
        3: 'שירותים וייעוץ פיננסי',
        4: 'שירותים ועיצוב',
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
        25: 'עולם התורה ותוכנה',
        26: 'עולם התורה ותוכנה',
        27: 'שבת ויודאיקה',
        28: 'שבת ויודאיקה',
        29: 'שבת ויודאיקה',
        30: 'מוצרי חשמל לבית',
        31: 'מוצרי חשמל לבית',
        32: 'מוצרי חשמל לבית',
        33: 'מוצרי חשמל לבית',
        34: 'מכשירי ניקיון וחשמל',
        35: 'מכשירי ניקיון וחשמל',
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
        48: 'טיפוח ואיפור',
        49: 'מזוודות ונסיעות',
        50: 'כושר ובריאות',
        51: 'צילום ומדיה',
        52: 'ביגוד והנעלה',
        53: 'משחקים ופנאי',
        54: 'מוצרי תינוקות',
        55: 'כלי עבודה ועשה זאת בעצמך',
        56: 'קולינריה ופנאי'
    }

    # Consolidated high-level category for aggregated charts
    broad_category_map = {
        'עולם התורה ותוכנה': 'עולם התורה וספרים',
        'צילום ומדיה': 'צילום ומדיה',
        'ריהוט ועיצוב הבית': 'ריהוט ועיצוב הבית',
        'כלי עבודה ועשה זאת בעצמך': 'כלי עבודה',
        'נופש ואירוח': 'נופש ואירוח',
        'מכשירי ניקיון וחשמל': 'מוצרי חשמל וניקיון',
        'מוצרי חשמל לבית': 'מוצרי חשמל וניקיון',
        'תכשיטים ויופי': 'תכשיטים וביגוד',
        'ביגוד והנעלה': 'תכשיטים וביגוד',
        'טיפוח ואיפור': 'טיפוח, פנאי ומשחקים',
        'טיפוח ופנאי': 'טיפוח, פנאי ומשחקים',
        'משחקים ופנאי': 'טיפוח, פנאי ומשחקים',
        'קולינריה ופנאי': 'טיפוח, פנאי ומשחקים',
        'כלי כסף ויודאיקה': 'יודאיקה ושבת',
        'שבת ויודאיקה': 'יודאיקה ושבת',
        'טקסטיל ונדוניה': 'ריהוט ועיצוב הבית',
        'מוצרי מטבח וקולינריה': 'מוצרי מטבח וכלי בית',
        'מזוודות ונסיעות': 'שונות ופנאי',
        'מוצרי תינוקות': 'שונות ופנאי',
        'אופטיקה וראייה': 'שונות ופנאי',
        'כושר ובריאות': 'שונות ופנאי',
        'שירותים וייעוץ פיננסי': 'שירותים וייעוץ',
        'שירותים ועיצוב': 'שירותים וייעוץ'
    }

    rows = []
    total_tickets = 0
    all_donors = set()

    for code_str, item in raffle_data.items():
        code = int(code_str)
        name = item['name']
        tickets = item['tickets']
        ticket_count = len(tickets)
        total_tickets += ticket_count
        donors = set(t['name'] for t in tickets)
        all_donors.update(donors)
        donor_count = len(donors)
        avg_per_donor = ticket_count / donor_count if donor_count > 0 else 0
        cat = categories_map.get(code, 'שונות')
        broad_cat = broad_category_map.get(cat, cat)
        
        rows.append({
            'code': code,
            'name': name,
            'category': cat,
            'broad_category': broad_cat,
            'tickets': ticket_count,
            'donors': donor_count,
            'avg_per_donor': avg_per_donor
        })

    # Sort rows by tickets descending
    rows_by_popularity = sorted(rows, key=lambda x: x['tickets'], reverse=True)
    for rank, r in enumerate(rows_by_popularity, 1):
        r['rank'] = rank
        r['percent'] = r['tickets'] / total_tickets if total_tickets > 0 else 0
        if r['tickets'] >= 200:
            r['tier'] = 'ביקוש שיא (טופ 4)'
        elif r['tickets'] >= 130:
            r['tier'] = 'ביקוש גבוה מאוד'
        elif r['tickets'] >= 70:
            r['tier'] = 'ביקוש גבוה'
        elif r['tickets'] >= 40:
            r['tier'] = 'ביקוש בינוני'
        else:
            r['tier'] = 'ביקוש נמוך'

    # Rows by code
    rows_by_code = sorted(rows_by_popularity, key=lambda x: x['code'])

    # Category aggregations
    cat_summary = {}
    for r in rows_by_popularity:
        c = r['category']
        if c not in cat_summary:
            cat_summary[c] = {'tickets': 0, 'raffles_count': 0, 'donors_set': set()}
        cat_summary[c]['tickets'] += r['tickets']
        cat_summary[c]['raffles_count'] += 1
        # Add donors from raffle
        item = raffle_data[str(r['code'])]
        for t in item['tickets']:
            cat_summary[c]['donors_set'].add(t['name'])

    cat_rows = []
    for c, data in cat_summary.items():
        cat_rows.append({
            'category': c,
            'tickets': data['tickets'],
            'percent': data['tickets'] / total_tickets,
            'raffles_count': data['raffles_count'],
            'avg_per_raffle': data['tickets'] / data['raffles_count'],
            'unique_donors': len(data['donors_set'])
        })
    cat_rows.sort(key=lambda x: x['tickets'], reverse=True)

    # Workbook setup
    workbook = xlsxwriter.Workbook(output_filename)
    
    # Formats
    title_fmt = workbook.add_format({
        'bold': True,
        'font_size': 16,
        'font_name': 'Segoe UI',
        'font_color': '#FFFFFF',
        'bg_color': '#1E3A8A', # Deep navy blue
        'align': 'center',
        'valign': 'vcenter'
    })
    
    subtitle_fmt = workbook.add_format({
        'font_size': 10,
        'font_name': 'Segoe UI',
        'font_color': '#DBEAFE',
        'bg_color': '#1E3A8A',
        'align': 'center',
        'valign': 'vcenter'
    })

    kpi_title_fmt = workbook.add_format({
        'font_size': 9,
        'font_name': 'Segoe UI',
        'font_color': '#475569',
        'bg_color': '#F1F5F9',
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'border_color': '#CBD5E1'
    })

    kpi_val_fmt = workbook.add_format({
        'bold': True,
        'font_size': 14,
        'font_name': 'Segoe UI',
        'font_color': '#1E3A8A',
        'bg_color': '#F8FAFC',
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'border_color': '#CBD5E1',
        'num_format': '#,##0'
    })

    kpi_val_dec_fmt = workbook.add_format({
        'bold': True,
        'font_size': 14,
        'font_name': 'Segoe UI',
        'font_color': '#1E3A8A',
        'bg_color': '#F8FAFC',
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'border_color': '#CBD5E1',
        'num_format': '0.0'
    })

    header_fmt = workbook.add_format({
        'bold': True,
        'font_name': 'Segoe UI',
        'font_size': 11,
        'font_color': '#FFFFFF',
        'bg_color': '#1E3A8A',
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'border_color': '#0F172A',
        'text_wrap': True
    })

    cell_center = workbook.add_format({
        'font_name': 'Segoe UI',
        'font_size': 10,
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'border_color': '#E2E8F0'
    })

    cell_left = workbook.add_format({
        'font_name': 'Segoe UI',
        'font_size': 10,
        'align': 'right',
        'valign': 'vcenter',
        'border': 1,
        'border_color': '#E2E8F0'
    })

    cell_bold_center = workbook.add_format({
        'bold': True,
        'font_name': 'Segoe UI',
        'font_size': 10,
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'border_color': '#E2E8F0'
    })

    cell_int = workbook.add_format({
        'font_name': 'Segoe UI',
        'font_size': 10,
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'border_color': '#E2E8F0',
        'num_format': '#,##0'
    })

    cell_bold_int = workbook.add_format({
        'bold': True,
        'font_name': 'Segoe UI',
        'font_size': 10,
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'border_color': '#E2E8F0',
        'num_format': '#,##0'
    })

    cell_percent = workbook.add_format({
        'font_name': 'Segoe UI',
        'font_size': 10,
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'border_color': '#E2E8F0',
        'num_format': '0.0%'
    })

    cell_decimal = workbook.add_format({
        'font_name': 'Segoe UI',
        'font_size': 10,
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'border_color': '#E2E8F0',
        'num_format': '0.0'
    })

    total_fmt = workbook.add_format({
        'bold': True,
        'font_name': 'Segoe UI',
        'font_size': 11,
        'bg_color': '#E2E8F0',
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'border_color': '#94A3B8',
        'num_format': '#,##0'
    })

    total_pct_fmt = workbook.add_format({
        'bold': True,
        'font_name': 'Segoe UI',
        'font_size': 11,
        'bg_color': '#E2E8F0',
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'border_color': '#94A3B8',
        'num_format': '0.0%'
    })

    total_dec_fmt = workbook.add_format({
        'bold': True,
        'font_name': 'Segoe UI',
        'font_size': 11,
        'bg_color': '#E2E8F0',
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'border_color': '#94A3B8',
        'num_format': '0.0'
    })

    # Tier badge formats
    tier_colors = {
        'ביקוש שיא (טופ 4)': workbook.add_format({'font_name': 'Segoe UI', 'font_size': 9, 'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#E2E8F0', 'bg_color': '#DCFCE7', 'font_color': '#166534'}),
        'ביקוש גבוה מאוד': workbook.add_format({'font_name': 'Segoe UI', 'font_size': 9, 'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#E2E8F0', 'bg_color': '#E0F2FE', 'font_color': '#0369A1'}),
        'ביקוש גבוה': workbook.add_format({'font_name': 'Segoe UI', 'font_size': 9, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#E2E8F0', 'bg_color': '#F0FDF4', 'font_color': '#15803D'}),
        'ביקוש בינוני': workbook.add_format({'font_name': 'Segoe UI', 'font_size': 9, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#E2E8F0', 'bg_color': '#FEF9C3', 'font_color': '#854D0E'}),
        'ביקוש נמוך': workbook.add_format({'font_name': 'Segoe UI', 'font_size': 9, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#E2E8F0', 'bg_color': '#FEE2E2', 'font_color': '#991B1B'})
    }

    # ==========================================
    # SHEET 1: דירוג פופולריות ותרשים
    # ==========================================
    ws1 = workbook.add_worksheet('דירוג פופולריות ותרשים')
    ws1.right_to_left()
    ws1.set_zoom(90)

    # Title block
    ws1.merge_range('A1:I1', 'מפעל התורה "היושבת בגנים" - ניתוח כרטיסי הגרלה ודירוג פופולריות', title_fmt)
    ws1.merge_range('A2:I2', 'דוח רשמי: התפלגות 5,814 כרטיסים על פני 56 הגרלות הפרסים', subtitle_fmt)
    ws1.set_row(0, 28)
    ws1.set_row(1, 18)

    # KPI Cards
    # Card 1: Total tickets
    ws1.merge_range('A4:B4', 'סה"כ כרטיסים שנרכשו', kpi_title_fmt)
    ws1.merge_range('A5:B5', total_tickets, kpi_val_fmt)

    # Card 2: Total lotteries
    ws1.merge_range('C4:D4', 'מספר הגרלות', kpi_title_fmt)
    ws1.merge_range('C5:D5', len(rows), kpi_val_fmt)

    # Card 3: Unique donors
    ws1.merge_range('E4:F4', 'סה"כ תורמים משתתפים', kpi_title_fmt)
    ws1.merge_range('E5:F5', len(all_donors), kpi_val_fmt)

    # Card 4: Avg per lottery
    ws1.write('G4', 'ממוצע כרטיסים להגרלה', kpi_title_fmt)
    ws1.write('G5', total_tickets / len(rows), kpi_val_dec_fmt)

    # Card 5: Median
    ws1.write('H4', 'חציון כרטיסים', kpi_title_fmt)
    median_val = (rows_by_popularity[27]['tickets'] + rows_by_popularity[28]['tickets']) / 2
    ws1.write('H5', median_val, kpi_val_dec_fmt)

    # Card 6: Top lottery
    ws1.write('I4', 'ההגרלה המבוקשת ביותר', kpi_title_fmt)
    ws1.write('I5', f"{rows_by_popularity[0]['name']} ({rows_by_popularity[0]['tickets']})", workbook.add_format({
        'bold': True, 'font_size': 10, 'font_name': 'Segoe UI', 'font_color': '#166534', 'bg_color': '#DCFCE7',
        'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#CBD5E1'
    }))

    ws1.set_row(3, 18)
    ws1.set_row(4, 26)

    # Table Headers
    headers = [
        'דירוג',
        'קוד הגרלה',
        'שם הפרס',
        'קטגוריה',
        'מספר כרטיסים',
        'אחוז מסך הכרטיסים',
        'תורמים משתתפים',
        'ממוצע לתורם',
        'רמת ביקוש'
    ]
    header_row = 6
    ws1.set_row(header_row, 26)
    for col_idx, h in enumerate(headers):
        ws1.write(header_row, col_idx, h, header_fmt)

    # Data Rows
    start_data_row = 7
    for i, r in enumerate(rows_by_popularity):
        curr_row = start_data_row + i
        ws1.set_row(curr_row, 20)
        
        ws1.write(curr_row, 0, r['rank'], cell_bold_center)
        ws1.write(curr_row, 1, r['code'], cell_center)
        ws1.write(curr_row, 2, r['name'], cell_left)
        ws1.write(curr_row, 3, r['category'], cell_center)
        ws1.write(curr_row, 4, r['tickets'], cell_bold_int)
        ws1.write_formula(curr_row, 5, f"=E{curr_row+1}/$E${start_data_row + len(rows_by_popularity) + 1}", cell_percent, r['percent'])
        ws1.write(curr_row, 6, r['donors'], cell_int)
        ws1.write_formula(curr_row, 7, f"=E{curr_row+1}/G{curr_row+1}", cell_decimal, r['avg_per_donor'])
        ws1.write(curr_row, 8, r['tier'], tier_colors.get(r['tier'], cell_center))

    # Total Row
    total_row_idx = start_data_row + len(rows_by_popularity)
    ws1.set_row(total_row_idx, 24)
    ws1.write(total_row_idx, 0, 'סה"כ', total_fmt)
    ws1.write(total_row_idx, 1, f"{len(rows)} הגרלות", total_fmt)
    ws1.write(total_row_idx, 2, '-', total_fmt)
    ws1.write(total_row_idx, 3, '-', total_fmt)
    ws1.write_formula(total_row_idx, 4, f"=SUM(E{start_data_row+1}:E{total_row_idx})", total_fmt, total_tickets)
    ws1.write_formula(total_row_idx, 5, f"=SUM(F{start_data_row+1}:F{total_row_idx})", total_pct_fmt, 1.0)
    ws1.write(total_row_idx, 6, len(all_donors), total_fmt)
    ws1.write_formula(total_row_idx, 7, f"=AVERAGE(H{start_data_row+1}:H{total_row_idx})", total_dec_fmt, total_tickets/len(rows_by_popularity))
    ws1.write(total_row_idx, 8, '-', total_fmt)

    # Column widths
    col_widths = [8, 12, 42, 24, 16, 18, 16, 14, 20]
    for idx, w in enumerate(col_widths):
        ws1.set_column(idx, idx, w)

    # Top 15 Chart embedded next to table or below KPIs
    chart1 = workbook.add_chart({'type': 'bar'})
    chart1.add_series({
        'name': 'מספר כרטיסים שנרכשו',
        'categories': ['דירוג פופולריות ותרשים', start_data_row, 2, start_data_row + 14, 2],
        'values': ['דירוג פופולריות ותרשים', start_data_row, 4, start_data_row + 14, 4],
        'fill': {'color': '#1E3A8A'},
        'data_labels': {'value': True, 'font': {'name': 'Segoe UI', 'size': 9, 'bold': True, 'color': '#0F172A'}}
    })
    chart1.set_title({
        'name': '15 ההגרלות המבוקשות ביותר (Top 15)',
        'name_font': {'name': 'Segoe UI', 'size': 14, 'bold': True, 'color': '#1E3A8A'}
    })
    chart1.set_x_axis({
        'name': 'מספר כרטיסים',
        'name_font': {'name': 'Segoe UI', 'size': 10, 'bold': True},
        'major_gridlines': {'visible': True, 'line': {'color': '#E2E8F0'}}
    })
    chart1.set_y_axis({
        'reverse_categories': True,
        'num_font': {'name': 'Segoe UI', 'size': 9}
    })
    chart1.set_legend({'position': 'none'})
    chart1.set_size({'width': 780, 'height': 560})
    ws1.insert_chart('K4', chart1, {'x_offset': 10, 'y_offset': 5})

    # ==========================================
    # SHEET 2: פילוח לפי קטגוריות
    # ==========================================
    ws2 = workbook.add_worksheet('פילוח לפי קטגוריות')
    ws2.right_to_left()
    ws2.set_zoom(95)

    ws2.merge_range('A1:F1', 'פילוח וניתוח כרטיסים לפי קטגוריות תוכן', title_fmt)
    ws2.merge_range('A2:F2', 'השוואת ביצועים ונתח שוק של תחומי הפרסים השונים', subtitle_fmt)
    ws2.set_row(0, 28)
    ws2.set_row(1, 18)

    cat_headers = [
        'דירוג',
        'קטגוריה',
        'מספר הגרלות בקטגוריה',
        'סה"כ כרטיסים',
        'אחוז מכלל הכרטיסים',
        'ממוצע כרטיסים להגרלה'
    ]
    cat_header_row = 4
    ws2.set_row(cat_header_row, 24)
    for col_idx, h in enumerate(cat_headers):
        ws2.write(cat_header_row, col_idx, h, header_fmt)

    cat_start_row = 5
    for i, cr in enumerate(cat_rows):
        curr_row = cat_start_row + i
        ws2.set_row(curr_row, 20)
        ws2.write(curr_row, 0, i + 1, cell_bold_center)
        ws2.write(curr_row, 1, cr['category'], cell_left)
        ws2.write(curr_row, 2, cr['raffles_count'], cell_int)
        ws2.write(curr_row, 3, cr['tickets'], cell_bold_int)
        ws2.write_formula(curr_row, 4, f"=D{curr_row+1}/$D${cat_start_row + len(cat_rows) + 1}", cell_percent, cr['percent'])
        ws2.write_formula(curr_row, 5, f"=D{curr_row+1}/C{curr_row+1}", cell_decimal, cr['avg_per_raffle'])

    cat_total_row = cat_start_row + len(cat_rows)
    ws2.set_row(cat_total_row, 24)
    ws2.write(cat_total_row, 0, 'סה"כ', total_fmt)
    ws2.write(cat_total_row, 1, f"{len(cat_rows)} קטגוריות", total_fmt)
    ws2.write_formula(cat_total_row, 2, f"=SUM(C{cat_start_row+1}:C{cat_total_row})", total_fmt, 56)
    ws2.write_formula(cat_total_row, 3, f"=SUM(D{cat_start_row+1}:D{cat_total_row})", total_fmt, total_tickets)
    ws2.write_formula(cat_total_row, 4, f"=SUM(E{cat_start_row+1}:E{cat_total_row})", total_pct_fmt, 1.0)
    ws2.write_formula(cat_total_row, 5, f"=AVERAGE(F{cat_start_row+1}:F{cat_total_row})", total_dec_fmt, total_tickets/56)

    cat_col_widths = [8, 28, 22, 16, 20, 22]
    for idx, w in enumerate(cat_col_widths):
        ws2.set_column(idx, idx, w)

    # Category Chart
    chart2 = workbook.add_chart({'type': 'column'})
    chart2.add_series({
        'name': 'סה"כ כרטיסים לקטגוריה',
        'categories': ['פילוח לפי קטגוריות', cat_start_row, 1, cat_start_row + len(cat_rows) - 1, 1],
        'values': ['פילוח לפי קטגוריות', cat_start_row, 3, cat_start_row + len(cat_rows) - 1, 3],
        'fill': {'color': '#2563EB'},
        'data_labels': {'value': True, 'font': {'name': 'Segoe UI', 'size': 9, 'bold': True}}
    })
    chart2.set_title({
        'name': 'התפלגות כרטיסים לפי קטגוריות פרסים',
        'name_font': {'name': 'Segoe UI', 'size': 13, 'bold': True, 'color': '#1E3A8A'}
    })
    chart2.set_x_axis({
        'name_font': {'name': 'Segoe UI', 'size': 9},
        'num_font': {'name': 'Segoe UI', 'size': 9, 'rotation': -45}
    })
    chart2.set_y_axis({
        'name': 'מספר כרטיסים',
        'major_gridlines': {'visible': True, 'line': {'color': '#E2E8F0'}}
    })
    chart2.set_legend({'position': 'none'})
    chart2.set_size({'width': 780, 'height': 420})
    ws2.insert_chart('H4', chart2, {'x_offset': 10, 'y_offset': 5})

    # ==========================================
    # SHEET 3: רשימה לפי סדר קוד הגרלה (1-56)
    # ==========================================
    ws3 = workbook.add_worksheet('לפי סדר קטלוג (1-56)')
    ws3.right_to_left()
    ws3.set_zoom(90)

    ws3.merge_range('A1:H1', 'רשימת ההגרלות לפי סדר קטלוגי (קוד 1 עד 56)', title_fmt)
    ws3.merge_range('A2:H2', 'עיון מהיר בנתוני כל פרס לפי מספרו הסידורי בקטלוג המתנות', subtitle_fmt)
    ws3.set_row(0, 28)
    ws3.set_row(1, 18)

    headers_code = [
        'קוד הגרלה',
        'שם הפרס',
        'קטגוריה',
        'מספר כרטיסים',
        'אחוז מסך הכרטיסים',
        'תורמים משתתפים',
        'ממוצע לתורם',
        'דירוג פופולריות'
    ]
    ws3.set_row(3, 24)
    for col_idx, h in enumerate(headers_code):
        ws3.write(3, col_idx, h, header_fmt)

    start_c_row = 4
    for i, r in enumerate(rows_by_code):
        curr_row = start_c_row + i
        ws3.set_row(curr_row, 19)
        ws3.write(curr_row, 0, r['code'], cell_bold_center)
        ws3.write(curr_row, 1, r['name'], cell_left)
        ws3.write(curr_row, 2, r['category'], cell_center)
        ws3.write(curr_row, 3, r['tickets'], cell_bold_int)
        ws3.write_formula(curr_row, 4, f"=D{curr_row+1}/$D${start_c_row + len(rows_by_code) + 1}", cell_percent, r['percent'])
        ws3.write(curr_row, 5, r['donors'], cell_int)
        ws3.write_formula(curr_row, 6, f"=D{curr_row+1}/F{curr_row+1}", cell_decimal, r['avg_per_donor'])
        ws3.write(curr_row, 7, f"#{r['rank']} מתוך 56", cell_center)

    total_c_row = start_c_row + len(rows_by_code)
    ws3.set_row(total_c_row, 24)
    ws3.write(total_c_row, 0, 'סה"כ', total_fmt)
    ws3.write(total_c_row, 1, '56 הגרלות', total_fmt)
    ws3.write(total_c_row, 2, '-', total_fmt)
    ws3.write_formula(total_c_row, 3, f"=SUM(D{start_c_row+1}:D{total_c_row})", total_fmt, total_tickets)
    ws3.write_formula(total_c_row, 4, f"=SUM(E{start_c_row+1}:E{total_c_row})", total_pct_fmt, 1.0)
    ws3.write(total_c_row, 5, len(all_donors), total_fmt)
    ws3.write_formula(total_c_row, 6, f"=AVERAGE(G{start_c_row+1}:G{total_c_row})", total_dec_fmt, total_tickets/56)
    ws3.write(total_c_row, 7, '-', total_fmt)

    for idx, w in enumerate([12, 42, 24, 16, 18, 16, 14, 18]):
        ws3.set_column(idx, idx, w)

    # Full distribution column chart
    chart3 = workbook.add_chart({'type': 'column'})
    chart3.add_series({
        'name': 'כרטיסים לפי הגרלה (1-56)',
        'categories': ['לפי סדר קטלוג (1-56)', start_c_row, 0, start_c_row + len(rows_by_code) - 1, 0],
        'values': ['לפי סדר קטלוג (1-56)', start_c_row, 3, start_c_row + len(rows_by_code) - 1, 3],
        'fill': {'color': '#3B82F6'}
    })
    chart3.set_title({
        'name': 'התפלגות כרטיסים על פני כל 56 ההגרלות (סדר קטלוגי)',
        'name_font': {'name': 'Segoe UI', 'size': 13, 'bold': True, 'color': '#1E3A8A'}
    })
    chart3.set_x_axis({
        'name': 'קוד הגרלה',
        'name_font': {'name': 'Segoe UI', 'size': 10, 'bold': True}
    })
    chart3.set_y_axis({
        'name': 'מספר כרטיסים',
        'major_gridlines': {'visible': True, 'line': {'color': '#E2E8F0'}}
    })
    chart3.set_legend({'position': 'none'})
    chart3.set_size({'width': 860, 'height': 420})
    ws3.insert_chart('J4', chart3, {'x_offset': 10, 'y_offset': 5})

    # ==========================================
    # SHEET 4: ניתוח ופרשנות
    # ==========================================
    ws4 = workbook.add_worksheet('ניתוח ופרשנות מעמיקה')
    ws4.right_to_left()
    ws4.set_zoom(100)

    ws4.merge_range('A1:G1', 'דוח פרשני: ניתוח העדפות המשתתפים וגורמי ההצלחה של ההגרלות', title_fmt)
    ws4.merge_range('A2:G2', 'תובנות סוציו-אקונומיות, דפוסי בחירה והמלצות אסטרטגיות לקמפיינים הבאים', subtitle_fmt)
    ws4.set_row(0, 28)
    ws4.set_row(1, 18)

    section_header_fmt = workbook.add_format({
        'bold': True,
        'font_name': 'Segoe UI',
        'font_size': 12,
        'font_color': '#1E3A8A',
        'bg_color': '#EFF6FF',
        'border': 1,
        'border_color': '#BFDBFE',
        'valign': 'vcenter'
    })

    text_box_fmt = workbook.add_format({
        'font_name': 'Segoe UI',
        'font_size': 10,
        'font_color': '#1E293B',
        'align': 'right',
        'valign': 'top',
        'text_wrap': True
    })

    sections = [
        ("1. תמונת המצב הכללית וריכוזיות הביקוש",
         "במסגרת קמפיין מפעל התורה 'היושבת בגנים' (גני איילון / אחיסמך) נרכשו בסך הכל 5,814 כרטיסים על פני 56 הגרלות שונות.\n"
         "ניתוח הנתונים מצביע על ריכוזיות ביקוש גבוהה במיוחד: 4 ההגרלות המובילות לבדן (אוצר החכמה, מצלמת קנון SX740, מיטת היי-ריזר, וסט כלי עבודה מקיטה) גרפו 1,111 כרטיסים - המהווים כ-19.1% מכלל הכרטיסים בקמפיין!\n"
         "10 ההגרלות המובילות גרפו 2,169 כרטיסים (37.3% מסך הכרטיסים), בעוד ש-10 ההגרלות בתחתית הרשימה גרפו יחד 323 כרטיסים בלבד (5.6%)."),

        ("2. מה הפך את ההגרלות המובילות למועדפות ביותר?",
         "א. עולם התורה וספרות תורנית דיגיטלית (אוצר החכמה - מקום 1 עם 319 כרטיסים; שו\"ת בר אילן - מקום 10 עם 163 כרטיסים):\n"
         "בקהילת היעד (אברכים ועמלי תורה), תוכנות תורניות אלו מהוות כלי עבודה יסודי ונחשק מאין כמותו, שמחירו בשוק החופשי מגיע לאלפי שקלים רבים. הזדמנות לזכות במוצר חלומי ויוקרתי זה בכרטיס הגרלה יצרה היענות שיא.\n\n"
         "ב. צילום ומדיה ללא סמארטפונים (מצלמת Canon SX740 - מקום 2 עם 308 כרטיסים; Canon SX720 - מקום 8 עם 173 כרטיסים):\n"
         "בציבור החרדי, שאינו משתמש בטלפונים חכמים, מצלמה קומפקטית איכותית ומתקדמת היא הדרך היחידה לתעד שמחות משפחתיות ורגעים יקרים. דגמי Canon PowerShot SX מהווים מוצר דגל מבוקש במיוחד ו'מוצר זהב' בתרבות הצריכה הקהילתית (סה\"כ 481 כרטיסים לשני הדגמים יחד!).\n\n"
         "ג. מענה מעשי ומדויק לצרכי הבית ברוך הילדים:\n"
         "• מיטת היי-ריזר כפולה/משולשת (מקום 3, 243 כרטיסים): פתרון אחסון ושינה קריטי בדירות מגורים צפופות למשפחות ברוכות ילדים.\n"
         "• סט כלי עבודה Makita מקצועי (מקום 4, 241 כרטיסים): מאפשר אחזקת בית עצמאית וחיסכון בעלויות אנשי מקצוע.\n"
         "• שואב שוטף Tineco S9 / מדיח DeLonghi (מקום 6, 189 כרטיסים): חיסכון בזמן עבודה יקר ומאמץ פיזי רב באחזקת הבית והכנות לשבת.\n\n"
         "ד. גמישות, נזילות וערך ודאי (תווי קניה וזיכויים):\n"
         "• יופי של תכשיטים - 1,000 ש\"ח בנובל (מקום 7, 178 כרטיסים) וזיכוי בגיר ע\"ס 1,500 ש\"ח (מקום 9, 169 כרטיסים):\n"
         "פרסים אלו מעניקים אפשרות בחירה מלאה לפי טעם אישי ומידה מדויקת, ללא סיכון לקבל פריט לא מתאים."),

        ("3. ניתוח ההגרלות בעלות הביקוש הנמוך ביותר - מדוע נרשמה היענות מועטה?",
         "• מוצרי חשמל ממותג שאינו נתפס כמותג פרימיום (מותג Germaine: מכונת כביסה - 32 כרטיסים, מקרר - 44 כרטיסים, מייבש - 53 כרטיסים):\n"
         "הציבור העדיף באופן מובהק מותגי יוקרה מוכרים (כמו Makita, Dyson/Tineco, DeLonghi, Canon) ונמנע מהשקעת כרטיסים במוצרים בסיסיים ממותגים פחות יוקרתיים.\n"
         "• שירותים נישתיים/מותנים (מחזור משכנתא - 37 כרטיסים; אדריכלות מטבח - 29 כרטיסים):\n"
         "פרסים אלו רלוונטיים רק למי שנמצא בעיצומו של תהליך שיפוץ או בעל תנאי משכנתא ספציפיים, ולכן אינם רלוונטיים לרוב המשתתפים.\n"
         "• פרסים בשווי כספי נתפס נמוך (חגיגה במטבח 2 - 19 כרטיסים; פינוק ליל שישי זוגי - 29 כרטיסים):\n"
         "המשתתפים העדיפו להפנות את כרטיסיהם להגרלות על פרסים של אלפי שקלים ולא על מוצרים קטנים שניתן לרכוש בקלות בחנות."),

        ("4. מסקנות והמלצות לקמפיינים עתידיים",
         "1. התמקדות במותגי פרימיום מובילים: שימוש במותגים בעלי שם עולמי (Canon, Makita, Bosch, DeLonghi) מגדיל משמעותית את כוח המשיכה.\n"
         "2. הגדלת היצע הפרסים התורניים: הרחבת הגרלות על ספריות תורניות, ש\"ס מפואר, ומחשבים ייעודיים ללימוד.\n"
         "3. העדפת תווי קניה וגמישות אישית על פני פריטים ספציפיים מדי שאינם מתאימים לכל בית.\n"
         "4. ביטול פרסים נישתיים בעלי תנאי מימוש מורכבים שאינם רלוונטיים למרבית הציבור.")
    ]

    curr_row = 4
    for title, content in sections:
        ws4.merge_range(curr_row, 0, curr_row, 6, title, section_header_fmt)
        ws4.set_row(curr_row, 24)
        curr_row += 1
        
        line_count = content.count('\n') + 2
        approx_height = max(60, line_count * 18)
        ws4.merge_range(curr_row, 0, curr_row + 1, 6, content, text_box_fmt)
        ws4.set_row(curr_row, approx_height // 2)
        ws4.set_row(curr_row + 1, approx_height // 2)
        curr_row += 3

    for idx, w in enumerate([18, 20, 20, 20, 20, 20, 20]):
        ws4.set_column(idx, idx, w)

    workbook.close()
    print(f"✓ Excel file '{output_filename}' created successfully!")

if __name__ == '__main__':
    create_full_analysis_excel()
