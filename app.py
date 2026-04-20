from flask import Flask, render_template, request, redirect, url_for, send_file, jsonify
import sqlite3
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os, io

app = Flask(__name__)
DATABASE = os.path.join(os.path.dirname(__file__), 'database.db')
EXCEL_SOURCE = r'C:\Users\Administrator\Desktop\androw_new.xlsx'

MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
MONTHS_AR = ['يناير','فبراير','مارس','أبريل','مايو','يونيو','يوليو','أغسطس','سبتمبر','أكتوبر','نوفمبر','ديسمبر']

# Antigens with drop-out rules
ANTIGENS = [
    {'key': 'BCG',    'label': 'BCG',         'color': '#d0e8ff'},
    {'key': 'OPV0',   'label': 'OPV 0',       'color': '#fff0cc'},
    {'key': 'Penta1', 'label': 'Penta 1',     'color': '#d0e8ff'},
    {'key': 'OPV1',   'label': 'OPV 1',       'color': '#fff0cc'},
    {'key': 'PCV1',   'label': 'PCV 1',       'color': '#d0e8ff'},
    {'key': 'Penta2', 'label': 'Penta 2',     'color': '#fff0cc'},
    {'key': 'OPV2',   'label': 'OPV 2',       'color': '#d0e8ff'},
    {'key': 'PCV2',   'label': 'PCV 2',       'color': '#fff0cc'},
    {'key': 'Penta3', 'label': 'Penta 3',     'color': '#d0e8ff'},
    {'key': 'OPV3',   'label': 'OPV 3',       'color': '#fff0cc'},
    {'key': 'PCV3',   'label': 'PCV 3',       'color': '#d0e8ff'},
    {'key': 'IPV',    'label': 'IPV',         'color': '#fff0cc'},
    {'key': 'MR1',    'label': 'MR 1 (MMR1)', 'color': '#d0e8ff'},
    {'key': 'MR2',    'label': 'MR 2 (MMR2)', 'color': '#fff0cc'},
    {'key': 'VitA1',  'label': 'Vit A (1st)', 'color': '#d0e8ff'},
    {'key': 'VitA2',  'label': 'Vit A (2nd)', 'color': '#fff0cc'},
]

STOCK_VACCINES = [
    {'key': 'BCG',   'label': 'BCG (20 جرعة/فيال)',                   'label_en': 'BCG (20 doses/vial)',                     'vial': 20},
    {'key': 'HepB',  'label': 'التهاب الكبد B (10 جرعة/فيال)',        'label_en': 'HepB Pediatric (10 doses/vial)',           'vial': 10},
    {'key': 'IPV',   'label': 'شلل الأطفال المحقون IPV (5 جرعة)',     'label_en': 'Inactivated Polio IPV (5 doses/vial)',     'vial': 5},
    {'key': 'bOPV',  'label': 'شلل الأطفال الفموي bOPV (10 جرعة)',    'label_en': 'Oral Polio bOPV (10 doses/vial)',          'vial': 10},
    {'key': 'Penta', 'label': 'الخماسي DTP-HepB-Hib (10 جرعة)',       'label_en': 'Pentavalent DTP-HepB-Hib (10 doses/vial)', 'vial': 10},
    {'key': 'Rota',  'label': 'الروتا Rotavac 2.5ml (5 جرعة)',        'label_en': 'Rotavirus Rotavac 2.5ml (5 doses/vial)',   'vial': 5},
    {'key': 'PCV',   'label': 'المكورات الرئوية PCV10 (5 جرعة)',      'label_en': 'Pneumococcal PCV10 (5 doses/vial)',        'vial': 5},
    {'key': 'MMR',   'label': 'الحصبة والحصبة الألمانية MMR (1 جرعة)','label_en': 'MMR (1 dose/vial)',                       'vial': 1},
    {'key': 'DTP18', 'label': 'ثلاثي DTP منشط (18 شهر، 10 جرعة)',     'label_en': 'DTP Booster 18m (10 doses/vial)',          'vial': 10},
    {'key': 'DT6',   'label': 'ثنائي DT (6 سنوات، 10 جرعة)',          'label_en': 'DT 6 years (10 doses/vial)',               'vial': 10},
    {'key': 'Td15',  'label': 'Td (15 سنة، 10 جرعة)',                  'label_en': 'Td 15 years (10 doses/vial)',              'vial': 10},
]

# Drop-out definitions: (label, numerator_key, denominator_key)
DROPOUTS = [
    ('Dropout 1: Penta1→Penta3',  'Penta1', 'Penta3'),
    ('Dropout 2: BCG→MR1',        'BCG',    'MR1'),
    ('Dropout 3: MR1→MR2',        'MR1',    'MR2'),
]

def get_db():
    conn = sqlite3.connect(DATABASE)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS facilities (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sno INTEGER, governorate TEXT, name TEXT,
        type TEXT, provider TEXT, status TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS target_population (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        facility_id INTEGER, year INTEGER,
        catchment_pop INTEGER DEFAULT 0,
        si_percent REAL DEFAULT 3.2,
        UNIQUE(facility_id, year)
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS immunization_data (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        facility_id INTEGER, year INTEGER,
        antigen TEXT, month INTEGER,
        doses INTEGER DEFAULT 0,
        UNIQUE(facility_id, year, antigen, month)
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS stock_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        facility_id INTEGER, year INTEGER, month INTEGER,
        antigen TEXT,
        opening INTEGER, received INTEGER, administered INTEGER, wastage INTEGER,
        UNIQUE(facility_id, year, month, antigen)
    )''')
    conn.commit()
    # Import facilities if table is empty
    c.execute('SELECT COUNT(*) FROM facilities')
    if c.fetchone()[0] == 0:
        import_facilities(conn)
    conn.close()

def import_facilities(conn=None):
    close = False
    if conn is None:
        conn = get_db()
        close = True
    try:
        df = pd.read_excel(EXCEL_SOURCE, sheet_name='Sheet1', header=1)
        c = conn.cursor()
        for _, row in df.iterrows():
            if pd.notna(row.get('Name of the Facility')):
                c.execute('''INSERT OR IGNORE INTO facilities
                    (sno, governorate, name, type, provider, status) VALUES (?,?,?,?,?,?)''',
                    (row.get('S.No.'), str(row.get('Governorate','')),
                     str(row.get('Name of the Facility','')),
                     str(row.get('Type of Facility','')),
                     str(row.get('Provider','')),
                     str(row.get('Status of Functionality',''))))
        conn.commit()
    except Exception as e:
        print(f"Import error: {e}")
    if close:
        conn.close()

# ─── Routes ───────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    year = request.args.get('year', 2025, type=int)
    search = request.args.get('search', '')
    gov = request.args.get('governorate', '')
    conn = get_db()
    query = 'SELECT * FROM facilities WHERE 1=1'
    params = []
    if search:
        query += ' AND name LIKE ?'
        params.append(f'%{search}%')
    if gov:
        query += ' AND governorate = ?'
        params.append(gov)
    query += ' ORDER BY governorate, sno'
    facilities = conn.execute(query, params).fetchall()
    govs = conn.execute('SELECT DISTINCT governorate FROM facilities ORDER BY governorate').fetchall()
    conn.close()
    return render_template('index.html', facilities=facilities, year=year,
                           governorates=govs, search=search, sel_gov=gov)

@app.route('/facility/<int:fid>', methods=['GET','POST'])
def facility(fid):
    year = request.args.get('year', 2025, type=int)
    conn = get_db()
    fac = conn.execute('SELECT * FROM facilities WHERE id=?', (fid,)).fetchone()
    if not fac:
        conn.close()
        return redirect(url_for('index'))

    if request.method == 'POST':
        year = int(request.form.get('year', 2025))
        catchment = int(request.form.get('catchment_pop', 0) or 0)
        si = float(request.form.get('si_percent', 3.2) or 3.2)
        conn.execute('''INSERT INTO target_population (facility_id, year, catchment_pop, si_percent)
            VALUES (?,?,?,?) ON CONFLICT(facility_id,year) DO UPDATE SET
            catchment_pop=excluded.catchment_pop, si_percent=excluded.si_percent''',
            (fid, year, catchment, si))
        for ag in ANTIGENS:
            for m in range(1, 13):
                val_str = request.form.get(f"{ag['key']}_{m}", '').strip()
                if val_str == '':
                    # Blank = month not yet reported → remove any existing record
                    conn.execute(
                        'DELETE FROM immunization_data WHERE facility_id=? AND year=? AND antigen=? AND month=?',
                        (fid, year, ag['key'], m))
                else:
                    val = int(val_str) if val_str.isdigit() else 0
                    conn.execute('''INSERT INTO immunization_data (facility_id,year,antigen,month,doses)
                        VALUES (?,?,?,?,?) ON CONFLICT(facility_id,year,antigen,month) DO UPDATE SET doses=excluded.doses''',
                        (fid, year, ag['key'], m, val))
        conn.commit()
        conn.close()
        return redirect(url_for('facility', fid=fid, year=year, saved=1))

    # Load existing data
    tp = conn.execute('SELECT * FROM target_population WHERE facility_id=? AND year=?',
                      (fid, year)).fetchone()
    rows = conn.execute('SELECT antigen, month, doses FROM immunization_data WHERE facility_id=? AND year=?',
                        (fid, year)).fetchall()
    conn.close()

    data = {}
    for r in rows:
        data[(r['antigen'], r['month'])] = r['doses']

    saved = request.args.get('saved', 0)
    return render_template('facility.html', fac=fac, year=year, tp=tp,
                           antigens=ANTIGENS, months=MONTHS, data=data,
                           dropouts=DROPOUTS, saved=saved)

@app.route('/add_facility', methods=['GET','POST'])
def add_facility():
    if request.method == 'POST':
        conn = get_db()
        conn.execute('''INSERT INTO facilities (governorate, name, type, provider, status)
            VALUES (?,?,?,?,?)''',
            (request.form['governorate'], request.form['name'],
             request.form['type'], request.form['provider'], 'Functional'))
        conn.commit()
        conn.close()
        return redirect(url_for('index'))
    return render_template('add_facility.html')

@app.route('/reimport')
def reimport():
    import_facilities()
    return redirect(url_for('index'))

@app.route('/export/<int:fid>')
def export_facility(fid):
    year = request.args.get('year', 2025, type=int)
    conn = get_db()
    fac = conn.execute('SELECT * FROM facilities WHERE id=?', (fid,)).fetchone()
    tp = conn.execute('SELECT * FROM target_population WHERE facility_id=? AND year=?',
                      (fid, year)).fetchone()
    rows = conn.execute('SELECT antigen, month, doses FROM immunization_data WHERE facility_id=? AND year=?',
                        (fid, year)).fetchall()
    conn.close()

    data = {}
    for r in rows:
        data[(r['antigen'], r['month'])] = r['doses']

    wb = build_excel(fac, year, tp, data)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    safe_name = str(fac['name']).replace(' ', '_').replace('/', '-')
    return send_file(buf, as_attachment=True,
                     download_name=f"EPI_{safe_name}_{year}.xlsx",
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/export_all')
def export_all():
    year = request.args.get('year', 2025, type=int)
    conn = get_db()
    facilities = conn.execute('SELECT * FROM facilities ORDER BY governorate, sno').fetchall()

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for fac in facilities:
        tp = conn.execute('SELECT * FROM target_population WHERE facility_id=? AND year=?',
                          (fac['id'], year)).fetchone()
        rows = conn.execute('SELECT antigen,month,doses FROM immunization_data WHERE facility_id=? AND year=?',
                            (fac['id'], year)).fetchall()
        data = {(r['antigen'], r['month']): r['doses'] for r in rows}
        sheet_name = str(fac['name'])[:31].replace('/', '-').replace('\\','-').replace('*','').replace('[','').replace(']','').replace(':','').replace('?','')
        add_sheet_to_wb(wb, sheet_name, fac, year, tp, data)

    conn.close()
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name=f"EPI_All_Facilities_{year}.xlsx",
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# ─── Excel Builder ────────────────────────────────────────────────────────────

HDR_FILL  = PatternFill('solid', fgColor='1FA2C8')
HDR2_FILL = PatternFill('solid', fgColor='3ABCD8')
BLUE_FILL = PatternFill('solid', fgColor='D0E8FF')
YELL_FILL = PatternFill('solid', fgColor='FFF0CC')
GRN_FILL  = PatternFill('solid', fgColor='C6EFCE')
RED_FILL  = PatternFill('solid', fgColor='FFC7CE')
WHT_FONT  = Font(name='Arial', bold=True, color='FFFFFF', size=9)
BLD_FONT  = Font(name='Arial', bold=True, size=9)
NRM_FONT  = Font(name='Arial', size=9)
CTR = Alignment(horizontal='center', vertical='center', wrap_text=True)
thin = Side(style='thin', color='888888')
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

def cell_style(ws, r, c, value='', fill=None, font=None, align=CTR, border=BORDER, number_format=None):
    cell = ws.cell(row=r, column=c, value=value)
    if fill:   cell.fill = fill
    if font:   cell.font = font
    if align:  cell.alignment = align
    if border: cell.border = border
    if number_format: cell.number_format = number_format
    return cell

def build_excel(fac, year, tp, data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = str(fac['name'])[:31]
    add_sheet_content(ws, fac, year, tp, data)
    return wb

def add_sheet_to_wb(wb, sheet_name, fac, year, tp, data):
    ws = wb.create_sheet(title=sheet_name)
    add_sheet_content(ws, fac, year, tp, data)

def add_sheet_content(ws, fac, year, tp, data):
    # ── Title ──
    ws.merge_cells('A1:AO1')
    t = ws['A1']
    t.value = f"Routine Immunization Monthly Monitoring  |  {fac['name']}  |  Year: {year}"
    t.font = Font(name='Arial', bold=True, size=12, color='FFFFFF')
    t.fill = HDR_FILL
    t.alignment = CTR

    # ── Facility info row ──
    ws.merge_cells('A2:D2')
    ws['A2'].value = f"Governorate: {fac['governorate']}"
    ws['A2'].font = BLD_FONT
    ws['A2'].fill = HDR2_FILL
    ws['A2'].font = WHT_FONT
    ws['A2'].alignment = CTR
    ws.merge_cells('E2:H2')
    ws['E2'].value = f"Provider: {fac['provider']}"
    ws['E2'].font = WHT_FONT
    ws['E2'].fill = HDR2_FILL
    ws['E2'].alignment = CTR
    ws.merge_cells('I2:L2')
    ws['I2'].value = f"Type: {fac['type']}"
    ws['I2'].font = WHT_FONT
    ws['I2'].fill = HDR2_FILL
    ws['I2'].alignment = CTR

    # ── Target Population Row ──
    catchment = tp['catchment_pop'] if tp else 0
    si        = tp['si_percent']    if tp else 3.2
    target_u1 = int(catchment * si / 100)

    ws.merge_cells('A3:C3')
    ws['A3'].value = 'Catchment Population'
    ws['A3'].font = BLD_FONT; ws['A3'].fill = HDR2_FILL
    ws['A3'].font = WHT_FONT; ws['A3'].alignment = CTR

    ws['D3'].value = catchment
    ws['D3'].font = BLD_FONT; ws['D3'].alignment = CTR; ws['D3'].border = BORDER
    ws['D3'].number_format = '#,##0'

    ws.merge_cells('E3:G3')
    ws['E3'].value = '% Surviving Infants (SI)'
    ws['E3'].font = WHT_FONT; ws['E3'].fill = HDR2_FILL; ws['E3'].alignment = CTR

    ws['H3'].value = si / 100
    ws['H3'].number_format = '0.0%'
    ws['H3'].font = BLD_FONT; ws['H3'].alignment = CTR; ws['H3'].border = BORDER

    ws.merge_cells('I3:K3')
    ws['I3'].value = 'Target (Under 1 yr)'
    ws['I3'].font = WHT_FONT; ws['I3'].fill = HDR2_FILL; ws['I3'].alignment = CTR

    ws['L3'].value = f'=D3*H3'
    ws['L3'].number_format = '#,##0'
    ws['L3'].font = BLD_FONT; ws['L3'].alignment = CTR; ws['L3'].border = BORDER

    # ── Column Headers ──
    # Row 4: Antigen | Jan (Tot, Cum) | Feb (Tot, Cum) ... | Dec (Tot,Cum) | Annual | Coverage%
    ROW_HDR = 4
    ws.merge_cells(f'A{ROW_HDR}:A{ROW_HDR+1}')
    cell_style(ws, ROW_HDR, 1, 'Antigen / Vaccine', HDR_FILL, WHT_FONT)

    COL_START = 2
    month_col = {}  # month -> (tot_col, cum_col)
    for i, m in enumerate(MONTHS):
        tc = COL_START + i * 2
        cc = tc + 1
        month_col[i+1] = (tc, cc)
        ws.merge_cells(f'{get_column_letter(tc)}{ROW_HDR}:{get_column_letter(cc)}{ROW_HDR}')
        cell_style(ws, ROW_HDR, tc, m, HDR_FILL, WHT_FONT)
        cell_style(ws, ROW_HDR+1, tc, 'Tot.', HDR2_FILL, WHT_FONT)
        cell_style(ws, ROW_HDR+1, cc, 'Cum.', HDR2_FILL, WHT_FONT)

    ANN_COL = COL_START + 24      # Annual total
    COV_COL = ANN_COL + 1         # Coverage %
    DO_COL  = COV_COL + 1         # Dropout %

    ws.merge_cells(f'{get_column_letter(ANN_COL)}{ROW_HDR}:{get_column_letter(ANN_COL)}{ROW_HDR+1}')
    cell_style(ws, ROW_HDR, ANN_COL, 'Annual Total', HDR_FILL, WHT_FONT)
    ws.merge_cells(f'{get_column_letter(COV_COL)}{ROW_HDR}:{get_column_letter(COV_COL)}{ROW_HDR+1}')
    cell_style(ws, ROW_HDR, COV_COL, 'Coverage %', HDR_FILL, WHT_FONT)

    # ── Data Rows ──
    DATA_START = ROW_HDR + 2
    antigen_rows = {}   # key -> excel row

    for ag_i, ag in enumerate(ANTIGENS):
        r = DATA_START + ag_i
        antigen_rows[ag['key']] = r
        fill = BLUE_FILL if ag_i % 2 == 0 else YELL_FILL

        cell_style(ws, r, 1, ag['label'], fill, BLD_FONT, Alignment(horizontal='left', vertical='center'))

        for m in range(1, 13):
            tc, cc = month_col[m]
            # None for unset months → blank cell (not 0)
            val = data.get((ag['key'], m), None)
            cell_style(ws, r, tc, val, fill, NRM_FONT, number_format='#,##0')

            # Cumulative stops at last entered month (ISBLANK check)
            mc_ref   = f'{get_column_letter(tc)}{r}'
            if m == 1:
                cum_formula = f'=IF(ISBLANK({mc_ref}),"",{mc_ref})'
            else:
                prev_cc_ref = f'{get_column_letter(month_col[m-1][1])}{r}'
                cum_formula = (
                    f'=IF(ISBLANK({mc_ref}),"",IF({prev_cc_ref}="",'
                    f'{mc_ref},{prev_cc_ref}+{mc_ref}))'
                )
            cell_style(ws, r, cc, cum_formula, fill, Font(name='Arial', size=9, italic=True),
                       number_format='#,##0')

        # Annual = SUM of monthly cols only (SUM ignores blank cells)
        tot_sum = 'SUM(' + ','.join(get_column_letter(month_col[m][0])+str(r) for m in range(1,13)) + ')'
        cell_style(ws, r, ANN_COL, f'={tot_sum}', fill, BLD_FONT, number_format='#,##0')

        # Coverage % = Annual / Target (blank if no data)
        ann_ref = f'{get_column_letter(ANN_COL)}{r}'
        cell_style(ws, r, COV_COL,
                   f'=IF(OR($L$3=0,{tot_sum}=0),"",{ann_ref}/$L$3)',
                   fill, Font(name='Arial', size=9, color='000000'), number_format='0.0%')

    # ── Drop-out Rows ──
    DO_START = DATA_START + len(ANTIGENS) + 1
    do_header_row = DO_START - 1
    ws.merge_cells(f'A{do_header_row}:{get_column_letter(DO_COL)}{do_header_row}')
    cell_style(ws, do_header_row, 1, 'Drop-out Rates', HDR_FILL, WHT_FONT,
               Alignment(horizontal='left', vertical='center'))

    for di, (label, num_key, den_key) in enumerate(DROPOUTS):
        r = DO_START + di
        fill = BLUE_FILL if di % 2 == 0 else YELL_FILL
        cell_style(ws, r, 1, label, fill, BLD_FONT,
                   Alignment(horizontal='left', vertical='center'))

        num_r = antigen_rows.get(num_key)
        den_r = antigen_rows.get(den_key)
        ann_num = f'{get_column_letter(ANN_COL)}{num_r}'
        ann_den = f'{get_column_letter(ANN_COL)}{den_r}'

        # Monthly dropout — blank if cumulative not yet reached that month
        for m in range(1, 13):
            tc, cc = month_col[m]
            num_cell = f'{get_column_letter(month_col[m][1])}{num_r}'
            den_cell = f'{get_column_letter(month_col[m][1])}{den_r}'
            # Show dropout only if num cumulative is present and > 0
            formula_do = (
                f'=IF(OR({num_cell}="",{num_cell}=0),"",IF({den_cell}="",1,'
                f'({num_cell}-{den_cell})/{num_cell}))'
            )
            do_cell = ws.cell(row=r, column=tc, value=formula_do)
            do_cell.number_format = '0.0%'
            do_cell.fill = fill
            do_cell.font = NRM_FONT
            do_cell.border = BORDER
            ws.merge_cells(f'{get_column_letter(tc)}{r}:{get_column_letter(cc)}{r}')

        # Annual dropout — based on SUM totals (not Dec cumul)
        ann_num_sum = 'SUM(' + ','.join(get_column_letter(month_col[m][0])+str(num_r) for m in range(1,13)) + ')'
        ann_den_sum = 'SUM(' + ','.join(get_column_letter(month_col[m][0])+str(den_r) for m in range(1,13)) + ')'
        do_ann = f'=IF({ann_num_sum}=0,"",({ann_num_sum}-{ann_den_sum})/{ann_num_sum})'
        ann_cell = ws.cell(row=r, column=ANN_COL, value=do_ann)
        ann_cell.number_format = '0.0%'
        ann_cell.fill = fill; ann_cell.font = BLD_FONT; ann_cell.border = BORDER

        # Conditional note
        note_cell = ws.cell(row=r, column=COV_COL,
                            value=f'=IF({get_column_letter(ANN_COL)}{r}>0.1,"HIGH ⚠","OK ✓")')
        note_cell.fill = fill; note_cell.font = BLD_FONT
        note_cell.alignment = CTR; note_cell.border = BORDER

    # ── Achievement Legend ──
    LEG_ROW = DO_START + len(DROPOUTS) + 2
    ws.merge_cells(f'A{LEG_ROW}:C{LEG_ROW}')
    cell_style(ws, LEG_ROW, 1, 'Coverage Achievement Legend:', None, BLD_FONT,
               Alignment(horizontal='left'))
    thresholds = [('≥100%', '00B050'), ('≥75%','92D050'), ('≥50%','FFEB84'), ('≥45%','FF9900'), ('<45%','FF0000')]
    for ti, (lbl, color) in enumerate(thresholds):
        c = 4 + ti
        cell = ws.cell(row=LEG_ROW, column=c, value=lbl)
        cell.fill = PatternFill('solid', fgColor=color)
        cell.font = Font(name='Arial', size=9, bold=True)
        cell.alignment = CTR
        cell.border = BORDER

    # ── Formula Notes ──
    NOTE_ROW = LEG_ROW + 2
    ws.merge_cells(f'A{NOTE_ROW}:H{NOTE_ROW}')
    ws[f'A{NOTE_ROW}'].value = 'Formulas:  Drop-out # = First Dose − Last Dose  |  Drop-out % = Drop-out # ÷ First Dose × 100  |  Coverage % = Annual Doses ÷ Target Population × 100'
    ws[f'A{NOTE_ROW}'].font = Font(name='Arial', size=8, italic=True, color='444444')

    # ── Column widths ──
    ws.column_dimensions['A'].width = 22
    for m in range(1, 13):
        tc, cc = month_col[m]
        ws.column_dimensions[get_column_letter(tc)].width = 7
        ws.column_dimensions[get_column_letter(cc)].width = 7
    ws.column_dimensions[get_column_letter(ANN_COL)].width = 12
    ws.column_dimensions[get_column_letter(COV_COL)].width = 12
    ws.row_dimensions[1].height = 22
    ws.freeze_panes = f'B{DATA_START}'

@app.route('/stock/<int:fid>', methods=['GET', 'POST'])
def stock(fid):
    year  = request.args.get('year',  2026, type=int)
    month = request.args.get('month', 1,    type=int)
    conn  = get_db()
    fac   = conn.execute('SELECT * FROM facilities WHERE id=?', (fid,)).fetchone()
    if not fac:
        conn.close()
        return redirect(url_for('index'))

    if request.method == 'POST':
        year  = int(request.form.get('year',  2026))
        month = int(request.form.get('month', 1))
        for v in STOCK_VACCINES:
            k = v['key']
            def fi(field):
                s = request.form.get(f'{k}_{field}', '').strip()
                return int(s) if s.isdigit() else None
            opening = fi('opening'); received = fi('received')
            administered = fi('administered'); wastage = fi('wastage')
            # Only save if at least one field filled
            if any(x is not None for x in [opening, received, administered, wastage]):
                conn.execute('''INSERT INTO stock_log
                    (facility_id,year,month,antigen,opening,received,administered,wastage)
                    VALUES (?,?,?,?,?,?,?,?)
                    ON CONFLICT(facility_id,year,month,antigen) DO UPDATE SET
                    opening=excluded.opening, received=excluded.received,
                    administered=excluded.administered, wastage=excluded.wastage''',
                    (fid, year, month, k, opening, received, administered, wastage))
            else:
                conn.execute('DELETE FROM stock_log WHERE facility_id=? AND year=? AND month=? AND antigen=?',
                             (fid, year, month, k))
        conn.commit()
        conn.close()
        return redirect(url_for('stock', fid=fid, year=year, month=month, saved=1))

    rows = conn.execute(
        'SELECT antigen,opening,received,administered,wastage FROM stock_log WHERE facility_id=? AND year=? AND month=?',
        (fid, year, month)).fetchall()
    data = {r['antigen']: r for r in rows}

    # Auto-fill opening from previous month's closing if not yet set
    prev_month = month - 1
    prev_year  = year
    if prev_month == 0:
        prev_month = 12
        prev_year  = year - 1
    prev_rows = conn.execute(
        'SELECT antigen,opening,received,administered,wastage FROM stock_log WHERE facility_id=? AND year=? AND month=?',
        (fid, prev_year, prev_month)).fetchall()
    prev_closing = {}
    for r in prev_rows:
        o = r['opening'] or 0; rc = r['received'] or 0
        a = r['administered'] or 0; w = r['wastage'] or 0
        prev_closing[r['antigen']] = o + rc - a - w

    conn.close()
    saved = request.args.get('saved', 0)
    return render_template('stock.html', fac=fac, year=year, month=month,
                           vaccines=STOCK_VACCINES, months=MONTHS, months_ar=MONTHS_AR,
                           data=data, prev_closing=prev_closing, saved=saved)


@app.route('/export_stock/<int:fid>')
def export_stock(fid):
    year = request.args.get('year', 2026, type=int)
    conn = get_db()
    fac  = conn.execute('SELECT * FROM facilities WHERE id=?', (fid,)).fetchone()
    rows = conn.execute(
        'SELECT month,antigen,opening,received,administered,wastage FROM stock_log WHERE facility_id=? AND year=? ORDER BY antigen,month',
        (fid, year)).fetchall()
    conn.close()

    # Build {antigen: {month: row}}
    sdata = {}
    for r in rows:
        sdata.setdefault(r['antigen'], {})[r['month']] = r

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Stock Log'

    # Title
    ws.merge_cells('A1:R1')
    ws['A1'].value = f'Stock Log – {fac["name"]} – {year}'
    ws['A1'].font  = Font(name='Arial', bold=True, size=12, color='FFFFFF')
    ws['A1'].fill  = HDR_FILL
    ws['A1'].alignment = CTR

    # Headers row 2: Vaccine | Jan (Open,Recv,Admin,Waste,Close) | Feb ... | Dec
    ws.cell(row=2, column=1, value='Vaccine').font = WHT_FONT
    ws.cell(row=2, column=1).fill = HDR_FILL
    ws.cell(row=2, column=1).alignment = CTR

    COL = 2
    month_cols = {}
    for mi, m in enumerate(MONTHS):
        ws.merge_cells(start_row=2, start_column=COL, end_row=2, end_column=COL+4)
        c = ws.cell(row=2, column=COL, value=m)
        c.font = WHT_FONT; c.fill = HDR2_FILL if mi%2 else HDR_FILL; c.alignment = CTR
        month_cols[mi+1] = COL
        for j, sub in enumerate(['Open','Recv','Admin','Waste','Close']):
            sc = ws.cell(row=3, column=COL+j, value=sub)
            sc.font = WHT_FONT; sc.fill = HDR2_FILL; sc.alignment = CTR; sc.border = BORDER
        COL += 5

    ws.row_dimensions[2].height = 18
    ws.row_dimensions[3].height = 14

    # Data rows
    for vi, v in enumerate(STOCK_VACCINES):
        r = 4 + vi
        fill = BLUE_FILL if vi%2==0 else YELL_FILL
        ws.cell(row=r, column=1, value=v['label']).font  = BLD_FONT
        ws.cell(row=r, column=1).fill = fill
        ws.cell(row=r, column=1).border = BORDER

        prev_close = None
        for mi in range(1, 13):
            sc = month_cols[mi]
            md = sdata.get(v['key'], {}).get(mi)
            o  = md['opening']      if md and md['opening']      is not None else (prev_close if prev_close is not None else None)
            rc = md['received']     if md and md['received']     is not None else None
            a  = md['administered'] if md and md['administered'] is not None else None
            w  = md['wastage']      if md and md['wastage']      is not None else None
            cl = (o or 0) + (rc or 0) - (a or 0) - (w or 0) if any(x is not None for x in [o,rc,a,w]) else None
            prev_close = cl

            for j, val in enumerate([o, rc, a, w, cl]):
                cell = ws.cell(row=r, column=sc+j, value=val)
                cell.fill = fill; cell.border = BORDER; cell.alignment = CTR
                if val is not None:
                    cell.number_format = '#,##0'
                if j == 4 and cl is not None:  # closing col
                    if cl < 0:   cell.font = Font(name='Arial', size=9, bold=True, color='FF0000')
                    elif cl == 0: cell.font = Font(name='Arial', size=9, bold=True, color='FF9900')
                    else:         cell.font = Font(name='Arial', size=9, bold=True, color='00B050')
                else:
                    cell.font = NRM_FONT

        ws.row_dimensions[r].height = 16

    ws.column_dimensions['A'].width = 20
    for mi in range(1, 13):
        for j in range(5):
            ws.column_dimensions[get_column_letter(month_cols[mi]+j)].width = 8
    ws.freeze_panes = 'B4'

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    safe = str(fac['name']).replace(' ','_').replace('/','–')
    return send_file(buf, as_attachment=True,
                     download_name=f'StockLog_{safe}_{year}.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/stock_summary/<int:fid>')
def stock_summary(fid):
    year = request.args.get('year', 2026, type=int)
    conn = get_db()
    fac  = conn.execute('SELECT * FROM facilities WHERE id=?', (fid,)).fetchone()
    if not fac:
        conn.close(); return redirect(url_for('index'))
    rows = conn.execute(
        'SELECT month,antigen,opening,received,administered,wastage FROM stock_log WHERE facility_id=? AND year=? ORDER BY antigen,month',
        (fid, year)).fetchall()
    conn.close()
    sdata = {}
    for r in rows:
        sdata.setdefault(r['antigen'], {})[r['month']] = r
    return render_template('stock_summary.html', fac=fac, year=year,
                           vaccines=STOCK_VACCINES, months=MONTHS, months_ar=MONTHS_AR,
                           sdata=sdata)


if __name__ == '__main__':
    init_db()
    app.run(debug=True, port=5000)
