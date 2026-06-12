import os
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

app = Flask(__name__)
app.secret_key = 'super_secret_permanent_key_123'
EXCEL_FILE = 'school_fees_by_class.xlsx'

font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
font_body = Font(name="Calibri", size=11)
font_locked = Font(name="Calibri", size=11, italic=True, color="595959")
fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
fill_locked = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
fill_pending = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

def init_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Grade 5"
        ws.views.sheetView[0].showGridLines = True
        headers = ["Record ID", "Student Name", "Category", "Particulars List", "Quantities List", "Amounts List", "Total Amount", "Status", "Is Locked"]
        ws.append(headers)
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal="center")
        wb.save(EXCEL_FILE)

def get_all_records_from_excel():
    if not os.path.exists(EXCEL_FILE):
        init_excel()
    wb = openpyxl.load_workbook(EXCEL_FILE)
    all_records = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for row in range(2, ws.max_row + 1):
            if ws.cell(row=row, column=1).value is None:
                continue
            all_records.append({
                "id": ws.cell(row=row, column=1).value,
                "name": ws.cell(row=row, column=2).value,
                "category": str(ws.cell(row=row, column=3).value or "Fees").strip(),
                "particulars": ws.cell(row=row, column=4).value or "",
                "quantities": ws.cell(row=row, column=5).value or "",
                "amounts": ws.cell(row=row, column=6).value or "",
                "total_amount": float(ws.cell(row=row, column=7).value or 0.0),
                "class": sheet_name,
                "status": ws.cell(row=row, column=8).value,
                "is_locked": int(ws.cell(row=row, column=9).value or 0)
            })
    return all_records

@app.route('/')
def index():
    if 'user_role' not in session:
        return redirect(url_for('login'))
    records = get_all_records_from_excel()
    return render_template('index.html', records=records, role=session.get('user_role'), username=session.get('username'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        if username == 'admin' and password == 'admin123':
            session['user_role'] = 'admin'
            session['username'] = 'Admin'
            return redirect(url_for('index'))
        elif username == 'staff' and password == 'staff123':
            session['user_role'] = 'user'
            session['username'] = 'Bursar Staff'
            return redirect(url_for('index'))
        else:
            flash('Invalid login credentials.')
    return render_template('index.html', login_page=True)

@app.route('/add_student', methods=['POST'])
def add_student():
    if 'user_role' not in session:
        return redirect(url_for('login'))
        
    student_name = request.form['student_name']
    student_class = request.form['student_class'].strip()
    category = request.form['category'].strip()
    particulars = request.form['particulars']
    quantities = request.form['quantities']
    amounts_str = request.form['amounts']
    
    try:
        amounts_list = [float(x.strip()) for x in amounts_str.split(',')]
        total_amount = sum(amounts_list)
    except ValueError:
        flash("Formatting Error: Ensure individual amounts are comma-separated numbers.")
        return redirect(url_for('index'))
        
    wb = openpyxl.load_workbook(EXCEL_FILE)
    if student_class not in wb.sheetnames:
        ws = wb.create_sheet(title=student_class)
        ws.views.sheetView[0].showGridLines = True
        headers = ["Record ID", "Student Name", "Category", "Particulars List", "Quantities List", "Amounts List", "Total Amount", "Status", "Is Locked"]
        ws.append(headers)
        for col, h in enumerate(headers, 1):
            ws.cell(row=1, column=col).font = font_header
            ws.cell(row=1, column=col).fill = fill_header
    else:
        ws = wb[student_class]
        
    all_records = get_all_records_from_excel()
    next_id = len(all_records) + 1
    
    ws.append([next_id, student_name, category, particulars, quantities, amounts_str, total_amount, "Pending", 0])
    
    new_row_idx = ws.max_row
    for col in range(1, 10):
        cell = ws.cell(row=new_row_idx, column=col)
        cell.fill = fill_pending
        cell.font = font_body
        if col == 7:
            cell.number_format = 'R#,##0.00'
            
    wb.save(EXCEL_FILE)
    flash(f"Record created successfully for {student_name} under {category}!")
    return redirect(url_for('index'))

@app.route('/download_database')
def download_database():
    if 'user_role' not in session:
        return redirect(url_for('login'))
    if os.path.exists(EXCEL_FILE):
        return send_file(EXCEL_FILE, as_attachment=True, download_name="vinayaka_school_database.xlsx")
    else:
        flash("Data registry document is not generated yet.")
        return redirect(url_for('index'))

@app.route('/generate_invoice/<sheet_name>/<int:row_id>')
def generate_invoice(sheet_name, row_id):
    if 'user_role' not in session:
        return redirect(url_for('login'))
    wb = openpyxl.load_workbook(EXCEL_FILE)
    record_data = {}
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for row in range(2, ws.max_row + 1):
            if ws.cell(row=row, column=1).value == row_id:
                ws.cell(row=row, column=8).value = "Invoiced"
                ws.cell(row=row, column=9).value = 1
                for col in range(1, 10):
                    ws.cell(row=row, column=col).fill = fill_locked
                    ws.cell(row=row, column=col).font = font_locked
                
                record_data = {
                    "id": row_id,
                    "name": ws.cell(row=row, column=2).value,
                    "category": str(ws.cell(row=row, column=3).value or "Fees").strip(),
                    "particulars": str(ws.cell(row=row, column=4).value).split(','),
                    "quantities": str(ws.cell(row=row, column=5).value).split(','),
                    "amounts": str(ws.cell(row=row, column=6).value).split(','),
                    "total": float(ws.cell(row=row, column=7).value or 0.0),
                    "class": sheet_name
                }
                wb.save(EXCEL_FILE)
                break

    table_rows_html = ""
    for i in range(len(record_data['particulars'])):
        part = record_data['particulars'][i].strip()
        qty = record_data['quantities'][i].strip() if i < len(record_data['quantities']) else "1"
        amt = record_data['amounts'][i].strip() if i < len(record_data['amounts']) else "0.00"
        table_rows_html += f"<tr><td class='text-center'>{i+1}</td><td>{part}</td>
