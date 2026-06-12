import os
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

app = Flask(__name__)
app.secret_key = 'super_secret_permanent_key_123'
EXCEL_FILE = '/data/school_fees_by_class.xlsx'

font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
font_body = Font(name="Calibri", size=11)
font_locked = Font(name="Calibri", size=11, italic=True, color="595959")
fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
fill_locked = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
fill_pending = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

def init_excel():
    dirname = os.path.dirname(EXCEL_FILE)
    if dirname and not os.path.exists(dirname):
        os.makedirs(dirname, exist_ok=True)
    if not os.path.exists(EXCEL_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Grade 5"
        ws.views.sheetView[0].showGridLines = True
        headers = ["Record ID", "Student Name", "Category", "Particulars List", "Quantities List", "Amounts List", "Total Amount", "Status", "Is Locked", "Invoice Number"]
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
            try:
                all_records.append({
                    "id": int(ws.cell(row=row, column=1).value),
                    "name": ws.cell(row=row, column=2).value,
                    "category": str(ws.cell(row=row, column=3).value or "Fees").strip(),
                    "particulars": str(ws.cell(row=row, column=4).value or ""),
                    "quantities": str(ws.cell(row=row, column=5).value or ""),
                    "amounts": str(ws.cell(row=row, column=6).value or ""),
                    "total_amount": float(ws.cell(row=row, column=7).value or 0.0),
                    "class": sheet_name,
                    "status": ws.cell(row=row, column=8).value,
                    "is_locked": int(ws.cell(row=row, column=9).value or 0),
                    "invoice_num": str(ws.cell(row=row, column=10).value or "DRAFT-UNASSIGNED")
                })
            except Exception:
                continue
    return all_records

def get_next_invoice_number_for_category(category):
    records = get_all_records_from_excel()
    count = 0
    for r in records:
        if r['category'] == category and r['is_locked'] == 1:
            count += 1
    prefix = {"Fees": "Tuition-", "Books": "Book-", "Uniform": "Uniform-"}
    return f"{prefix.get(category, 'INV-')}{count + 1:04d}"

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
    student_class = request.form
