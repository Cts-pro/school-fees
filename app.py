import os
from flask import Flask, render_template, request, redirect, url_for, session, flash
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
        headers = ["Record ID", "Student Name", "Fee Amount ($)", "Status", "Is Locked"]
        ws.append(headers)
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal="center")
        samples = [
            (1, "John Doe", 500.00, "Invoiced", 1),
            (2, "Alice Smith", 500.00, "Pending", 0)
        ]
        for r_idx, row in enumerate(samples, 2):
            ws.append(row)
            current_fill = fill_locked if row[4] == 1 else fill_pending
            current_font = font_locked if row[4] == 1 else font_body
            for c_idx in range(1, 6):
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.font = current_font
                cell.fill = current_fill
                if c_idx == 3:
                    cell.number_format = '$#,##0.00'
                    
        ws2 = wb.create_sheet(title="Grade 6")
        ws2.views.sheetView[0].showGridLines = True
        ws2.append(headers)
        for col, h in enumerate(headers, 1):
            ws2.cell(row=1, column=col).font = font_header
            ws2.cell(row=1, column=col).fill = fill_header
        ws2.append((3, "Jane Smith", 600.00, "Pending", 0))
        for c_idx in range(1, 6):
            ws2.cell(row=2, column=c_idx).fill = fill_pending
            ws2.cell(row=2, column=c_idx).font = font_body
        wb.save(EXCEL_FILE)

def get_all_records_from_excel():
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
                "class": sheet_name,
                "amount": ws.cell(row=row, column=3).value,
                "status": ws.cell(row=row, column=4).value,
                "is_locked": ws.cell(row=row, column=5).value
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

@app.route('/update_fee/<sheet_name>/<int:row_id>', methods=['POST'])
def update_fee(sheet_name, row_id):
    if 'user_role' not in session:
        return redirect(url_for('login'))
    new_amount = float(request.form['fee_amount'])
    wb = openpyxl.load_workbook(EXCEL_FILE)
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for row in range(2, ws.max_row + 1):
            if ws.cell(row=row, column=1).value == row_id:
                if ws.cell(row=row, column=5).value == 1:
                    flash('CRITICAL EXCEPTION: Edit denied. Row is permanently locked.')
                else:
                    ws.cell(row=row, column=3).value = new_amount
                    wb.save(EXCEL_FILE)
                    flash(f'Successfully updated record {row_id} inside sheet tab "{sheet_name}".')
                break
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
                ws.cell(row=row, column=4).value = "Invoiced"
                ws.cell(row=row, column=5).value = 1
                for col in range(1, 6):
                    ws.cell(row=row, column=col).fill = fill_locked
                    ws.cell(row=row, column=col).font = font_locked
                record_data = {
                    "id": row_id,
                    "name": ws.cell(row=row, column=2).value,
                    "class": sheet_name,
                    "amount": ws.cell(row=row, column=3).value,
                    "status": "Invoiced"
                }
                wb.save(EXCEL_FILE)
                break
    return f"""
    <html>
    <head><title>Invoice - {record_data['name']}</title><link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css"></head>
    <body class="p-5">
        <div class="card p-4 mx-auto" style="max-width: 600px; border: 2px dashed #333;">
            <h3 class="text-center text-uppercase fw-bold">Official School Receipt</h3>
            <hr>
            <p><strong>System Tracking ID:</strong> REC-{record_data['id']}</p>
            <p><strong>Class Sheet Reference:</strong> {record_data['class']}</p>
            <p><strong>Student Target:</strong> {record_data['name']}</p>
            <p><strong>Status Rule:</strong> <span class="badge bg-danger">Archived & Immutable</span></p>
            <hr>
            <h4>Amount Authenticated: ${record_data['amount']:.2f}</h4>
            <hr>
            <button class="btn btn-dark d-print-none w-100" onclick="window.print()">Print Output Document</button>
            <a href="/" class="btn btn-outline-secondary d-print-none w-100 mt-2">Return to Dashboard</a>
        </div>
    </body>
    </html>
    """

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

init_excel()

if __name__ == '__main__':
    app.run(debug=False)
