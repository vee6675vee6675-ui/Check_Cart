from flask import Flask, render_template, request, jsonify, send_file
import sqlite3
import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import os

app = Flask(__name__)
DB_NAME = "cart_inspection.db"

CHECK_ITEMS = [
    ("c_1_1", "1.1 สภาพผิวยางล้อ"),
    ("c_1_2", "1.2 การหมุนของลูกปืนล้อ"),
    ("c_1_3", "1.3 การเลี้ยว/คอล้อ (Swivel)"),
    ("c_1_4", "1.4 สิ่งแปลกปลอม"),
    ("c_1_5", "1.5 ระบบเบรกล้อ"),
    ("c_2_1", "2.1 น็อตยึดแป้นล้อ"),
    ("c_2_2", "2.2 น็อตและสลักจุดพับ/มือจับ"),
    ("c_3_1", "3.1 ความคมของขอบเหล็ก"),
    ("c_3_2", "3.2 ความสมบูรณ์ของรอยเชื่อม"),
    ("c_3_3", "3.3 การบิดงอ/เสียรูปทรง"),
    ("c_3_4", "3.4 สนิมและสภาพสี"),
    ("c_4_1", "4.1 ความแข็งแรงของมือจับ"),
    ("c_4_2", "4.2 ระบบพับ/ล็อกมือจับ"),
    ("c_4_3", "4.3 ป้ายรหัสและบาร์โค้ด")
]

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    columns_sql = ", ".join([f"{key} TEXT" for key, _ in CHECK_ITEMS])
    c.execute(f'''
        CREATE TABLE IF NOT EXISTS inspections (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT,
            cart_id TEXT,
            inspector_id TEXT,
            {columns_sql},
            overall_status TEXT,
            remark TEXT
        )
    ''')
    conn.commit()
    conn.close()

init_db()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/submit', methods=['POST'])
def submit():
    try:
        data = request.json
        cart_id = data.get('cart_id', '').strip()
        inspector_id = data.get('inspector_id', '').strip()
        remark = data.get('remark', '').strip()
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        if not cart_id or not inspector_id:
            return jsonify({"success": False, "error": "กรุณากรอกรหัสรถเข็นและรหัสผู้ตรวจ"}), 400

        item_values = []
        has_fail = False

        for key, _ in CHECK_ITEMS:
            val = data.get(key, 'PASS')
            if val == 'FAIL':
                has_fail = True
            item_values.append(val)

        overall_status = "FAIL" if has_fail else "PASS"

        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        col_names = ", ".join(["timestamp", "cart_id", "inspector_id"] + [key for key, _ in CHECK_ITEMS] + ["overall_status", "remark"])
        placeholders = ", ".join(["?"] * (len(CHECK_ITEMS) + 5))
        
        c.execute(f'''
            INSERT INTO inspections ({col_names})
            VALUES ({placeholders})
        ''', [timestamp, cart_id, inspector_id] + item_values + [overall_status, remark])
        
        conn.commit()
        conn.close()

        return jsonify({"success": True, "status": overall_status})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/download-excel')
def download_excel():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT * FROM inspections ORDER BY id DESC")
    rows = c.fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "บันทึกตรวจสภาพรถเข็น"

    headers = ["ลำดับ", "วัน-เวลาที่ตรวจ", "รหัสรถเข็น", "รหัสผู้ตรวจ"] + [name for _, name in CHECK_ITEMS] + ["ผลการตรวจรวม", "หมายเหตุ"]
    ws.append(headers)

    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    pass_fill = PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid")
    fail_fill = PatternFill(start_color="FFE4E6", end_color="FFE4E6", fill_type="solid")
    na_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    thin_border = Border(left=Side(style='thin', color='CBD5E1'), right=Side(style='thin', color='CBD5E1'),
                         top=Side(style='thin', color='CBD5E1'), bottom=Side(style='thin', color='CBD5E1'))

    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for r_idx, row in enumerate(rows, start=2):
        ws.append(list(row))
        for c_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.border = thin_border
            val = str(cell.value)
            if val == "PASS":
                cell.fill = pass_fill
                cell.alignment = Alignment(horizontal="center")
            elif val == "FAIL":
                cell.fill = fail_fill
                cell.font = Font(color="DC2626", bold=True)
                cell.alignment = Alignment(horizontal="center")
            elif val == "N/A":
                cell.fill = na_fill
                cell.alignment = Alignment(horizontal="center")

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    excel_file = "cart_inspection_report.xlsx"
    wb.save(excel_file)
    return send_file(excel_file, as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
