import os
import sqlite3
from datetime import datetime
from flask import Flask, render_template, request, jsonify, send_file
import openpyxl

app = Flask(__name__)
DB_NAME = "trolley_data.db"
EXCEL_NAME = "รายงานตรวจเช็กรถเข็น.xlsx"

def init_db():
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS inspections (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                cart_id TEXT NOT NULL,
                inspector_id TEXT NOT NULL,
                wheel_check TEXT NOT NULL,
                bolt_check TEXT NOT NULL,
                edge_check TEXT NOT NULL,
                frame_check TEXT NOT NULL,
                overall_status TEXT NOT NULL,
                remark TEXT,
                inspect_time TEXT NOT NULL
            )
        """)
        conn.commit()

def append_to_excel(row_data):
    if not os.path.exists(EXCEL_NAME):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "ประวัติการตรวจเช็ก"
        ws.append(["วัน-เวลา", "รหัสรถเข็น", "รหัสผู้ตรวจ", "ล้อ", "น็อต", "ลบคม", "โครงสร้าง", "ผลรวม", "หมายเหตุ"])
    else:
        wb = openpyxl.load_workbook(EXCEL_NAME)
        ws = wb.active
    ws.append(row_data)
    wb.save(EXCEL_NAME)

init_db()

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/submit", methods=["POST"])
def submit_data():
    data = request.json
    cart_id = data.get("cart_id", "").strip().upper()
    inspector_id = data.get("inspector_id", "").strip().upper()
    wheel = data.get("wheel")
    bolt = data.get("bolt")
    edge = data.get("edge")
    frame = data.get("frame")
    remark = data.get("remark", "").strip() or "-"

    is_pass = all(item == "PASS" for item in [wheel, bolt, edge, frame])
    status = "PASS" if is_pass else "REPAIR"
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO inspections (cart_id, inspector_id, wheel_check, bolt_check, edge_check, frame_check, overall_status, remark, inspect_time)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (cart_id, inspector_id, wheel, bolt, edge, frame, status, remark, now))
        conn.commit()

    append_to_excel([now, cart_id, inspector_id, wheel, bolt, edge, frame, status, remark])
    return jsonify({"success": True, "status": status, "time": now})

@app.route("/api/history", methods=["GET"])
def get_history():
    with sqlite3.connect(DB_NAME) as conn:
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM inspections ORDER BY id DESC LIMIT 50")
        rows = cursor.fetchall()
        return jsonify([dict(row) for row in rows])

@app.route("/download-excel")
def download_excel():
    if os.path.exists(EXCEL_NAME):
        return send_file(EXCEL_NAME, as_attachment=True)
    return "ยังไม่มีข้อมูล", 404

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
