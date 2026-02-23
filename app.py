import os
from datetime import datetime, date
from typing import Optional, Tuple

import pandas as pd
from flask import Flask, request, redirect, url_for, render_template_string, abort
from flask_sqlalchemy import SQLAlchemy

# ---------------------------
# App & DB Config
# ---------------------------
app = Flask(__name__)

# Render ใช้ /opt/render/project/src เป็น working dir
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "sanko_training.db")

app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{DB_PATH}"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)

# ---------------------------
# Database Model
# ---------------------------
class Employee(db.Model):
    id = db.Column(db.Integer, primary_key=True)

    # Key fields
    no = db.Column(db.Integer, nullable=True)
    em_id = db.Column(db.String(50), unique=True, nullable=False, index=True)
    id_card = db.Column(db.String(50), nullable=True)

    # Thai name split
    title_th = db.Column(db.String(30), nullable=True)
    first_name_th = db.Column(db.String(100), nullable=True)
    last_name_th = db.Column(db.String(100), nullable=True)

    # English name split
    title_en = db.Column(db.String(30), nullable=True)
    first_name_en = db.Column(db.String(100), nullable=True)
    last_name_en = db.Column(db.String(100), nullable=True)

    # Org & Job
    position = db.Column(db.String(120), nullable=True)
    section = db.Column(db.String(120), nullable=True)
    department = db.Column(db.String(120), nullable=True)

    # Dates
    start_work = db.Column(db.Date, nullable=True)
    resign_date = db.Column(db.Date, nullable=True)

    # Status & Education
    status = db.Column(db.String(50), nullable=True)
    degree = db.Column(db.String(50), nullable=True)
    major = db.Column(db.String(120), nullable=True)

    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def name_th_full(self) -> str:
        parts = [self.title_th, self.first_name_th, self.last_name_th]
        return " ".join([p for p in parts if p])

    def name_en_full(self) -> str:
        parts = [self.title_en, self.first_name_en, self.last_name_en]
        return " ".join([p for p in parts if p])


# ---------------------------
# Helpers
# ---------------------------
def ensure_db():
    with app.app_context():
        db.create_all()

def parse_date(value) -> Optional[date]:
    """Parse date from Excel cell: supports datetime/date/string formats."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None

    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    s = str(value).strip()
    if not s or s.lower() in {"nan", "none"}:
        return None

    # รองรับ dd-mm-yy / dd-mm-yyyy / yyyy-mm-dd
    # ถ้าเป็น 01-01-96 จะกลายเป็น 1996 โดย default ของ strptime %y => 1996
    for fmt in ("%d-%m-%Y", "%d-%m-%y", "%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def split_th_name(fullname: str) -> Tuple[str, str, str]:
    """
    ตัวอย่าง:
    - "นาย สมชาย ใจดี"
    - "นางสาว สุจิตรา พิศาล"
    - "Mr. สมชาย ใจดี" (ถ้ามีหลุดมา ก็ยังแยกได้)
    """
    if not fullname:
        return "", "", ""
    s = str(fullname).strip()
    s = " ".join(s.split())
    parts = s.split()

    if len(parts) == 1:
        return "", parts[0], ""
    if len(parts) == 2:
        return "", parts[0], parts[1]

    title_candidates = {"นาย", "นาง", "นางสาว", "Mr.", "Mrs.", "Ms.", "Miss", "คุณ"}
    if parts[0] in title_candidates:
        title = parts[0]
        first = parts[1]
        last = " ".join(parts[2:])
        return title, first, last

    # ไม่มีคำนำหน้า
    title = ""
    first = parts[0]
    last = " ".join(parts[1:])
    return title, first, last


def split_en_name(fullname: str) -> Tuple[str, str, str]:
    """
    ตัวอย่าง:
    - "Mr.Masami Katsumoto"
    - "Mr. Masami Katsumoto"
    - "Masami Katsumoto"
    """
    if not fullname:
        return "", "", ""
    s = str(fullname).strip()
    s = s.replace("Mr.", "Mr. ").replace("Mrs.", "Mrs. ").replace("Ms.", "Ms. ").replace("Miss", "Miss ")
    s = " ".join(s.split())
    parts = s.split()

    if len(parts) == 1:
        return "", parts[0], ""
    if len(parts) == 2:
        return "", parts[0], parts[1]

    title_candidates = {"Mr.", "Mrs.", "Ms.", "Miss"}
    if parts[0] in title_candidates:
        title = parts[0]
        first = parts[1]
        last = " ".join(parts[2:])
        return title, first, last

    title = ""
    first = parts[0]
    last = " ".join(parts[1:])
    return title, first, last


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normalize columns to match your header names.
    Expected headers:
    No., Em. ID, ID Card, Name-TH, Name-EN, Position, Section, Department,
    Start work, Resign, Status, Degree, Major
    """
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    # แก้ชื่อที่พบบ่อยให้เป็นมาตรฐาน
    rename_map = {
        "No": "No.",
        "Emp. ID": "Em. ID",
        "Employee ID": "Em. ID",
        "IDCard": "ID Card",
        "Name TH": "Name-TH",
        "Name EN": "Name-EN",
        "Start work": "Start work",
        "Start_work": "Start work",
        "Resign": "Resign",
        "Resign date": "Resign",
    }
    df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns}, inplace=True)

    return df


# ---------------------------
# Simple HTML Templates (inline)
# ---------------------------
BASE_HTML = """
<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>{{ title }}</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 24px; }
    a { color: #0b57d0; }
    .top { display:flex; align-items:center; gap:12px; }
    .badge { background:#f1f3f4; padding:4px 10px; border-radius:999px; font-size:12px; }
    table { border-collapse: collapse; width: 100%; margin-top: 12px; }
    th, td { border: 1px solid #ddd; padding: 8px; text-align:left; font-size:14px; }
    th { background: #f7f7f7; }
    input, select { padding: 8px; width: 100%; box-sizing:border-box; }
    .row { display:grid; grid-template-columns: 1fr 1fr; gap:12px; }
    .row3 { display:grid; grid-template-columns: 1fr 1fr 1fr; gap:12px; }
    .btn { padding: 10px 14px; border: 1px solid #ddd; background:#fff; cursor:pointer; border-radius:8px; }
    .btn-primary { background:#111827; color:#fff; border-color:#111827; }
    .btn-danger { background:#b91c1c; color:#fff; border-color:#b91c1c; }
    .actions { display:flex; gap:8px; }
    .muted { color:#666; font-size:13px; }
    .card { border:1px solid #e5e7eb; border-radius:12px; padding:14px; margin-top:12px; }
    .hr { height:1px; background:#eee; margin:16px 0; }
  </style>
</head>
<body>
  <div class="top">
    <h1 style="margin:0;">{{ header }}</h1>
    <span class="badge">SANKO Training System 🚀</span>
  </div>
  <div class="muted" style="margin-top:6px;">{{ subtitle }}</div>
  <div class="hr"></div>
  {{ body|safe }}
</body>
</html>
"""

HOME_BODY = """
<h2>เมนูหลัก</h2>
<ul>
  <li><a href="/employees">Employee Report</a></li>
  <li><a href="/employees/import">Import Employee Data (Excel)</a></li>
  <li><a href="/training">Training Matrix</a></li>
</ul>
"""

TRAINING_BODY = """
<h2>Training Matrix</h2>
<p class="muted">หน้านี้เป็นตัวอย่าง (ต่อไปเราจะทำ Matrix จริงจากหลักสูตร/ตำแหน่ง)</p>
<a href="/">กลับหน้าเมนูหลัก</a>
"""

EMP_LIST_BODY = """
<div class="actions">
  <a class="btn btn-primary" href="/employees/new">+ เพิ่มพนักงาน</a>
  <a class="btn" href="/employees/import">นำเข้า Excel</a>
  <a class="btn" href="/">กลับหน้าเมนูหลัก</a>
</div>

<form method="get" style="margin-top:12px;">
  <div class="row3">
    <div>
      <label>ค้นหา (Em. ID / ชื่อ / แผนก)</label>
      <input name="q" value="{{ q }}" placeholder="เช่น M0096001 / สมชาย / Production">
    </div>
    <div>
      <label>Department</label>
      <input name="dept" value="{{ dept }}" placeholder="เช่น Production">
    </div>
    <div style="display:flex; align-items:end; gap:8px;">
      <button class="btn btn-primary" type="submit">ค้นหา</button>
      <a class="btn" href="/employees">ล้าง</a>
    </div>
  </div>
</form>

<table>
  <tr>
    <th>No.</th>
    <th>Em. ID</th>
    <th>ID Card</th>
    <th>Name-TH</th>
    <th>Name-EN</th>
    <th>Position</th>
    <th>Section</th>
    <th>Department</th>
    <th>Start work</th>
    <th>Resign</th>
    <th>Status</th>
    <th>Degree</th>
    <th>Major</th>
    <th>Action</th>
  </tr>

  {% for e in employees %}
  <tr>
    <td>{{ e.no or "" }}</td>
    <td><b>{{ e.em_id }}</b></td>
    <td>{{ e.id_card or "" }}</td>
    <td>{{ e.name_th_full() }}</td>
    <td>{{ e.name_en_full() }}</td>
    <td>{{ e.position or "" }}</td>
    <td>{{ e.section or "" }}</td>
    <td>{{ e.department or "" }}</td>
    <td>{{ e.start_work or "" }}</td>
    <td>{{ e.resign_date or "" }}</td>
    <td>{{ e.status or "" }}</td>
    <td>{{ e.degree or "" }}</td>
    <td>{{ e.major or "" }}</td>
    <td>
      <div class="actions">
        <a class="btn" href="/employees/{{ e.id }}/edit">แก้ไข</a>
        <form method="post" action="/employees/{{ e.id }}/delete" onsubmit="return confirm('ลบพนักงาน {{e.em_id}} ?');">
          <button class="btn btn-danger" type="submit">ลบ</button>
        </form>
      </div>
    </td>
  </tr>
  {% endfor %}
</table>

<p class="muted">ทั้งหมด {{ total }} รายการ</p>
"""

EMP_FORM_BODY = """
<div class="actions">
  <a class="btn" href="/employees">← กลับ</a>
</div>

<form method="post" class="card">
  <div class="row3">
    <div>
      <label>No.</label>
      <input name="no" value="{{ e.no or '' }}" />
    </div>
    <div>
      <label>Em. ID (ห้ามซ้ำ)</label>
      <input name="em_id" required value="{{ e.em_id or '' }}" />
    </div>
    <div>
      <label>ID Card</label>
      <input name="id_card" value="{{ e.id_card or '' }}" />
    </div>
  </div>

  <div class="hr"></div>

  <h3>ชื่อไทย</h3>
  <div class="row3">
    <div><label>คำนำหน้า (TH)</label><input name="title_th" value="{{ e.title_th or '' }}"></div>
    <div><label>ชื่อ (TH)</label><input name="first_name_th" value="{{ e.first_name_th or '' }}"></div>
    <div><label>สกุล (TH)</label><input name="last_name_th" value="{{ e.last_name_th or '' }}"></div>
  </div>

  <h3 style="margin-top:16px;">ชื่ออังกฤษ</h3>
  <div class="row3">
    <div><label>Title (EN)</label><input name="title_en" value="{{ e.title_en or '' }}"></div>
    <div><label>First name (EN)</label><input name="first_name_en" value="{{ e.first_name_en or '' }}"></div>
    <div><label>Last name (EN)</label><input name="last_name_en" value="{{ e.last_name_en or '' }}"></div>
  </div>

  <div class="hr"></div>

  <h3>ตำแหน่ง/หน่วยงาน</h3>
  <div class="row3">
    <div><label>Position</label><input name="position" value="{{ e.position or '' }}"></div>
    <div><label>Section</label><input name="section" value="{{ e.section or '' }}"></div>
    <div><label>Department</label><input name="department" value="{{ e.department or '' }}"></div>
  </div>

  <div class="hr"></div>

  <h3>วันที่/สถานะ</h3>
  <div class="row3">
    <div><label>Start work (dd-mm-yy หรือ yyyy-mm-dd)</label><input name="start_work" value="{{ e.start_work or '' }}"></div>
    <div><label>Resign (dd-mm-yy หรือ yyyy-mm-dd)</label><input name="resign_date" value="{{ e.resign_date or '' }}"></div>
    <div><label>Status</label><input name="status" value="{{ e.status or '' }}"></div>
  </div>

  <div class="row3" style="margin-top:12px;">
    <div><label>Degree</label><input name="degree" value="{{ e.degree or '' }}"></div>
    <div><label>Major</label><input name="major" value="{{ e.major or '' }}"></div>
    <div style="display:flex; align-items:end;">
      <button class="btn btn-primary" type="submit">บันทึก</button>
    </div>
  </div>
</form>
"""

IMPORT_BODY = """
<div class="actions">
  <a class="btn" href="/employees">← กลับ</a>
</div>

<div class="card">
  <h2>Import Employee Data (Excel)</h2>
  <p class="muted">รองรับ .xlsx และหัวคอลัมน์: No., Em. ID, ID Card, Name-TH, Name-EN, Position, Section, Department, Start work, Resign, Status, Degree, Major</p>

  <form method="post" enctype="multipart/form-data">
    <input type="file" name="file" accept=".xlsx" required>
    <div style="margin-top:12px;">
      <button class="btn btn-primary" type="submit">อัปโหลด & Import</button>
    </div>
  </form>

  {% if msg %}
    <p style="margin-top:12px;"><b>{{ msg }}</b></p>
  {% endif %}
</div>
"""

# ---------------------------
# Routes
# ---------------------------
@app.get("/")
def home():
    return render_template_string(
        BASE_HTML,
        title="SANKO Training System",
        header="SANKO Training System 🚀",
        subtitle="ระบบกำลังทำงานอยู่",
        body=HOME_BODY,
    )

@app.get("/training")
def training():
    return render_template_string(
        BASE_HTML,
        title="Training Matrix",
        header="Training Matrix",
        subtitle="ตัวอย่างหน้า Training",
        body=TRAINING_BODY,
    )

@app.route("/employees", methods=["GET"])
def employees_list():
    q = (request.args.get("q") or "").strip()
    dept = (request.args.get("dept") or "").strip()

    query = Employee.query

    if q:
        like = f"%{q}%"
        query = query.filter(
            db.or_(
                Employee.em_id.ilike(like),
                Employee.id_card.ilike(like),
                Employee.first_name_th.ilike(like),
                Employee.last_name_th.ilike(like),
                Employee.first_name_en.ilike(like),
                Employee.last_name_en.ilike(like),
                Employee.department.ilike(like),
                Employee.section.ilike(like),
                Employee.position.ilike(like),
            )
        )

    if dept:
        query = query.filter(Employee.department.ilike(f"%{dept}%"))

    employees = query.order_by(Employee.department.asc(), Employee.section.asc(), Employee.em_id.asc()).all()

    return render_template_string(
        BASE_HTML,
        title="Employee Report",
        header="Employee Report",
        subtitle="ฐานข้อมูลพนักงาน (เพิ่ม/แก้ไข/ลบ/นำเข้า Excel)",
        body=render_template_string(EMP_LIST_BODY, employees=employees, total=len(employees), q=q, dept=dept),
    )

@app.route("/employees/new", methods=["GET", "POST"])
def employees_new():
    if request.method == "POST":
        em_id = (request.form.get("em_id") or "").strip()
        if not em_id:
            return "Em. ID required", 400

        exists = Employee.query.filter_by(em_id=em_id).first()
        if exists:
            return f"Em. ID ซ้ำ: {em_id}", 400

        e = Employee(em_id=em_id)
        fill_employee_from_form(e, request.form)
        db.session.add(e)
        db.session.commit()
        return redirect(url_for("employees_list"))

    e = Employee()
    return render_template_string(
        BASE_HTML,
        title="Add Employee",
        header="เพิ่มพนักงาน",
        subtitle="กรอกข้อมูลแล้วกดบันทึก",
        body=render_template_string(EMP_FORM_BODY, e=e),
    )

@app.route("/employees/<int:emp_id>/edit", methods=["GET", "POST"])
def employees_edit(emp_id: int):
    e = Employee.query.get_or_404(emp_id)

    if request.method == "POST":
        new_em_id = (request.form.get("em_id") or "").strip()
        if not new_em_id:
            return "Em. ID required", 400

        # ถ้าเปลี่ยน Em. ID ต้องเช็คซ้ำ
        if new_em_id != e.em_id:
            exists = Employee.query.filter_by(em_id=new_em_id).first()
            if exists:
                return f"Em. ID ซ้ำ: {new_em_id}", 400

        e.em_id = new_em_id
        fill_employee_from_form(e, request.form)
        db.session.commit()
        return redirect(url_for("employees_list"))

    return render_template_string(
        BASE_HTML,
        title="Edit Employee",
        header=f"แก้ไขพนักงาน: {e.em_id}",
        subtitle="แก้ไขข้อมูลแล้วกดบันทึก",
        body=render_template_string(EMP_FORM_BODY, e=e),
    )

@app.post("/employees/<int:emp_id>/delete")
def employees_delete(emp_id: int):
    e = Employee.query.get_or_404(emp_id)
    db.session.delete(e)
    db.session.commit()
    return redirect(url_for("employees_list"))


@app.route("/employees/import", methods=["GET", "POST"])
def employees_import():
    msg = ""
    if request.method == "POST":
        f = request.files.get("file")
        if not f:
            return "No file", 400

        df = pd.read_excel(f)
        df = normalize_columns(df)

        imported = 0
        updated = 0

        for _, row in df.iterrows():
            em_id = str(row.get("Em. ID", "")).strip()
            if not em_id or em_id.lower() == "nan":
                continue

            emp = Employee.query.filter_by(em_id=em_id).first()
            is_new = False
            if not emp:
                emp = Employee(em_id=em_id)
                db.session.add(emp)
                is_new = True

            # no
            try:
                n = row.get("No.", None)
                emp.no = int(n) if pd.notna(n) and str(n).strip() != "" else None
            except Exception:
                emp.no = None

            emp.id_card = str(row.get("ID Card", "")).strip()

            # Names
            th_full = str(row.get("Name-TH", "")).strip()
            en_full = str(row.get("Name-EN", "")).strip()

            t_th, f_th, l_th = split_th_name(th_full)
            t_en, f_en, l_en = split_en_name(en_full)

            emp.title_th = t_th
            emp.first_name_th = f_th
            emp.last_name_th = l_th

            emp.title_en = t_en
            emp.first_name_en = f_en
            emp.last_name_en = l_en

            # Org
            emp.position = str(row.get("Position", "")).strip()
            emp.section = str(row.get("Section", "")).strip()
            emp.department = str(row.get("Department", "")).strip()

            # Dates
            emp.start_work = parse_date(row.get("Start work"))
            emp.resign_date = parse_date(row.get("Resign"))

            # Status
            emp.status = str(row.get("Status", "")).strip()
            emp.degree = str(row.get("Degree", "")).strip()
            emp.major = str(row.get("Major", "")).strip()

            if is_new:
                imported += 1
            else:
                updated += 1

        db.session.commit()
        msg = f"Import สำเร็จ ✅ เพิ่มใหม่ {imported} รายการ | อัปเดต {updated} รายการ"

    return render_template_string(
        BASE_HTML,
        title="Import Employees",
        header="Import Employee Data",
        subtitle="อัปโหลด Excel เพื่อเพิ่ม/อัปเดตฐานข้อมูลพนักงาน",
        body=render_template_string(IMPORT_BODY, msg=msg),
    )


def fill_employee_from_form(e: Employee, form):
    # Numeric
    no = (form.get("no") or "").strip()
    try:
        e.no = int(no) if no else None
    except Exception:
        e.no = None

    e.id_card = (form.get("id_card") or "").strip()

    e.title_th = (form.get("title_th") or "").strip()
    e.first_name_th = (form.get("first_name_th") or "").strip()
    e.last_name_th = (form.get("last_name_th") or "").strip()

    e.title_en = (form.get("title_en") or "").strip()
    e.first_name_en = (form.get("first_name_en") or "").strip()
    e.last_name_en = (form.get("last_name_en") or "").strip()

    e.position = (form.get("position") or "").strip()
    e.section = (form.get("section") or "").strip()
    e.department = (form.get("department") or "").strip()

    e.start_work = parse_date(form.get("start_work"))
    e.resign_date = parse_date(form.get("resign_date"))

    e.status = (form.get("status") or "").strip()
    e.degree = (form.get("degree") or "").strip()
    e.major = (form.get("major") or "").strip()


# ---------------------------
# Render / Production
# ---------------------------
# สำคัญ: Render จะรันผ่าน gunicorn "app:app"
# ไม่ต้อง app.run() ใน production
if __name__ == "__main__":
    ensure_db()
    app.run(host="0.0.0.0", port=5000, debug=True)
else:
    ensure_db()
