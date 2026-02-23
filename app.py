from flask import Flask, request, redirect, url_for, render_template_string, send_file, flash
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
import pandas as pd
import io

app = Flask(__name__)
app.secret_key = "change-me"

# --- DB (เริ่มแบบ SQLite ก่อน) ---
app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///data.db"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
db = SQLAlchemy(app)


class Employee(db.Model):
    id = db.Column(db.Integer, primary_key=True)

    em_id = db.Column(db.String(50), unique=True, index=True)   # Em. ID
    id_card = db.Column(db.String(50), index=True)              # ID Card

    title_th = db.Column(db.String(20))
    first_name_th = db.Column(db.String(100))
    last_name_th = db.Column(db.String(100))

    title_en = db.Column(db.String(20))
    first_name_en = db.Column(db.String(100))
    last_name_en = db.Column(db.String(100))

    position = db.Column(db.String(120))
    section = db.Column(db.String(120))
    department = db.Column(db.String(120))

    start_work = db.Column(db.Date)
    resign_date = db.Column(db.Date)

    status = db.Column(db.String(30))
    degree = db.Column(db.String(50))
    major = db.Column(db.String(120))

    created_at = db.Column(db.DateTime, default=datetime.utcnow)


with app.app_context():
    db.create_all()


def parse_date(v):
    if v is None or (isinstance(v, float) and pd.isna(v)) or (isinstance(v, str) and v.strip() == ""):
        return None
    # รองรับทั้ง "01-01-96" หรือ date/datetime จาก excel
    try:
        if isinstance(v, (datetime, )):
            return v.date()
        return pd.to_datetime(v, dayfirst=True, errors="coerce").date()
    except Exception:
        return None


HOME_HTML = """
<h1>SANKO Training System 🚀</h1>
<h3>เมนูหลัก</h3>
<ul>
  <li><a href="/employees">Employee Data</a></li>
  <li><a href="/employees/import">Import Excel</a></li>
  <li><a href="/employees/export">Export CSV</a></li>
</ul>
"""

LIST_HTML = """
<h2>Employee Data</h2>

<form method="get">
  <input name="q" placeholder="ค้นหา (Em.ID/ชื่อ/แผนก)" value="{{q}}" />
  <button type="submit">ค้นหา</button>
  <a href="/employees/new">+ เพิ่มพนักงาน</a>
  | <a href="/">กลับหน้าแรก</a>
</form>

<br>

<table border="1" cellpadding="6">
  <tr>
    <th>Em.ID</th>
    <th>ID Card</th>
    <th>Name-TH</th>
    <th>Name-EN</th>
    <th>Position</th>
    <th>Section</th>
    <th>Department</th>
    <th>Start</th>
    <th>Resign</th>
    <th>Status</th>
    <th>Degree</th>
    <th>Major</th>
    <th>Action</th>
  </tr>

  {% for e in rows %}
  <tr>
    <td>{{e.em_id}}</td>
    <td>{{e.id_card}}</td>
    <td>{{e.title_th}} {{e.first_name_th}} {{e.last_name_th}}</td>
    <td>{{e.title_en}} {{e.first_name_en}} {{e.last_name_en}}</td>
    <td>{{e.position}}</td>
    <td>{{e.section}}</td>
    <td>{{e.department}}</td>
    <td>{{e.start_work}}</td>
    <td>{{e.resign_date}}</td>
    <td>{{e.status}}</td>
    <td>{{e.degree}}</td>
    <td>{{e.major}}</td>
    <td>
      <a href="/employees/{{e.id}}/edit">แก้ไข</a>
      <form method="post" action="/employees/{{e.id}}/delete" style="display:inline">
        <button type="submit" onclick="return confirm('ลบรายการนี้?')">ลบ</button>
      </form>
    </td>
  </tr>
  {% endfor %}
</table>
"""

FORM_HTML = """
<h2>{{'แก้ไข' if emp else 'เพิ่ม'}}พนักงาน</h2>
<form method="post">
  Em.ID: <input name="em_id" value="{{emp.em_id if emp else ''}}" required><br><br>
  ID Card: <input name="id_card" value="{{emp.id_card if emp else ''}}"><br><br>

  Title-TH: <input name="title_th" value="{{emp.title_th if emp else ''}}">
  First-TH: <input name="first_name_th" value="{{emp.first_name_th if emp else ''}}">
  Last-TH: <input name="last_name_th" value="{{emp.last_name_th if emp else ''}}">
  <br><br>

  Title-EN: <input name="title_en" value="{{emp.title_en if emp else ''}}">
  First-EN: <input name="first_name_en" value="{{emp.first_name_en if emp else ''}}">
  Last-EN: <input name="last_name_en" value="{{emp.last_name_en if emp else ''}}">
  <br><br>

  Position: <input name="position" value="{{emp.position if emp else ''}}"><br><br>
  Section: <input name="section" value="{{emp.section if emp else ''}}"><br><br>
  Department: <input name="department" value="{{emp.department if emp else ''}}"><br><br>

  Start work (dd-mm-yy): <input name="start_work" value="{{emp.start_work if emp else ''}}"><br><br>
  Resign (dd-mm-yy): <input name="resign_date" value="{{emp.resign_date if emp else ''}}"><br><br>

  Status: <input name="status" value="{{emp.status if emp else ''}}"><br><br>
  Degree: <input name="degree" value="{{emp.degree if emp else ''}}"><br><br>
  Major: <input name="major" value="{{emp.major if emp else ''}}"><br><br>

  <button type="submit">บันทึก</button>
  <a href="/employees">ยกเลิก</a>
</form>
"""

IMPORT_HTML = """
<h2>Import Excel (.xlsx)</h2>
<p>ต้องมีคอลัมน์ใกล้เคียง: Em. ID, ID Card, Name-TH, Name-EN, Position, Section, Department, Start work, Resign, Status, Degree, Major</p>

<form method="post" enctype="multipart/form-data">
  <input type="file" name="file" accept=".xlsx" required>
  <button type="submit">อัปโหลด & อิมพอร์ต</button>
</form>
<br>
<a href="/">กลับหน้าแรก</a>
"""


@app.get("/")
def home():
    return render_template_string(HOME_HTML)


@app.get("/employees")
def employees_list():
    q = request.args.get("q", "").strip()
    query = Employee.query
    if q:
        like = f"%{q}%"
        query = query.filter(
            (Employee.em_id.like(like)) |
            (Employee.first_name_th.like(like)) |
            (Employee.last_name_th.like(like)) |
            (Employee.department.like(like)) |
            (Employee.section.like(like))
        )
    rows = query.order_by(Employee.created_at.desc()).limit(500).all()
    return render_template_string(LIST_HTML, rows=rows, q=q)


@app.route("/employees/new", methods=["GET", "POST"])
def employees_new():
    if request.method == "POST":
        emp = Employee(
            em_id=request.form.get("em_id"),
            id_card=request.form.get("id_card"),
            title_th=request.form.get("title_th"),
            first_name_th=request.form.get("first_name_th"),
            last_name_th=request.form.get("last_name_th"),
            title_en=request.form.get("title_en"),
            first_name_en=request.form.get("first_name_en"),
            last_name_en=request.form.get("last_name_en"),
            position=request.form.get("position"),
            section=request.form.get("section"),
            department=request.form.get("department"),
            start_work=parse_date(request.form.get("start_work")),
            resign_date=parse_date(request.form.get("resign_date")),
            status=request.form.get("status"),
            degree=request.form.get("degree"),
            major=request.form.get("major"),
        )
        db.session.add(emp)
        db.session.commit()
        return redirect(url_for("employees_list"))
    return render_template_string(FORM_HTML, emp=None)


@app.route("/employees/<int:emp_id>/edit", methods=["GET", "POST"])
def employees_edit(emp_id):
    emp = Employee.query.get_or_404(emp_id)
    if request.method == "POST":
        for f in ["em_id","id_card","title_th","first_name_th","last_name_th","title_en",
                  "first_name_en","last_name_en","position","section","department","status","degree","major"]:
            setattr(emp, f, request.form.get(f))
        emp.start_work = parse_date(request.form.get("start_work"))
        emp.resign_date = parse_date(request.form.get("resign_date"))
        db.session.commit()
        return redirect(url_for("employees_list"))
    return render_template_string(FORM_HTML, emp=emp)


@app.post("/employees/<int:emp_id>/delete")
def employees_delete(emp_id):
    emp = Employee.query.get_or_404(emp_id)
    db.session.delete(emp)
    db.session.commit()
    return redirect(url_for("employees_list"))


@app.route("/employees/import", methods=["GET", "POST"])
def employees_import():
    if request.method == "POST":
        f = request.files.get("file")
        if not f:
            return "No file", 400

        df = pd.read_excel(f)

        # ตัวอย่างแมพคอลัมน์ (ปรับชื่อคอลัมน์ตามไฟล์จริงได้)
        # แนะนำให้ตั้งหัวคอลัมน์ใน Excel ให้ชัด
        col_map = {
            "Em. ID": "em_id",
            "ID Card": "id_card",
            "Title-TH": "title_th",
            "First-TH": "first_name_th",
            "Last-TH": "last_name_th",
            "Title-EN": "title_en",
            "First-EN": "first_name_en",
            "Last-EN": "last_name_en",
            "Position": "position",
            "Section": "section",
            "Department": "department",
            "Start work": "start_work",
            "Resign": "resign_date",
            "Status": "status",
            "Degree": "degree",
            "Major": "major",
        }

        # รองรับกรณีไฟล์คุณใช้ชื่อคอลัมน์แบบในรูป: "Em. ID", "Start wo", "Resigr"
        # เพิ่ม alias ให้มันเจอได้
        alias = {
            "Em. ID": "Em. ID",
            "Start wo": "Start work",
            "Resigr": "Resign",
            "Name-TH": None,
            "Name-EN": None
        }

        # rename columns ตามที่มี
        fixed_cols = {}
        for c in df.columns:
            if c in alias and alias[c]:
                fixed_cols[c] = alias[c]
        if fixed_cols:
            df = df.rename(columns=fixed_cols)

        imported = 0
        for _, row in df.iterrows():
            em_id = str(row.get("Em. ID", "")).strip()
            if not em_id:
                continue

            emp = Employee.query.filter_by(em_id=em_id).first()
            if not emp:
                emp = Employee(em_id=em_id)
                db.session.add(emp)

            emp.id_card = str(row.get("ID Card", "")).strip()

            # ถ้าไฟล์มี Name-TH เป็นคำรวม (เช่น "นาย สมชาย ใจดี") จะยังไม่แยกให้
            # แนะนำให้ทำคอลัมน์ Title/First/Last แยกมา จะดีที่สุด
            emp.position = str(row.get("Position", "")).strip()
            emp.section = str(row.get("Section", "")).strip()
            emp.department = str(row.get("Department", "")).strip()

            emp.start_work = parse_date(row.get("Start work"))
            emp.resign_date = parse_date(row.get("Resign"))

            emp.status = str(row.get("Status", "")).strip()
            emp.degree = str(row.get("Degree", "")).strip()
            emp.major = str(row.get("Major", "")).strip()

            imported += 1

        db.session.commit()
        return redirect(url_for("employees_list"))

    return render_template_string(IMPORT_HTML)


@app.get("/employees/export")
def employees_export():
    rows = Employee.query.order_by(Employee.id.asc()).all()
    data = []
    for e in rows:
        data.append({
            "Em. ID": e.em_id,
            "ID Card": e.id_card,
            "Title-TH": e.title_th,
            "First-TH": e.first_name_th,
            "Last-TH": e.last_name_th,
            "Title-EN": e.title_en,
            "First-EN": e.first_name_en,
            "Last-EN": e.last_name_en,
            "Position": e.position,
            "Section": e.section,
            "Department": e.department,
            "Start work": e.start_work,
            "Resign": e.resign_date,
            "Status": e.status,
            "Degree": e.degree,
            "Major": e.major,
        })
    df = pd.DataFrame(data)
    buf = io.StringIO()
    df.to_csv(buf, index=False)
    mem = io.BytesIO(buf.getvalue().encode("utf-8-sig"))
    return send_file(mem, mimetype="text/csv", as_attachment=True, download_name="employees.csv")


if __name__ == "__main__":
    app.run()
