import os
from datetime import datetime
from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    flash,
)
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import or_, nullslast
from sqlalchemy.exc import IntegrityError, DataError
from io import BytesIO
from flask import send_file
from openpyxl import load_workbook
from openpyxl import Workbook
from sqlalchemy import func  
from datetime import date


# -------------------------------------------------
# App Config
# -------------------------------------------------
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "super-secret-key")

# DATABASE
db_url = (os.environ.get("DATABASE_URL") or "").strip()

if db_url.startswith("postgres://"):
    db_url = db_url.replace("postgres://", "postgresql://", 1)

app.config["SQLALCHEMY_DATABASE_URI"] = db_url or "sqlite:///employee.db"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)


# -------------------------------------------------
# Model
# -------------------------------------------------
class Employee(db.Model):
    __tablename__ = "employees"

    def th_full(self):
        """คืนชื่อ-สกุลภาษาไทยแบบรวม"""
        first = (self.first_name_th or "").strip()
        last = (self.last_name_th or "").strip()
        full = f"{first} {last}".strip()
        return full

    def en_full(self):   # ← ต้องเยื้อง 4 ช่อง
        """คืนชื่อ-สกุลภาษาอังกฤษแบบรวม"""
        first = (self.first_name_en or "").strip()
        last = (self.last_name_en or "").strip()
        full = f"{first} {last}".strip()
        return full
    
    id = db.Column(db.Integer, primary_key=True)
    no = db.Column(db.Integer, nullable=True)
    em_id = db.Column(db.String(40), unique=True, nullable=False)
    id_card = db.Column(db.String(60), nullable=True)

    title_th = db.Column(db.String(50), nullable=True)
    first_name_th = db.Column(db.String(120), nullable=True)
    last_name_th = db.Column(db.String(120), nullable=True)

    title_en = db.Column(db.String(50), nullable=True)
    first_name_en = db.Column(db.String(120), nullable=True)
    last_name_en = db.Column(db.String(120), nullable=True)

    position = db.Column(db.String(150), nullable=True)
    section = db.Column(db.String(150), nullable=True)
    department = db.Column(db.String(150), nullable=True)

    start_work = db.Column(db.Date, nullable=True)
    resign = db.Column(db.Date, nullable=True)

    status = db.Column(db.String(50), nullable=True)
    degree = db.Column(db.String(80), nullable=True)
    major = db.Column(db.String(150), nullable=True)

    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(
        db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow
    )


def init_db():
    db.create_all()


# -------------------------------------------------
# Helper Functions
# -------------------------------------------------
def safe_str(x):
    if not x:
        return None
    x = str(x).strip()
    return x if x else None


def safe_int(x):
    try:
        return int(x)
    except:
        return None


def safe_date(x):
    try:
        return datetime.strptime(x, "%Y-%m-%d").date()
    except:
        return None


# -------------------------------------------------
# Routes
# -------------------------------------------------
@app.get("/")
def root():
    return redirect(url_for("employees_list"))


@app.get("/healthz")
def healthz():
    return {"status": "ok"}, 200


@app.get("/employees")
def employees_list():
    q = (request.args.get("q") or "").strip()

    query = Employee.query  # ✅ ต้องอยู่ตรงนี้ (ก่อน if q)

    if q:
        like = f"%{q}%"
        query = query.filter(
            or_(
                Employee.em_id.ilike(like),
                Employee.id_card.ilike(like),
                Employee.first_name_th.ilike(like),
                Employee.last_name_th.ilike(like),
                Employee.first_name_en.ilike(like),
                Employee.last_name_en.ilike(like),
                Employee.position.ilike(like),
                Employee.department.ilike(like),
                Employee.status.ilike(like),
            )
        )

    employees = query.order_by(nullslast(Employee.no.asc()), Employee.em_id.asc()).all()

    return render_template(
        "employees.html",
        employees=employees,
        q=q,
        total=len(employees),
    )
    
@app.route("/employees/new", methods=["GET", "POST"])
def employee_new():
    if request.method == "POST":
        em_id = request.form.get("em_id", "").strip()
        first_name_th = request.form.get("first_name_th", "").strip()

        if not em_id:
            flash("กรุณากรอกรหัสพนักงาน", "error")
            return redirect(url_for("employee_new"))

        if Employee.query.filter_by(em_id=em_id).first():
            flash("em_id นี้มีอยู่แล้ว", "error")
            return redirect(url_for("employee_new"))

        emp = Employee(em_id=em_id, first_name_th=first_name_th)
        db.session.add(emp)
        db.session.commit()

        flash("เพิ่มพนักงานเรียบร้อย", "success")
        return redirect(url_for("employees_list"))

    return render_template("employee_form.html", employee=None)

from sqlalchemy.exc import IntegrityError, DataError

def normalize_status(x):
    if not x:
        return None
    s = str(x).strip().upper()
    if s in ("W", "WORKING", "ACTIVE", "ทำงาน", "ยังอยู่"):
        return "W"
    if s in ("RS", "RESIGN", "RESIGNED", "ลาออก", "ออก"):
        return "RS"
    return s

@app.route("/employees/<string:em_id>/edit", methods=["GET", "POST"])
def employee_edit(em_id):
    emp = Employee.query.filter_by(em_id=em_id).first_or_404()

    if request.method == "POST":
        try:
            emp.id_card = safe_str(request.form.get("id_card"))
            emp.first_name_th = safe_str(request.form.get("first_name_th"))
            emp.last_name_th  = safe_str(request.form.get("last_name_th"))
            emp.first_name_en = safe_str(request.form.get("first_name_en"))
            emp.last_name_en  = safe_str(request.form.get("last_name_en"))
            emp.status = normalize_status(request.form.get("status"))

            db.session.commit()
            flash("แก้ไขข้อมูลเรียบร้อย", "success")
            return redirect(url_for("employees_list"))

        except (IntegrityError, DataError) as e:
            db.session.rollback()
            flash("บันทึกไม่สำเร็จ: ข้อมูลไม่ถูกต้องหรือซ้ำในระบบ", "error")
        except Exception as e:
            db.session.rollback()
            flash(f"เกิดข้อผิดพลาด: {e}", "error")
    return render_template("employee_form.html", employee=emp)

@app.route("/employees/<string:em_id>/delete", methods=["POST"])
def employee_delete(em_id):
    emp = Employee.query.filter_by(em_id=em_id).first_or_404()
    db.session.delete(emp)
    db.session.commit()
    flash("ลบข้อมูลเรียบร้อย", "success")
    return redirect(url_for("employees_list"))

@app.route("/employees/import", methods=["GET", "POST"])
def employees_import():
    if request.method == "GET":
        return render_template("import.html")

    f = request.files.get("file")
    if not f or f.filename == "":
        flash("กรุณาเลือกไฟล์ Excel (.xlsx) ก่อน", "error")
        return redirect(url_for("employees_import"))

    if not f.filename.lower().endswith(".xlsx"):
        flash("รองรับเฉพาะไฟล์ .xlsx เท่านั้น", "error")
        return redirect(url_for("employees_import"))

    wb = load_workbook(f, data_only=True)
    ws = wb.active

    added = 0
    updated = 0
    skipped = 0

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row:
            skipped += 1
            continue

        no_raw = row[0] if len(row) > 0 else None
        em_id_raw = row[1] if len(row) > 1 else None
        if not em_id_raw:
            skipped += 1
            continue

        em_id = str(em_id_raw).strip()

        emp = Employee.query.filter_by(em_id=em_id).first()
        is_new = emp is None
        if is_new:
            emp = Employee(em_id=em_id)

        # No
        try:
            if no_raw is not None and str(no_raw).strip() != "":
                emp.no = int(float(no_raw))
        except:
            pass

        # ID Card
        if len(row) > 2 and row[2]:
            emp.id_card = str(row[2]).strip()

        # TH
        if len(row) > 3 and row[3]:
            emp.title_th = str(row[3]).strip()
        if len(row) > 4 and row[4]:
            emp.first_name_th = str(row[4]).strip()
        if len(row) > 5 and row[5]:
            emp.last_name_th = str(row[5]).strip()

        # EN full name -> split
        if len(row) > 6 and row[6]:
            name_en = str(row[6]).strip()
            parts = [p for p in name_en.split(" ") if p]
            if len(parts) >= 2:
                emp.first_name_en = " ".join(parts[:-1])
                emp.last_name_en = parts[-1]
            else:
                emp.first_name_en = name_en

        # Position / Section / Department
        if len(row) > 7 and row[7]:
            emp.position = str(row[7]).strip()
        if len(row) > 8 and row[8]:
            emp.section = str(row[8]).strip()
        if len(row) > 9 and row[9]:
            emp.department = str(row[9]).strip()

        # Start work / Resign (แปลงเป็น Date ให้เข้ากับ model)
        if len(row) > 10 and row[10]:
            if isinstance(row[10], datetime):
                emp.start_work = row[10].date()
            else:
                emp.start_work = safe_date(str(row[10]).strip()) or emp.start_work

        if len(row) > 11 and row[11]:
            if isinstance(row[11], datetime):
                emp.resign = row[11].date()
            else:
                emp.resign = safe_date(str(row[11]).strip()) or emp.resign

        # Status / Degree / Major
        if len(row) > 12 and row[12]:
            emp.status = str(row[12]).strip()
        if len(row) > 13 and row[13]:
            emp.degree = str(row[13]).strip()
        if len(row) > 14 and row[14]:
            emp.major = str(row[14]).strip()

        if is_new:
            db.session.add(emp)
            added += 1
        else:
            updated += 1

    db.session.commit()
    flash(f"Import สำเร็จ: เพิ่มใหม่ {added} | อัปเดต {updated} | ข้าม {skipped}", "success")
    return redirect(url_for("employees_list"))

@app.get("/employees/export")
def employees_export():
    # รองรับ export ตามคำค้นเหมือนหน้า list
    q = (request.args.get("q") or "").strip()
    status = (request.args.get("status") or "").strip().upper()  # optional: W / RS

    query = Employee.query

    if q:
        like = f"%{q}%"
        query = query.filter(
            or_(
                Employee.em_id.ilike(like),
                Employee.id_card.ilike(like),
                Employee.first_name_th.ilike(like),
                Employee.last_name_th.ilike(like),
                Employee.first_name_en.ilike(like),
                Employee.last_name_en.ilike(like),
                Employee.position.ilike(like),
                Employee.section.ilike(like),
                Employee.department.ilike(like),
                Employee.status.ilike(like),
            )
        )

    if status in ("W", "RS"):
        query = query.filter(Employee.status == status)

    employees = query.order_by(nullslast(Employee.no.asc()), Employee.em_id.asc()).all()

    # สร้าง Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Employees"

    headers = [
        "No", "Em ID", "ID Card",
        "Title TH", "First TH", "Last TH",
        "Title EN", "First EN", "Last EN",
        "Position", "Section", "Department",
        "Start Work", "Resign", "Status",
        "Degree", "Major",
    ]
    ws.append(headers)

    for e in employees:
        ws.append([
            e.no,
            e.em_id,
            e.id_card,
            e.title_th,
            e.first_name_th,
            e.last_name_th,
            e.title_en,
            e.first_name_en,
            e.last_name_en,
            e.position,
            e.section,
            e.department,
            e.start_work.isoformat() if e.start_work else "",
            e.resign.isoformat() if e.resign else "",
            e.status,
            e.degree,
            e.major,
        ])

    # ทำไฟล์เป็น bytes ส่งกลับ
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    filename = "employees_export.xlsx"
    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/dashboard")
def dashboard():
    # รับค่า year/month จาก query string เช่น /dashboard?year=2026&month=2
    today = date.today()
    year = request.args.get("year", type=int) or today.year
    month = request.args.get("month", type=int)  # อาจเป็น None

    # รวมทั้งหมด / active / resigned (ของทั้งระบบ)
    total = db.session.query(func.count(Employee.id)).scalar() or 0
    active = db.session.query(func.count(Employee.id)).filter(Employee.status == "W").scalar() or 0
    resigned_total = db.session.query(func.count(Employee.id)).filter(Employee.status == "RS").scalar() or 0

    # ----- เข้า/ออก "ปีนี้" -----
    joined_year = (
        db.session.query(func.count(Employee.id))
        .filter(Employee.start_work.isnot(None))
        .filter(func.extract("year", Employee.start_work) == year)
        .scalar()
        or 0
    )
    resigned_year = (
        db.session.query(func.count(Employee.id))
        .filter(Employee.resign.isnot(None))
        .filter(func.extract("year", Employee.resign) == year)
        .scalar()
        or 0
    )

    # ----- เข้า/ออก "เดือนนี้" (ถ้าเลือก month) -----
    joined_month = None
    resigned_month = None
    if month:
        joined_month = (
            db.session.query(func.count(Employee.id))
            .filter(Employee.start_work.isnot(None))
            .filter(func.extract("year", Employee.start_work) == year)
            .filter(func.extract("month", Employee.start_work) == month)
            .scalar()
            or 0
        )
        resigned_month = (
            db.session.query(func.count(Employee.id))
            .filter(Employee.resign.isnot(None))
            .filter(func.extract("year", Employee.resign) == year)
            .filter(func.extract("month", Employee.resign) == month)
            .scalar()
            or 0
        )

    # ----- สรุปเข้า/ออก รายเดือน (1-12) ของปีที่เลือก -----
    join_rows = (
        db.session.query(func.extract("month", Employee.start_work).label("m"), func.count(Employee.id))
        .filter(Employee.start_work.isnot(None))
        .filter(func.extract("year", Employee.start_work) == year)
        .group_by("m")
        .all()
    )
    resign_rows = (
        db.session.query(func.extract("month", Employee.resign).label("m"), func.count(Employee.id))
        .filter(Employee.resign.isnot(None))
        .filter(func.extract("year", Employee.resign) == year)
        .group_by("m")
        .all()
    )

    join_map = {int(m): c for m, c in join_rows if m is not None}
    resign_map = {int(m): c for m, c in resign_rows if m is not None}

    month_summary = []
    for m in range(1, 13):
        month_summary.append({
            "month": m,
            "joined": int(join_map.get(m, 0)),
            "resigned": int(resign_map.get(m, 0)),
        })

    return render_template(
        "dashboard.html",
        total=total,
        active=active,
        resigned_total=resigned_total,
        year=year,
        month=month,
        joined_year=joined_year,
        resigned_year=resigned_year,
        joined_month=joined_month,
        resigned_month=resigned_month,
        month_summary=month_summary,
    )

# -------------------------------------------------
# Run (Local Only)
# -------------------------------------------------
def init_db():
    with app.app_context():
        db.create_all()

init_db()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
