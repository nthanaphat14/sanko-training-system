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

from datetime import datetime, date

class TrainingRecord(db.Model):
    __tablename__ = "training_records"
    id = db.Column(db.Integer, primary_key=True)

    seq = db.Column(db.Integer, nullable=True)               # ลำดับ
    year = db.Column(db.Integer, nullable=True)              # Year.
    month = db.Column(db.Integer, nullable=True)             # Month

    emp_id = db.Column(db.String(50), nullable=False, index=True)  # Emp ID

    prefix = db.Column(db.String(50), nullable=True)
    full_name = db.Column(db.String(200), nullable=True)
    last_name = db.Column(db.String(200), nullable=True)
    
    department = db.Column(db.String(150), nullable=True)    # แผนก
    position = db.Column(db.String(150), nullable=True)      # ตำแหน่ง

    course_code = db.Column(db.String(100), nullable=True)   # รหัสหลักสูตร
    course_name = db.Column(db.String(255), nullable=True)   # ชื่อหลักสูตร
    course_type = db.Column(db.String(100), nullable=True)   # ประเภท

    start_date = db.Column(db.Date, nullable=True)           # StartDate
    end_date = db.Column(db.Date, nullable=True)             # EndDate
    hours = db.Column(db.Float, nullable=True)               # ชั่วโมง

    evaluate_method = db.Column(db.String(150), nullable=True)  # วิธีประเมิน
    result = db.Column(db.String(50), nullable=True)            # ผล
    score = db.Column(db.Float, nullable=True)                  # คะแนน
    evaluator = db.Column(db.String(150), nullable=True)        # ผู้ประเมิน

    expire_date = db.Column(db.Date, nullable=True)          # วันหมดอายุ
    remark = db.Column(db.Text, nullable=True)               # หมายเหตุ

    created_at = db.Column(db.DateTime, default=datetime.utcnow)


def init_db():
    with app.app_context():
        db.create_all()


# -------------------------------------------------
# Helper Functions
# -------------------------------------------------
    
def safe_str(v):
    if v is None:
        return ""
    return str(v).strip()

def safe_int(x):
    try:
        if x is None: return None
        s = str(x).strip()
        if s == "": return None
        return int(float(s))
    except:
        return None
        
def safe_month(v):
    if v is None:
        return None

    # ถ้าเป็นเลขอยู่แล้ว
    try:
        iv = int(v)
        if 1 <= iv <= 12:
            return iv
    except:
        pass

    s = str(v).strip().lower()
    m = {
        "jan": 1, "january": 1,
        "feb": 2, "february": 2,
        "mar": 3, "march": 3,
        "apr": 4, "april": 4,
        "may": 5,
        "jun": 6, "june": 6,
        "jul": 7, "july": 7,
        "aug": 8, "august": 8,
        "sep": 9, "sept": 9, "september": 9,
        "oct": 10, "october": 10,
        "nov": 11, "november": 11,
        "dec": 12, "december": 12,
    }
    return m.get(s)       

def safe_date(v):
    """
    รับได้ทั้ง:
    - datetime/date
    - ตัวเลข excel date (บางครั้ง openpyxl จะให้เป็น datetime อยู่แล้ว)
    - string หลาย format เช่น 2026-02-25, 25/02/2026, 25-02-2026
    """
    if v is None:
        return None

    # ถ้าเป็น date/datetime อยู่แล้ว
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v

    s = str(v).strip()
    if not s or s.lower() in ("none", "nan"):
        return None

    # ลอง format ที่พบบ่อย
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except:
            pass

    return None
    
    MONTH_MAP = {
        "jan": 1, "january": 1,
        "feb": 2, "february": 2,
        "mar": 3, "march": 3,
        "apr": 4, "april": 4,
        "may": 5,
        "jun": 6, "june": 6,
        "jul": 7, "july": 7,
        "aug": 8, "august": 8,
        "sep": 9, "september": 9,
        "oct": 10, "october": 10,
        "nov": 11, "november": 11,
        "dec": 12, "december": 12,
    }

def month_to_int(v):
    s = safe_str(v).strip().lower()
    if s.isdigit():
        return int(s)
    return MONTH_MAP.get(s[:3]) or MONTH_MAP.get(s)

def safe_float(x):
    try:
        if x is None: return None
        s = str(x).strip()
        if s == "": return None
        return float(s)
    except:
        return None

def parse_date(v):
    # รองรับ date/datetime จาก Excel + string หลายรูปแบบ
    if not v:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except:
            pass
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

    start_date = date(year, month, 1) if month else date(year, 1, 1)
    end_date = date(year + (1 if month == 12 else 0), (1 if month == 12 else (month + 1)), 1) if month else date(year + 1, 1, 1)

# ---- เข้า: start_work ตามช่วง ----
    dept_join_rows = (
    db.session.query(Employee.section, func.count(Employee.id))
    .filter(Employee.start_work >= start_date, Employee.start_work < end_date)
    .group_by(Employee.section)
    .all()
    )

# ---- ออก: resign ตามช่วง ----
    dept_resign_rows = (
    db.session.query(Employee.section, func.count(Employee.id))
    .filter(Employee.resign >= start_date, Employee.resign < end_date)
    .group_by(Employee.section)
    .all()
    )

# map เป็น dict เพื่อ merge กันง่าย
    join_map = { (d or "ไม่ระบุ"): int(c) for d, c in dept_join_rows }
    resign_map = { (d or "ไม่ระบุ"): int(c) for d, c in dept_resign_rows }

# รวมเป็น list เดียว
    dept_inout = []
    all_depts = set(join_map.keys()) | set(resign_map.keys())
    for d in all_depts:
        dept_inout.append({
        "section": d,
        "joined": join_map.get(d, 0),
        "resigned": resign_map.get(d, 0),
        "net": join_map.get(d, 0) - resign_map.get(d, 0),
    })

# เรียง: ออกมากสุดก่อน หรือเข้าออกมากสุดก่อนก็ได้
    dept_inout.sort(key=lambda x: (x["joined"] + x["resigned"]), reverse=True)

# เอา Top 10
    dept_inout_top10 = dept_inout[:10]
    
    return render_template(
        "dashboard.html",
        total=total,
        active=active,
        resigned_total=resigned_total,
        resigned=resigned_total,
        year=year,
        month=month,
        joined_year=joined_year,
        resigned_year=resigned_year,
        joined_month=joined_month,
        resigned_month=resigned_month,
        month_summary=month_summary,
        dept_inout_top10=dept_inout_top10,
        start_date=start_date,
        end_date=end_date,
    )
    
@app.get("/trainings")
def trainings_list():
    q = (request.args.get("q") or "").strip()
    year = (request.args.get("year") or "").strip()
    month = (request.args.get("month") or "").strip()

    query = TrainingRecord.query

    if q:
        like = f"%{q}%"
        query = query.filter(
            db.or_(
                TrainingRecord.emp_id.ilike(like),
                TrainingRecord.full_name.ilike(like),
                TrainingRecord.course_code.ilike(like),
                TrainingRecord.course_name.ilike(like),
                TrainingRecord.department.ilike(like),
                TrainingRecord.position.ilike(like),
            )
        )

    if year.isdigit():
        query = query.filter(TrainingRecord.year == int(year))
    if month.isdigit():
        query = query.filter(TrainingRecord.month == int(month))

    rows = query.order_by(
        TrainingRecord.start_date.desc().nullslast(),
        TrainingRecord.id.desc()
    ).limit(500).all()

    total = query.count()

    return render_template("trainings_list.html", rows=rows, total=total, q=q, year=year, month=month)

@app.route("/trainings/import", methods=["GET", "POST"])
def trainings_import():
    if request.method == "GET":
        return render_template("training_import.html")

    f = request.files.get("file")
    if not f or f.filename == "":
        flash("กรุณาเลือกไฟล์ Excel", "error")
        return redirect(url_for("trainings_import"))

    wb = load_workbook(f, data_only=True)
    ws = wb["Record Training"] if "Record Training" in wb.sheetnames else wb.active

# --- header map (row 1) ---
def norm(s: str) -> str:
    return safe_str(s).lower().replace(" ", "").replace("_", "").replace("-", "")

headers = [safe_str(ws.cell(1, c).value) for c in range(1, ws.max_column + 1)]
header_map = {norm(h): i + 1 for i, h in enumerate(headers)}

ALIASES = {
    "year.": ["year.", "year", "ปี"],
    "month": ["month", "mon", "เดือน"],
    "empid": ["empid", "emp id", "รหัสพนักงาน", "emp"],
    "คำนำหน้า": ["คำนำหน้า", "prefix"],
    "ชื่อ": ["ชื่อ", "firstname", "first name"],
    "นามสกุล": ["นามสกุล", "lastname", "last name"],
    "แผนก": ["แผนก", "section", "department"],
    "ตำแหน่ง": ["ตำแหน่ง", "position"],
    "รหัสหลักสูตร": ["รหัสหลักสูตร", "coursecode", "course code"],
    "ชื่อหลักสูตร": ["ชื่อหลักสูตร", "coursename", "course name"],
    "ประเภท": ["ประเภท", "category", "coursetype", "course type"],
    "startdate": ["startdate", "start date", "วันที่เริ่ม"],
    "enddate": ["enddate", "end date", "วันที่จบ"],
    "ชั่วโมง": ["ชั่วโมง", "hours", "hour"],
    "วิธีประเมิน": ["วิธีประเมิน", "evaluatemethod", "evaluate method"],
    "ผล": ["ผล", "result"],
    "คะแนน": ["คะแนน", "score"],
    "ผู้ประเมิน": ["ผู้ประเมิน", "evaluator"],
    "วันหมดอายุ": ["วันหมดอายุ", "expiredate", "expire date"],
    "หมายเหตุ": ["หมายเหตุ", "remark", "note"],
}

def col(name: str):
    k = norm(name)
    if k in header_map:
        return header_map[k]
    for alt in ALIASES.get(k, []):
        kk = norm(alt)
        if kk in header_map:
            return header_map[kk]
    return None

def cellv(r, name):
    idx = col(name)
    if not idx:
        return None
    return ws.cell(r, idx).value

    added = 0
    skipped = 0

for r in range(2, ws.max_row + 1):

    emp_id = safe_str(cellv(r, "Emp ID"))
    if not emp_id:
        skipped += 1
        continue

    prefix = safe_str(cellv(r, "คำนำหน้า"))

    # --- ชื่อ/นามสกุล ---
    # กรณีไฟล์มีคอลัมน์ "ชื่อ-สกุล" (รวม)
    full_name_raw = safe_str(cellv(r, "ชื่อ-สกุล"))
    if full_name_raw:
        full_name = full_name_raw
        # ถ้าไม่มีคอลัมน์ "นามสกุล" ก็ปล่อยว่าง
        last_name = safe_str(cellv(r, "นามสกุล"))
    else:
        # กรณีแยก "ชื่อ" + "นามสกุล"
        full_name = safe_str(cellv(r, "ชื่อ"))
        last_name = safe_str(cellv(r, "นามสกุล"))

    # --- Section / Position ---
    section = safe_str(cellv(r, "แผนก"))     # ในไฟล์ = Quality Control (Section)
    position = safe_str(cellv(r, "ตำแหน่ง"))  # ในไฟล์ = Operator (Position)

    tr = TrainingRecord(
        year=safe_int(cellv(r, "Year.")),
        month=safe_month(cellv(r, "Month")),

        emp_id=emp_id,
        prefix=prefix,
        full_name=full_name,
        last_name=last_name,

        department=section,   # เก็บ section ลงช่องเดิมของ DB (department)
        position=position,

        course_code=safe_str(cellv(r, "รหัสหลักสูตร")),
        course_name=safe_str(cellv(r, "ชื่อหลักสูตร")),
        course_type=safe_str(cellv(r, "ประเภท")),

        start_date=safe_date(cellv(r, "StartDate")),
        end_date=safe_date(cellv(r, "EndDate")),
        hours=safe_float(cellv(r, "ชั่วโมง")),

        evaluate_method=safe_str(cellv(r, "วิธีประเมิน")),
        result=safe_str(cellv(r, "ผล")),
        score=safe_float(cellv(r, "คะแนน")),
        evaluator=safe_str(cellv(r, "ผู้ประเมิน")),

        expire_date=safe_date(cellv(r, "วันหมดอายุ")),
        remark=safe_str(cellv(r, "หมายเหตุ")),
    )
            
        db.session.add(tr)
        added += 1

    db.session.commit()
    flash(f"Import สำเร็จ: {added} รายการ | ข้าม: {skipped} แถว", "success")
    return redirect(url_for("trainings_list"))

@app.route("/trainings/new", methods=["GET", "POST"])
def trainings_new():
    if request.method == "GET":
        return render_template("trainings_new.html")

    emp_id = safe_str(request.form.get("emp_id"))
    if not emp_id:
        flash("กรุณากรอก Emp ID", "error")
        return redirect(url_for("trainings_new"))

    tr = TrainingRecord(
        year=safe_int(request.form.get("year")),
        month=safe_int(request.form.get("month")),
        emp_id=emp_id,
        prefix=safe_str(request.form.get("prefix")),
        full_name=safe_str(request.form.get("full_name")),
        last_name=safe_str(request.form.get("last_name")),
        department=safe_str(request.form.get("department")),
        position=safe_str(request.form.get("position")),
        course_code=safe_str(request.form.get("course_code")),
        course_name=safe_str(request.form.get("course_name")),
        course_type=safe_str(request.form.get("course_type")),
        start_date=safe_date(request.form.get("start_date")),
        end_date=safe_date(request.form.get("end_date")),
        hours=safe_float(request.form.get("hours")),
        evaluate_method=safe_str(request.form.get("evaluate_method")),
        result=safe_str(request.form.get("result")),
        score=safe_float(request.form.get("score")),
        evaluator=safe_str(request.form.get("evaluator")),
        expire_date=safe_date(request.form.get("expire_date")),
        remark=safe_str(request.form.get("remark")),
    )

    db.session.add(tr)
    db.session.commit()
    flash("บันทึก Training Record แล้ว", "success")
    return redirect(url_for("trainings_list"))

    
# -------------------------------------------------
# Run (Local Only)
# -------------------------------------------------
def init_db():
    with app.app_context():
        db.create_all()

init_db()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
