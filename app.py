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

print("DATABASE =", app.config["SQLALCHEMY_DATABASE_URI"])  # 👈 ใส่ตรงนี้

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
    __table_args__ = (
        db.UniqueConstraint("emp_id", "start_date", name="uq_training_emp_start"),
    )
    id = db.Column(db.Integer, primary_key=True)

    seq = db.Column(db.Integer, nullable=True)               # ลำดับ
    year = db.Column(db.Integer, nullable=True)              # Year.
    month = db.Column(db.Integer, nullable=True)             # Month

    emp_id = db.Column(db.String(50), nullable=False, index=True)  # Emp ID

    prefix = db.Column(db.String(50), nullable=True)
    first_name = db.Column(db.String(200), nullable=True)
    last_name = db.Column(db.String(200), nullable=True)
    
    section = db.Column(db.String(150), nullable=True)    # แผนก
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
    parse(s[:3]) or MONTH_MAP.get(s)

def safe_float(x):
    try:
        if x is None: return None
        s = str(x).strip()
        if s == "": return None
        return float(s)
    except:
        return None

def safe_date(v):
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
    return "Import Employees Page"
    
@app.route("/trainings/import", methods=["GET", "POST"])
def trainings_import():
    if request.method == "GET":
        return render_template("training_import.html")

    f = request.files.get("file")
    if not f or f.filename == "":
        flash("กรุณาเลือกไฟล์ Excel", "error")
        return redirect(url_for("trainings_import"))

    try:
        wb = load_workbook(f, data_only=True)
        ws = wb["Record Training"] if "Record Training" in wb.sheetnames else wb.active

        def norm(x):
            s = safe_str(x).strip().lower()
            for ch in ["\u00a0", ".", "-", "_", "/", "(", ")", "[", "]"]:
                s = s.replace(ch, " ")
            return " ".join(s.split())

        ALIASES = {
            "emp_id": ["emp id", "empid", "รหัสพนักงาน", "รหัส"],
            "prefix": ["คำนำหน้า", "prefix"],
            "frist_name": ["ชื่อ", "first name", "firstname", "first name", "ชื่อสกุล", "ชื่อ-สกุล"],
            "last_name": ["นามสกุล", "last name", "lastname"],
            "section ": ["แผนก", "section", "department"],
            "position": ["ตำแหน่ง", "position"],
            "course_code": ["รหัสหลักสูตร", "course code"],
            "course_name": ["ชื่อหลักสูตร", "course name"],
            "course_type": ["ประเภท", "category", "type", "course type"],

            "start_date": ["startdate", "start date", "วันที่เริ่ม", "เริ่ม"],
            "end_date": ["enddate", "end date", "วันที่จบ", "จบ"],
            "hours": ["ชั่วโมง", "hours", "hour"],

            "evaluate_method": ["วิธีประเมิน", "evaluate method"],
            "result": ["ผล", "result"],
            "score": ["คะแนน", "score"],
            "evaluator": ["ผู้ประเมิน", "evaluator"],
            "expire_date": ["วันหมดอายุ", "expire date"],
            "remark": ["หมายเหตุ", "remark"],

            "year": ["year", "ปี", "year."],
            "month": ["month", "เดือน", "mon"],
        }

        def find_header_row(scan_rows=10):
            best_row, best_score = 1, -1
            max_r = min(scan_rows, ws.max_row or 1)
            max_c = ws.max_column or 1
            for r in range(1, max_r + 1):
                vals = [norm(ws.cell(r, c).value) for c in range(1, max_c + 1)]
                score = 0
                for must in ["emp_id", "full_name", "course_name", "start_date"]:
                    if any(norm(a) in vals for a in ALIASES[must]):
                        score += 1
                if score > best_score:
                    best_score, best_row = score, r
            return best_row

        header_row = find_header_row()
        max_c = ws.max_column or 1

        headers = [norm(ws.cell(header_row, c).value) for c in range(1, max_c + 1)]
        header_map = {h: i + 1 for i, h in enumerate(headers) if h}

        def col(key):
            for alt in ALIASES.get(key, []):
                k = norm(alt)
                if k in header_map:
                    return header_map[k]
            return None

        def cellv(r, key):
            idx = col(key)
            if not idx:
                return None
            return ws.cell(r, idx).value

        added, skipped, duplicated = 0, 0, 0

        for r in range(header_row + 1, (ws.max_row or 1) + 1):
            emp_id = safe_str(cellv(r, "emp_id"))
            if not emp_id:
                skipped += 1
                continue

            prefix = safe_str(cellv(r, "prefix"))

            # ชื่อ / นามสกุล
            frist_name = safe_str(cellv(r, "frist_name"))
            last_name = safe_str(cellv(r, "last_name"))

            # แผนก/ตำแหน่ง
            section = safe_str(cellv(r, "section"))
            position = safe_str(cellv(r, "position"))

            course_code = safe_str(cellv(r, "course_code"))
            course_name = safe_str(cellv(r, "course_name"))
            course_type = safe_str(cellv(r, "course_type"))

            start_date = safe_date(cellv(r, "start_date"))
            end_date = safe_date(cellv(r, "end_date"))
            hours = safe_float(cellv(r, "hours"))

            evaluate_method = safe_str(cellv(r, "evaluate_method"))
            result = safe_str(cellv(r, "result"))
            score = safe_float(cellv(r, "score"))
            evaluator = safe_str(cellv(r, "evaluator"))
            expire_date = safe_date(cellv(r, "expire_date"))
            remark = safe_str(cellv(r, "remark"))

            year = safe_int(cellv(r, "year"))
            month = safe_month(cellv(r, "month"))

            # ✅ เงื่อนไขซ้ำ: emp_id + start_date เหมือนกัน = ซ้ำ
            if start_date:
                exists = (
                    TrainingRecord.query
                    .filter(TrainingRecord.emp_id == emp_id)
                    .filter(TrainingRecord.start_date == start_date)
                    .first()
                )
                if exists:
                    duplicated += 1
                    continue

            tr = TrainingRecord(
                year=year,
                month=month,
                emp_id=emp_id,
                prefix=prefix,
                first_name=first_name,
                last_name=last_name,
                section=section,
                position=position,
                course_code=course_code,
                course_name=course_name,
                course_type=course_type,
                start_date=start_date,
                end_date=end_date,
                hours=hours,
                evaluate_method=evaluate_method,
                result=result,
                score=score,
                evaluator=evaluator,
                expire_date=expire_date,
                remark=remark,
            )

            db.session.add(tr)
            added += 1

        db.session.commit()
        flash(f"Import สำเร็จ: เพิ่ม {added} | ซ้ำ {duplicated} | ข้าม {skipped}", "success")
        return redirect(url_for("trainings_list"))

    except Exception as e:
        db.session.rollback()
        flash(f"Import ไม่สำเร็จ: {e}", "error")
        return redirect(url_for("trainings_import"))

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

@app.route("/trainings/<int:tr_id>/edit", methods=["GET", "POST"])
def trainings_edit(tr_id):
    tr = TrainingRecord.query.get_or_404(tr_id)

    if request.method == "POST":
        tr.year = safe_int(request.form.get("year"))
        tr.month = safe_month(request.form.get("month"))

        tr.emp_id = safe_str(request.form.get("emp_id"))
        tr.prefix = safe_str(request.form.get("prefix"))
        tr.first_name = safe_str(request.form.get("first_name"))
        tr.last_name = safe_str(request.form.get("last_name"))

        tr.section = safe_str(request.form.get("section"))
        tr.position = safe_str(request.form.get("position"))

        tr.course_code = safe_str(request.form.get("course_code"))
        tr.course_name = safe_str(request.form.get("course_name"))
        tr.course_type = safe_str(request.form.get("course_type"))

        tr.start_date = safe_date(request.form.get("start_date"))
        tr.end_date = safe_date(request.form.get("end_date"))
        tr.hours = safe_float(request.form.get("hours"))

        tr.evaluate_method = safe_str(request.form.get("evaluate_method"))
        tr.result = safe_str(request.form.get("result"))
        tr.score = safe_float(request.form.get("score"))
        tr.evaluator = safe_str(request.form.get("evaluator"))

        tr.expire_date = safe_date(request.form.get("expire_date"))
        tr.remark = safe_str(request.form.get("remark"))

        db.session.commit()
        flash("แก้ไข Training Record เรียบร้อย", "success")
        return redirect(url_for("trainings_list"))

    return render_template("trainings_form.html", tr=tr, mode="edit")

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
        first_name=safe_str(request.form.get("first_name")),
        last_name=safe_str(request.form.get("last_name")),
        section=safe_str(request.form.get("section")),
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

@app.route("/trainings/<int:tr_id>/delete", methods=["POST"])
def trainings_delete(tr_id):
    tr = TrainingRecord.query.get_or_404(tr_id)
    db.session.delete(tr)
    db.session.commit()
    flash("ลบ Training Record แล้ว", "success")
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
