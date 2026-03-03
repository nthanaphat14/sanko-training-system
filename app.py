import os
from datetime import datetime
from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    flash,
    session,
    g,
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
from functools import wraps
from flask import session, abort
from flask import session, g
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import timedelta
from flask import redirect, url_for, abort, flash
from flask import request
from werkzeug.security import check_password_hash
from sqlalchemy import or_, func
from sqlalchemy import or_
from math import ceil
from datetime import datetime
from openpyxl import Workbook

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
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
# ถ้ารันบน HTTPS (Render) ให้เปิด True ได้เลย
app.config["SESSION_COOKIE_SECURE"] = True
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(minutes=30)

print("DATABASE =", app.config["SQLALCHEMY_DATABASE_URI"])  # 👈 ใส่ตรงนี้

db = SQLAlchemy(app)


# -------------------------------------------------
# Model
# -------------------------------------------------
class User(db.Model):
    __tablename__ = "users"
    id = db.Column(db.Integer, primary_key=True)
    email = db.Column(db.String(180), unique=True, nullable=False, index=True)
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(20), nullable=False, default="viewer")  # admin / viewer
    is_active = db.Column(db.Boolean, default=True)

    failed_attempts = db.Column(db.Integer, default=0)
    locked_until = db.Column(db.DateTime, nullable=True)

    last_login_at = db.Column(db.DateTime, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

class AuditLog(db.Model):
    __tablename__ = "audit_logs"
    id = db.Column(db.Integer, primary_key=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, index=True)

    user_email = db.Column(db.String(255), nullable=True, index=True)
    action = db.Column(db.String(80), nullable=False, index=True)
    detail = db.Column(db.Text, nullable=True)
    ip = db.Column(db.String(64), nullable=True)

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

class ImportBatch(db.Model):
    __tablename__ = "import_batches"
    id = db.Column(db.Integer, primary_key=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    # สรุปผลรอบนี้
    added = db.Column(db.Integer, default=0)
    updated = db.Column(db.Integer, default=0)
    duplicated = db.Column(db.Integer, default=0)
    skipped = db.Column(db.Integer, default=0)

    # ไว้ให้ดูไฟล์อะไร import (optional)
    filename = db.Column(db.String(255), nullable=True)

class ImportItem(db.Model):
    __tablename__ = "import_items"
    id = db.Column(db.Integer, primary_key=True)
    batch_id = db.Column(db.Integer, db.ForeignKey("import_batches.id"), nullable=False, index=True)

    # Added / Updated / Duplicate / Skipped
    status = db.Column(db.String(20), nullable=False)
    reason = db.Column(db.String(255), nullable=True)

    # เก็บ snapshot สำหรับ “ดูว่าใคร”
    row_no = db.Column(db.Integer, nullable=True)
    emp_id = db.Column(db.String(50), nullable=True)
    prefix = db.Column(db.String(50), nullable=True)
    first_name = db.Column(db.String(200), nullable=True)
    last_name = db.Column(db.String(200), nullable=True)
    section = db.Column(db.String(150), nullable=True)
    position = db.Column(db.String(150), nullable=True)

    course_code = db.Column(db.String(100), nullable=True)
    course_name = db.Column(db.String(255), nullable=True)
    course_type = db.Column(db.String(100), nullable=True)

    start_date = db.Column(db.Date, nullable=True)
    end_date = db.Column(db.Date, nullable=True)

    created_at = db.Column(db.DateTime, default=datetime.utcnow)

# -------------------------------------------------
# Helper Functions
# -------------------------------------------------
def get_current_user():
    uid = session.get("uid")
    if not uid:
        return None
    return db.session.get(User, uid)  # SQLAlchemy 2.x

def build_training_query(args):
    """
    ใช้ args = request.args (หรือ dict ที่มี key เหมือนกัน)
    ทำให้ /trainings และ /trainings/export ใช้ filter ชุดเดียวกัน 100%
    """
    q = (args.get("q") or "").strip()
    year = (args.get("year") or "").strip()
    month = (args.get("month") or "").strip()

    query = TrainingRecord.query

    if q:
        like = f"%{q}%"

        # เลือก field ชื่อแบบยืดหยุ่น (กัน model ไม่ตรง)
        name_field = None
        for cand in ["full_name", "employee_name", "name", "emp_name", "first_name", "last_name"]:
            if hasattr(TrainingRecord, cand):
                name_field = getattr(TrainingRecord, cand)
                break

        conds = []

        # ✅ สำคัญ: ใช้ชื่อ field ให้ตรงกับ model ของคุณ
        # จาก template คุณใช้ t.emp_id / t.course_code / t.course_name
        if hasattr(TrainingRecord, "emp_id"):
            conds.append(TrainingRecord.emp_id.ilike(like))
        if hasattr(TrainingRecord, "employee_code"):
            conds.append(TrainingRecord.employee_code.ilike(like))

        # ค้นหาชื่อ/นามสกุล ถ้ามี
        if hasattr(TrainingRecord, "first_name"):
            conds.append(TrainingRecord.first_name.ilike(like))
        if hasattr(TrainingRecord, "last_name"):
            conds.append(TrainingRecord.last_name.ilike(like))

        # หรือชื่อรวม (ถ้ามี)
        if name_field is not None:
            conds.append(name_field.ilike(like))

        if hasattr(TrainingRecord, "course_code"):
            conds.append(TrainingRecord.course_code.ilike(like))
        if hasattr(TrainingRecord, "course_name"):
            conds.append(TrainingRecord.course_name.ilike(like))

        if conds:
            query = query.filter(or_(*conds))

    if year.isdigit() and hasattr(TrainingRecord, "year"):
        query = query.filter(TrainingRecord.year == int(year))
    if month.isdigit() and hasattr(TrainingRecord, "month"):
        query = query.filter(TrainingRecord.month == int(month))

    return query, q, year, month

@app.context_processor
def inject_helpers():
    return {
        "current_user": get_current_user(),
        "mask_id_card": mask_id_card,
        "format_id_card": format_id_card,
    }

def format_id_card(id_card: str):
    if not id_card or len(id_card) != 13:
        return ""
    return f"{id_card[0]}-{id_card[1:5]}-{id_card[5:10]}-{id_card[10:12]}-{id_card[12]}"

def mask_id_card(id_card):
    if not id_card:
        return ""
    return id_card[:6] + "XXXXX" + id_card[-3:]

app.jinja_env.globals.update(mask_id_card=mask_id_card)

def audit(action, detail=None, user_email=None):
    try:
        ip = (request.headers.get("X-Forwarded-For") or request.remote_addr or "")
        ip = ip.split(",")[0].strip() if ip else None

        if user_email is None:
            u = get_current_user()  # ถ้ามี session อยู่จะได้ email
            user_email = getattr(u, "email", None) if u else None

        db.session.add(AuditLog(
            action=str(action),
            detail=detail,
            user_email=user_email,
            ip=ip,
        ))
        db.session.commit()
    except Exception:
        db.session.rollback()
        # ห้ามให้ audit ทำให้ระบบล่ม
        pass

@app.before_request
def require_login_globally():
    # บางครั้ง endpoint เป็น None (เช่น 404) กันพังไว้
    if request.endpoint is None:
        return

    allow = {"login", "login_post", "logout", "static"}

    if request.endpoint in allow:
        return

    u = get_current_user()
    if not u or not u.is_active:
        return redirect(url_for("login"))
        
def role_required(*roles):
    def deco(fn):
        @wraps(fn)
        def wrapper(*args, **kwargs):
            u = get_current_user()
            if not u or not u.is_active:
                return redirect(url_for("login"))

            if roles:
                if getattr(u, "role", None) not in roles:
                    abort(403)

            return fn(*args, **kwargs)
        return wrapper
    return deco

def seed_users_if_missing():
    defaults = [
        ("hr02@sankothai.net", "Sanko1996", "admin"),
        ("hr@sankothai.net", "Sanko1996", "viewer"),
        ("hr01@sankothai.net", "Sanko1996", "viewer"),
    ]

    for email, pw, role in defaults:
        exists = User.query.filter_by(email=email).first()
        if not exists:
            db.session.add(User(
                email=email,
                password_hash=generate_password_hash(pw),
                role=role,
                is_active=True
            ))

    db.session.commit()

def login_required(fn):
    # ใช้ role_required แบบไม่กำหนด role (แค่ต้อง login)
    return role_required()(fn)
    
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
@app.get("/login")
def login():
    u = get_current_user()
    if u and u.is_active:
        return redirect(url_for("employees_list"))
    return render_template("login.html")


@app.post("/login")
def login_post():
    email = (request.form.get("email") or "").strip().lower()
    password = request.form.get("password") or ""

    u = User.query.filter_by(email=email).first()
    if not u or not u.is_active:
        audit("LOGIN_FAIL", f"email={email}")
        flash("User หรือ Password ไม่ถูกต้อง", "error")
        return redirect(url_for("login"))

    if u.locked_until and u.locked_until > datetime.utcnow():
        audit("LOGIN_LOCKED", f"email={email}")
        flash("บัญชีถูกล็อกชั่วคราว (ลองใหม่อีกครั้งภายหลัง)", "error")
        return redirect(url_for("login"))

    if not check_password_hash(u.password_hash, password):
        u.failed_attempts = (u.failed_attempts or 0) + 1
        if u.failed_attempts >= 8:
            u.locked_until = datetime.utcnow() + timedelta(minutes=10)
            u.failed_attempts = 0
        db.session.commit()

        audit("LOGIN_FAIL", f"email={email}")
        flash("User หรือ Password ไม่ถูกต้อง", "error")
        return redirect(url_for("login"))

    # success
    u.failed_attempts = 0
    u.locked_until = None
    u.last_login_at = datetime.utcnow()
    db.session.commit()

    session.clear()
    session["uid"] = u.id          # ✅ สำคัญ: ใช้ key ให้ตรงกับ get_current_user()
    session.permanent = True

    audit("LOGIN_SUCCESS", f"email={email}")
    flash("เข้าสู่ระบบสำเร็จ", "success")
    return redirect(url_for("employees_list"))

@app.get("/logout")
def logout():
    audit("LOGOUT", "")
    session.clear()
    return redirect(url_for("login"))

@app.route("/change-password", methods=["GET", "POST"])
@login_required
def change_password():
    u = get_current_user()
    if request.method == "GET":
        return render_template("change_password.html", user=u)

    old_pw = request.form.get("old_password") or ""
    new_pw = request.form.get("new_password") or ""
    new_pw2 = request.form.get("new_password2") or ""

    if not check_password_hash(u.password_hash, old_pw):
        flash("รหัสเดิมไม่ถูกต้อง", "error")
        return redirect(url_for("change_password"))

    if new_pw != new_pw2:
        flash("ยืนยันรหัสใหม่ไม่ตรงกัน", "error")
        return redirect(url_for("change_password"))

    if len(new_pw) < 10:
        flash("รหัสใหม่ต้องยาวอย่างน้อย 10 ตัวอักษร", "error")
        return redirect(url_for("change_password"))

    u.password_hash = generate_password_hash(new_pw)
    db.session.commit()
    audit("CHANGE_PASSWORD", f"user={u.email}")
    flash("เปลี่ยนรหัสผ่านสำเร็จ", "success")
    return redirect(url_for("employees_list"))


# ---------- Admin: ดูรายชื่อ user + reset ----------
@app.get("/admin/users")
@role_required("admin")
def admin_users():
    users = User.query.order_by(User.email.asc()).all()
    return render_template("admin_users.html", users=users)


@app.post("/admin/users/<int:user_id>/reset-password")
@role_required("admin")
def admin_reset_password(user_id):
    target = User.query.get_or_404(user_id)
    new_pw = request.form.get("new_password") or ""

    if len(new_pw) < 10:
        flash("รหัสใหม่ต้องยาวอย่างน้อย 10 ตัวอักษร", "error")
        return redirect(url_for("admin_users"))

    target.password_hash = generate_password_hash(new_pw)
    db.session.commit()
    audit("RESET_PASSWORD", f"admin={session.get('user_email')} reset={target.email}")
    flash(f"รีเซ็ตรหัสผ่านให้ {target.email} สำเร็จ", "success")
    return redirect(url_for("admin_users"))

@app.before_request
def require_login_globally():
    open_paths = set(["/login", "/healthz"])
    if request.path.startswith("/static/"):
        return None
    if request.path in open_paths:
        return None

    # ปล่อยให้ POST /login ผ่าน
    if request.path == "/login" and request.method == "POST":
        return None

    u = get_current_user()
    if not u or not u.is_active:
        return redirect(url_for("login"))

@app.get("/")
def root():
    return redirect(url_for("employees_list"))


@app.get("/healthz")
def healthz():
    return {"status": "ok"}, 200


@app.get("/employees")
@login_required
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

@app.post("/employees/bulk-delete")
@role_required("admin")
def employees_bulk_delete():
    ids = request.form.getlist("ids")

    if not ids:
        flash("ยังไม่ได้เลือกพนักงาน", "error")
        return redirect(url_for("employees_list"))

    Employee.query.filter(Employee.em_id.in_(ids)).delete(synchronize_session=False)
    db.session.commit()

    audit("EMPLOYEE_BULK_DELETE", f"count={len(ids)}")
    flash(f"ลบแล้ว {len(ids)} รายการ", "success")
    return redirect(url_for("employees_list"))

@app.route("/employees/import", methods=["GET", "POST"])
@role_required("admin")
def employees_import():
    if request.method == "GET":
        return render_template("employees_import.html")

    f = request.files.get("file")
    if not f:
        flash("กรุณาเลือกไฟล์ .xlsx", "error")
        return redirect(url_for("employees_import"))

    try:
        wb = load_workbook(f, data_only=True)

        added = 0
        updated = 0
        skipped = 0

        for ws in wb.worksheets:
            sheet_title = (ws.title or "").strip().lower()
            default_status = "Resign" if ("ลาออก" in sheet_title or "resign" in sheet_title) else "Active"

            # อ่านหัวตาราง (row 1) ของชีตนี้
            # หาแถวหัวตารางที่มีคำว่า Em. ID
            header_row = None
            for r in range(1, 30):
                values = [(str(c.value).strip() if c.value else "") for c in ws[r]]
                if any("em. id" in v.lower() or "em id" in v.lower() for v in values):
                    header_row = r
                    break

            if not header_row:
                flash("ไม่พบแถวหัวตาราง (Em. ID)", "error")
                return redirect(url_for("employees_import"))

            # อ่านหัวตารางจริง
            headers = []
            for c in ws[header_row]:
                headers.append((str(c.value).strip() if c.value else ""))
                

            def col(name):
                name = name.strip().lower()
                for i, h in enumerate(headers):
                    if (h or "").strip().lower() == name:
                        return i
                return None

            # คอลัมน์ตามไฟล์คุณ
            c_no = col("no.")
            c_em = col("em. id") or col("em id") or col("employee id")
            c_idcard = col("id card")
            c_title_th = col("Prefix")
            c_first_th = col("name-TH")
            c_last_th = col("last-TH")
            c_name_en = col("name-en") or col("name-en ")
            c_position = col("position")
            c_section = col("section")
            c_dept = col("department")
            c_start = col("start work") or col("วันเริ่มงาน ")
            c_resign = col("resign")
            c_status = col("status")
            c_degree = col("Education")
            c_major = col("major")

            if c_em is None:
                # ถ้าบางชีตไม่ใช่ข้อมูลพนักงาน ให้ข้ามชีตนั้นไป
                continue

            # วนอ่านตั้งแต่แถว 2
            for row in ws.iter_rows(min_row=header_row+1, values_only=True):
                em_id = safe_str(row[c_em]) if c_em is not None else ""
                em_id = em_id.strip()
                
                import re

                # รองรับรหัสขึ้นต้นด้วย M หรือ D เท่านั้น
                if not re.match(r"^[MD]\d{5,10}$", em_id, re.IGNORECASE):
                    continue

                if not em_id:
                    skipped += 1
                    continue

                emp = Employee.query.filter_by(em_id=em_id).first()
                if not emp:
                    emp = Employee(em_id=em_id)
                    db.session.add(emp)
                    added += 1
                else:
                    updated += 1

                name_th = safe_str(row[c_name_th]) if c_name_th is not None else ""
                name_en = safe_str(row[c_name_en]) if c_name_en is not None else ""

                emp.no = safe_int(row[c_no]) if c_no is not None else None
                emp.id_card = safe_str(row[c_idcard]) if c_idcard is not None else ""
                emp.title_th = prefix
                emp.first_name_th = name-th
                emp.last_name_th = last-th
                emp.first_name_en = name_en
                emp.position = safe_str(row[c_position]) if c_position is not None else ""
                emp.section = safe_str(row[c_section]) if c_section is not None else ""
                emp.department = safe_str(row[c_dept]) if c_dept is not None else ""
                emp.start_work = safe_date(row[c_start]) if c_start is not None else None
                emp.resign = safe_date(row[c_resign]) if c_resign is not None else None

                file_status = safe_str(row[c_status]) if c_status is not None else ""
                emp.status = file_status or default_status

                emp.degree = safe_str(row[c_degree]) if c_degree is not None else ""
                emp.major = safe_str(row[c_major]) if c_major is not None else ""

        db.session.commit()
        flash(f"Import Employees สำเร็จ: เพิ่มใหม่ {added} | อัปเดต {updated} | ข้าม {skipped}", "success")
        return redirect(url_for("employees_list"))

    except Exception as e:
        db.session.rollback()
        flash(f"Import Employees ล้มเหลว: {e}", "error")
        return redirect(url_for("employees_import"))
    
@app.route("/trainings/import", methods=["GET", "POST"])
@role_required("admin")
def trainings_import():
    if request.method == "GET":
        return render_template("training_import.html")

    f = request.files.get("file")
    if not f or f.filename == "":
        flash("กรุณาเลือกไฟล์ Excel", "error")
        return redirect(url_for("trainings_import"))

    # ---- สร้าง Batch รอบนี้ ----
    batch = ImportBatch(filename=f.filename)
    db.session.add(batch)
    db.session.commit()  # ต้อง commit ก่อนเพื่อได้ batch.id

    def log_item(status, reason=None, row_no=None, **kw):
        it = ImportItem(
            batch_id=batch.id,
            status=status,
            reason=reason,
            row_no=row_no,
            emp_id=kw.get("emp_id"),
            prefix=kw.get("prefix"),
            first_name=kw.get("first_name"),
            last_name=kw.get("last_name"),
            section=kw.get("section"),
            position=kw.get("position"),
            course_code=kw.get("course_code"),
            course_name=kw.get("course_name"),
            course_type=kw.get("course_type"),
            start_date=kw.get("start_date"),
            end_date=kw.get("end_date"),
        )
        db.session.add(it)

    try:
        wb = load_workbook(f, data_only=True)
        ws = wb["Record Training"] if "Record Training" in wb.sheetnames else wb.active

        # ======= NORMALIZE HEADER =======
        def norm(x):
            s = safe_str(x).strip().lower()
            for ch in ["\u00a0", ".", "-", "_", "/", "(", ")", "[", "]", ":"]:
                s = s.replace(ch, " ")
            return " ".join(s.split())

        # ✅ เพิ่ม alias ให้ตรงไฟล์คุณ (ชื่อไทย / name-th ฯลฯ)
        ALIASES = {
            "seq": ["ลำดับ", "no", "seq", "#"],
            "year": ["year", "year.", "ปี"],
            "month": ["month", "mon", "เดือน"],

            "emp_id": ["em id", "emp id", "empid", "รหัสพนักงาน", "รหัส", "employee id"],
            "prefix": ["คำนำหน้า", "prefix"],
            "first_name": ["ชื่อ", "ชื่อไทย", "name th", "thai name", "first name", "firstname"],
            "last_name": ["นามสกุล", "last name", "lastname"],
            "section": ["section", "แผนก", "ฝ่าย", "หน่วยงาน", "department"],
            "position": ["position", "ตำแหน่ง"],

            "course_code": ["course code", "coursecode", "รหัสหลักสูตร"],
            "course_name": ["course name", "coursename", "ชื่อหลักสูตร"],
            "course_type": ["course type", "type", "category", "ประเภท"],

            "start_date": ["startdate", "start date", "วันที่เริ่ม", "เริ่ม"],
            "end_date": ["enddate", "end date", "วันที่จบ", "จบ"],
            "hours": ["ชั่วโมง", "hours", "hour"],

            "evaluate_method": ["วิธีประเมิน", "evaluate method", "ประเมิน"],
            "result": ["ผล", "result"],
            "score": ["คะแนน", "score"],
            "evaluator": ["ผู้ประเมิน", "evaluator"],
            "expire_date": ["วันหมดอายุ", "expire date", "expiry"],
            "remark": ["หมายเหตุ", "remark", "note"],
        }

        def find_header_row(scan_rows=10):
            best_row, best_score = 1, -1
            max_r = min(scan_rows, ws.max_row or 1)
            max_c = ws.max_column or 1
            must_keys = ["emp_id", "course_code", "course_name", "start_date"]

            for r in range(1, max_r + 1):
                vals = [norm(ws.cell(r, c).value) for c in range(1, max_c + 1)]
                score = 0
                for k in must_keys:
                    if any(norm(a) in vals for a in ALIASES[k]):
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

        # ======= counters =======
        added = updated = duplicated = skipped = 0

        # ======= loop data =======
        for r in range(header_row + 1, (ws.max_row or 1) + 1):
            emp_id = safe_str(cellv(r, "emp_id"))
            course_code = safe_str(cellv(r, "course_code"))
            start_date = safe_date(cellv(r, "start_date"))
            end_date = safe_date(cellv(r, "end_date"))

            prefix = safe_str(cellv(r, "prefix"))
            first_name = safe_str(cellv(r, "first_name"))
            last_name = safe_str(cellv(r, "last_name"))
            section = safe_str(cellv(r, "section"))
            position = safe_str(cellv(r, "position"))

            year = safe_int(cellv(r, "year"))
            month = safe_month(cellv(r, "month"))

            course_name = safe_str(cellv(r, "course_name"))
            course_type = safe_str(cellv(r, "course_type"))
            hours = safe_float(cellv(r, "hours"))
            evaluate_method = safe_str(cellv(r, "evaluate_method"))
            result = safe_str(cellv(r, "result"))
            score = safe_float(cellv(r, "score"))
            evaluator = safe_str(cellv(r, "evaluator"))
            expire_date = safe_date(cellv(r, "expire_date"))
            remark = safe_str(cellv(r, "remark"))

            # ---- ถ้าแถวว่างมาก ๆ ให้ข้ามแบบเงียบ ----
            if not emp_id and not course_code and not course_name and not start_date:
                continue

            # ---- SKIPPED: ข้อมูลขั้นต่ำไม่ครบ ----
            if not emp_id or not start_date or not end_date or not course_code:
                skipped += 1
                log_item(
                    "Skipped",
                    reason="ข้อมูลขั้นต่ำไม่ครบ (ต้องมี Emp ID + StartDate + EndDate + Course code)",
                    row_no=r,
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
                )
                continue

            # ======================================================
            # KEY “ตัวเดิม” = emp_id + start_date + end_date + course_code
            # - ถ้า key นี้ไม่เคยมี → Added
            # - ถ้าเคยมี → พยายามเติมช่องว่าง (Updated) ไม่งั้น Duplicate
            # ======================================================
            existing = (
                TrainingRecord.query
                .filter(TrainingRecord.emp_id == emp_id)
                .filter(TrainingRecord.start_date == start_date)
                .filter(TrainingRecord.end_date == end_date)
                .filter(TrainingRecord.course_code == course_code)
                .first()
            )

            if existing is None:
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
                log_item(
                    "Added",
                    row_no=r,
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
                )
                continue

            # ---- มีตัวเดิมแล้ว: UPDATE เฉพาะช่องที่เดิมว่าง + ใหม่มีค่า ----
            updated_flag = False

            def fill_if_empty(field, new_value):
                nonlocal updated_flag
                old = getattr(existing, field)
                old_empty = (old is None) or (isinstance(old, str) and old.strip() == "")
                new_ok = (new_value is not None) and (not (isinstance(new_value, str) and new_value.strip() == ""))
                if old_empty and new_ok:
                    setattr(existing, field, new_value)
                    updated_flag = True
                    
            fill_if_empty("course_name", course_name)
            fill_if_empty("course_type", course_type)
            fill_if_empty("hours", hours)
            fill_if_empty("evaluate_method", evaluate_method)
            fill_if_empty("result", result)
            fill_if_empty("score", score)
            fill_if_empty("evaluator", evaluator)
            fill_if_empty("expire_date", expire_date)
            fill_if_empty("remark", remark)

            # ✅ แนะนำ: เติม prefix/name/section/position ได้ด้วย (เพราะตอนนี้ของคุณ first_name/section ว่างอยู่)
            fill_if_empty("prefix", prefix)
            fill_if_empty("first_name", first_name)
            fill_if_empty("last_name", last_name)
            fill_if_empty("section", section)
            fill_if_empty("position", position)
            fill_if_empty("year", year)
            fill_if_empty("month", month)

            if updated_flag:
                updated += 1
                log_item(
                    "Updated",
                    row_no=r,
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
                )
            else:
                duplicated += 1
                log_item(
                    "Duplicate",
                    reason="พบรายการเดิมแล้ว และไม่มีข้อมูลใหม่เพื่อเติม",
                    row_no=r,
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
                )

        # ---- บันทึกผลลง batch ----
        batch.added = added
        batch.updated = updated
        batch.duplicated = duplicated
        batch.skipped = skipped

        db.session.commit()

        flash(
            f"Import สำเร็จ: เพิ่ม {added} | อัปเดต {updated} | ซ้ำ {duplicated} | ข้าม {skipped}",
            "success"
        )
        return redirect(url_for("import_batch_detail", batch_id=batch.id))

    except Exception as e:
        db.session.rollback()
        flash(f"Import ไม่สำเร็จ: {e}", "error")
        return redirect(url_for("trainings_import"))


# =========================================================
# IMPORT HISTORY LIST
# =========================================================
@app.get("/training-imports")
def import_batches_list():
    batches = ImportBatch.query.order_by(ImportBatch.id.desc()).limit(50).all()
    return render_template("training-imports_list.html", batches=batches)


# =========================================================
# IMPORT HISTORY DETAIL
# =========================================================
@app.get("/training-imports/<int:batch_id>")
def import_batch_detail(batch_id):
    batch = ImportBatch.query.get_or_404(batch_id)

    added = ImportItem.query.filter_by(batch_id=batch_id, status="Added").order_by(ImportItem.id.asc()).all()
    updated = ImportItem.query.filter_by(batch_id=batch_id, status="Updated").order_by(ImportItem.id.asc()).all()
    duplicated = ImportItem.query.filter_by(batch_id=batch_id, status="Duplicate").order_by(ImportItem.id.asc()).all()
    skipped = ImportItem.query.filter_by(batch_id=batch_id, status="Skipped").order_by(ImportItem.id.asc()).all()

    return render_template(
        "training-imports_detail.html",
        batch=batch,
        added=added,
        updated=updated,
        duplicated=duplicated,
        skipped=skipped,
    )                    

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
    query, q, year, month = build_training_query(request.args)

    # pagination
    page = int(request.args.get("page", 1) or 1)
    per_page = int(request.args.get("per_page", 50) or 50)
    if per_page not in (25, 50, 100, 200):
        per_page = 50
    if page < 1:
        page = 1

    total = query.count()
    total_pages = max(1, (total + per_page - 1) // per_page)

    rows = (query
        .order_by(TrainingRecord.start_date.desc().nullslast(), TrainingRecord.id.desc())
        .offset((page - 1) * per_page)
        .limit(per_page)
        .all()
    )

    return render_template(
        "trainings_list.html",
        rows=rows,
        total=total,
        q=q, year=year, month=month,
        page=page,
        per_page=per_page,
        total_pages=total_pages,
    )

@app.get("/trainings/export")
@role_required("admin", "viewer")  # ✅ viewer export ได้
def trainings_export():
    query, q, year, month = build_training_query(request.args)

    # กัน export เยอะเกินจนล่ม (ปรับเลขได้)
    total = query.count()
    if total > 50000:
        flash(f"ผลลัพธ์ {total:,} รายการมากเกินไป กรุณากรองเพิ่มก่อน Export", "error")
        return redirect(url_for("trainings_list", q=q, year=year, month=month))

    # log สำหรับ audit
    audit("EXPORT_TRAININGS", f"q={q}&year={year}&month={month}&total={total}")

    # ใช้ write_only ลด memory
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("TrainingRecord")

    headers = [
        "Year","Month","Emp ID","Prefix","First Name","Last Name","Section","Position",
        "Course Code","Course Name","Type","StartDate","EndDate","Hours",
        "Eval Method","Result","Score","Evaluator","Expire Date","Remark"
    ]
    ws.append(headers)

    def val(x):
        if x is None:
            return ""
        # date/datetime -> string
        try:
            if hasattr(x, "strftime"):
                return x.strftime("%Y-%m-%d")
        except Exception:
            pass
        return str(x)

    # ดึงทีละ batch ลด RAM
    # (yield_per ใช้ได้กับหลาย DB; ถ้าเจอปัญหา ค่อยปรับ)
    for t in query.order_by(TrainingRecord.id.desc()).yield_per(1000):
        ws.append([
            val(getattr(t, "year", "")),
            val(getattr(t, "month", "")),
            val(getattr(t, "emp_id", getattr(t, "employee_code", ""))),
            val(getattr(t, "prefix", "")),
            val(getattr(t, "first_name", "")),
            val(getattr(t, "last_name", "")),
            val(getattr(t, "section", "")),
            val(getattr(t, "position", "")),
            val(getattr(t, "course_code", "")),
            val(getattr(t, "course_name", "")),
            val(getattr(t, "course_type", "")),
            val(getattr(t, "start_date", "")),
            val(getattr(t, "end_date", "")),
            val(getattr(t, "hours", "")),
            val(getattr(t, "evaluate_method", "")),
            val(getattr(t, "result", "")),
            val(getattr(t, "score", "")),
            val(getattr(t, "evaluator", "")),
            val(getattr(t, "expire_date", "")),
            val(getattr(t, "remark", "")),
        ])

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)

    stamp = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    filename = f"training_record_{stamp}.xlsx"

    return send_file(
        bio,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

@app.route("/trainings/<int:tr_id>/edit", methods=["GET", "POST"])
@role_required("admin")
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

    return render_template("trainings_edit.html", tr=tr)

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

@app.route("/trainings/bulk-delete", methods=["POST"])
@role_required("admin")
def trainings_bulk_delete():

    print("BULK DELETE IDS:", request.form.getlist("ids"))
    
    ids = request.form.getlist("ids")  # ✅ ต้อง getlist เท่านั้น

    if not ids:
        flash("ยังไม่ได้เลือกข้อมูลที่จะลบ", "error")
        return redirect(url_for("trainings_list"))

    try:
        # แปลงเป็น int กันค่ามั่ว
        ids_int = []
        for x in ids:
            try:
                ids_int.append(int(x))
            except:
                pass

        if not ids_int:
            flash("รายการที่เลือกไม่ถูกต้อง", "error")
            return redirect(url_for("trainings_list"))

        deleted = (
            TrainingRecord.query
            .filter(TrainingRecord.id.in_(ids_int))
            .delete(synchronize_session=False)
        )

        db.session.commit()
        flash(f"ลบสำเร็จ {deleted} รายการ", "success")
        return redirect(url_for("trainings_list"))

    except Exception as e:
        db.session.rollback()
        flash(f"ลบไม่สำเร็จ: {e}", "error")
        return redirect(url_for("trainings_list"))

@app.route("/trainings/<int:tr_id>/delete", methods=["POST"])
@role_required("admin")
def trainings_delete(tr_id):
    tr = TrainingRecord.query.get_or_404(tr_id)
    try:
        db.session.delete(tr)
        db.session.commit()
        flash("ลบรายการแล้ว", "success")
    except Exception as e:
        db.session.rollback()
        flash(f"ลบไม่สำเร็จ: {e}", "error")
    return redirect(url_for("trainings_list"))

# -------------------------------------------------
# Run (Local Only)
# -------------------------------------------------
    with app.app_context():
        db.create_all()
        seed_users_if_missing()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
