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
from openpyxl import load_workbook

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

    return render_template("employee_form.html")
    
@app.route("/employees/<string:em_id>/edit", methods=["GET", "POST"])
def employee_edit(em_id):
    emp = Employee.query.filter_by(em_id=em_id).first_or_404()

    if request.method == "POST":
        emp.first_name_th = request.form.get("first_name_th")
        emp.last_name_th = request.form.get("last_name_th")
        emp.first_name_en = request.form.get("first_name_en")
        emp.last_name_en = request.form.get("last_name_en")
        emp.id_card = request.form.get("id_card")

        db.session.commit()
        flash("แก้ไขข้อมูลเรียบร้อย", "success")
        return redirect(url_for("employees_list"))

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
        flash("กรุณาเลือกไฟล์ Excel ก่อน", "error")
        return redirect(url_for("employees_import"))

    wb = load_workbook(f)
    ws = wb.active

    added = 0
updated = 0
skipped = 0

for row in ws.iter_rows(min_row=2, values_only=True):
    # ป้องกันแถวว่าง
    if not row:
        skipped += 1
        continue

    # --- mapping ตาม “ไฟล์ของคุณ” ---
    # สมมติคอลัมน์เรียงแบบ: No, Em.ID, ID Card, TitleTH, FirstTH, LastTH, Name-EN, Position, Section, Department, Start work, Resign, Status, Degree, Major
    # ถ้าของคุณไม่ตรง ให้บอก ผมจะจัด mapping ให้ตรง 100%
    no_raw = row[0] if len(row) > 0 else None
    em_id_raw = row[1] if len(row) > 1 else None

    if not em_id_raw:
        skipped += 1
        continue

    em_id = str(em_id_raw).strip()

    # ✅ upsert: ถ้ามีอยู่แล้วให้ update
    emp = Employee.query.filter_by(em_id=em_id).first()
    is_new = emp is None
    if is_new:
        emp = Employee(em_id=em_id)

    # --- no: รองรับ 1 / 1.0 / "001" ---
    try:
        if no_raw is not None and str(no_raw).strip() != "":
            emp.no = int(float(no_raw))
    except:
        pass

    # --- id_card ---
    if len(row) > 2 and row[2]:
        emp.id_card = str(row[2]).strip()

    # --- TH name ---
    if len(row) > 3 and row[3]:
        emp.title_th = str(row[3]).strip()
    if len(row) > 4 and row[4]:
        emp.first_name_th = str(row[4]).strip()
    if len(row) > 5 and row[5]:
        emp.last_name_th = str(row[5]).strip()

    # --- Name-EN (ชื่อเต็ม) -> แยก first/last แบบง่าย ---
    if len(row) > 6 and row[6]:
        name_en = str(row[6]).strip()
        parts = [p for p in name_en.split(" ") if p]
        if len(parts) >= 2:
            emp.first_name_en = " ".join(parts[:-1])
            emp.last_name_en = parts[-1]
        else:
            emp.first_name_en = name_en

    # --- Position / Section / Department ---
    if len(row) > 7 and row[7]:
        emp.position = str(row[7]).strip()
    if len(row) > 8 and row[8]:
        emp.section = str(row[8]).strip()
    if len(row) > 9 and row[9]:
        emp.department = str(row[9]).strip()

    # --- Start work / Resign / Status / Degree / Major ---
    if len(row) > 10 and row[10]:
        emp.start_work = str(row[10]).strip()
    if len(row) > 11 and row[11]:
        emp.resign = str(row[11]).strip()
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

        db.session.add(emp)
        added += 1

    db.session.commit()

    flash(f"Import สำเร็จ ✅ เพิ่ม {added} รายการ", "ok")
    return redirect(url_for("employees_list"))
@app.get("/employees/export")
def employees_export():
    return "Export logic here"

# -------------------------------------------------
# Run (Local Only)
# -------------------------------------------------
def init_db():
    with app.app_context():
        db.create_all()

init_db()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
