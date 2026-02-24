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

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[0]:
            continue

        em_id = str(row[0]).strip()

        if Employee.query.filter_by(em_id=em_id).first():
            continue

        emp = Employee(
            em_id=em_id,
            first_name_th=str(row[1]).strip() if row[1] else None,
            last_name_th=str(row[2]).strip() if row[2] else None,
        )

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
