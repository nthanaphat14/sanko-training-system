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
from sqlalchemy import or_

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
                Employee.department.ilike(like),
                Employee.status.ilike(like),
            )
        )

    employees = query.order_by(Employee.em_id.asc()).all()

    return render_template(
        "employees.html",
        employees=employees,
        q=q,
        total=len(employees),
    )


@app.route("/employees/new", methods=["GET", "POST"])
def employee_new():
    if request.method == "POST":
        em_id = safe_str(request.form.get("em_id"))

        if not em_id:
            flash("กรุณากรอก Em. ID", "error")
            return redirect(url_for("employee_new"))

        if Employee.query.filter_by(em_id=em_id).first():
            flash("Em. ID นี้มีอยู่แล้ว", "error")
            return redirect(url_for("employee_new"))

        emp = Employee(
            no=safe_int(request.form.get("no")),
            em_id=em_id,
            id_card=safe_str(request.form.get("id_card")),
            title_th=safe_str(request.form.get("title_th")),
            first_name_th=safe_str(request.form.get("first_name_th")),
            last_name_th=safe_str(request.form.get("last_name_th")),
            title_en=safe_str(request.form.get("title_en")),
            first_name_en=safe_str(request.form.get("first_name_en")),
            last_name_en=safe_str(request.form.get("last_name_en")),
            position=safe_str(request.form.get("position")),
            section=safe_str(request.form.get("section")),
            department=safe_str(request.form.get("department")),
            start_work=safe_date(request.form.get("start_work")),
            resign=safe_date(request.form.get("resign")),
            status=safe_str(request.form.get("status")),
            degree=safe_str(request.form.get("degree")),
            major=safe_str(request.form.get("major")),
        )

        db.session.add(emp)
        db.session.commit()

        flash("เพิ่มพนักงานเรียบร้อย ✅", "ok")
        return redirect(url_for("employees_list"))

    return render_template("employee_form.html", mode="new", emp=None)


@app.route("/employees/<em_id>/edit", methods=["GET", "POST"])
def employee_edit(em_id):
    emp = Employee.query.filter_by(em_id=em_id).first_or_404()

    if request.method == "POST":
        emp.no = safe_int(request.form.get("no"))
        emp.id_card = safe_str(request.form.get("id_card"))
        emp.title_th = safe_str(request.form.get("title_th"))
        emp.first_name_th = safe_str(request.form.get("first_name_th"))
        emp.last_name_th = safe_str(request.form.get("last_name_th"))
        emp.title_en = safe_str(request.form.get("title_en"))
        emp.first_name_en = safe_str(request.form.get("first_name_en"))
        emp.last_name_en = safe_str(request.form.get("last_name_en"))
        emp.position = safe_str(request.form.get("position"))
        emp.section = safe_str(request.form.get("section"))
        emp.department = safe_str(request.form.get("department"))
        emp.start_work = safe_date(request.form.get("start_work"))
        emp.resign = safe_date(request.form.get("resign"))
        emp.status = safe_str(request.form.get("status"))
        emp.degree = safe_str(request.form.get("degree"))
        emp.major = safe_str(request.form.get("major"))

        db.session.commit()
        flash("บันทึกการแก้ไขเรียบร้อย ✅", "ok")
        return redirect(url_for("employees_list"))

    return render_template("employee_form.html", mode="edit", emp=emp)


from flask import request, redirect, url_for, flash

@app.route("/employees/import", methods=["GET", "POST"])
def employees_import():

    # เปิดหน้า import
    if request.method == "GET":
        return render_template("import.html")
        
    # กดอัปโหลดไฟล์ (POST)
    f = request.files.get("file")
    if not f or f.filename == "":
        flash("กรุณาเลือกไฟล์ Excel ก่อน", "error")
        return redirect(url_for("employees_import"))

    flash("อัปโหลดไฟล์สำเร็จ ✅", "ok")
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
