import os
import re
from io import BytesIO
from datetime import datetime, date

import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, flash, send_file
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import or_

# -------------------------
# App config
# -------------------------
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret")

db_url = (os.environ.get("DATABASE_URL") or "").strip()
if db_url.startswith("postgres://"):
    db_url = db_url.replace("postgres://", "postgresql://", 1)

app.config["SQLALCHEMY_DATABASE_URI"] = db_url or "sqlite:///employee.db"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
db = SQLAlchemy(app)


# -------------------------
# Model
# -------------------------
class Employee(db.Model):
    __tablename__ = "employees"

    id = db.Column(db.Integer, primary_key=True)

    no = db.Column(db.Integer, nullable=True)              # No.
    em_id = db.Column(db.String(40), unique=True, nullable=False)  # Em. ID
    id_card = db.Column(db.String(60), nullable=True)      # ID Card

    # TH
    title_th = db.Column(db.String(50), nullable=True)
    first_name_th = db.Column(db.String(120), nullable=True)
    last_name_th = db.Column(db.String(120), nullable=True)

    # EN
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
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    def th_full(self):
        parts = [self.title_th, self.first_name_th, self.last_name_th]
        return " ".join([p for p in parts if p])

    def en_full(self):
        parts = [self.title_en, self.first_name_en, self.last_name_en]
        return " ".join([p for p in parts if p])


with app.app_context():
    db.create_all()


# -------------------------
# Helpers
# -------------------------
TH_TITLES = {"นาย", "นาง", "นางสาว", "ด.ช.", "ด.ญ.", "คุณ"}
EN_TITLES = {"mr", "mrs", "ms"}

def _is_nan(x):
    try:
        return pd.isna(x)
    except Exception:
        return False

def s(x):
    """safe string (NaN/empty => None)"""
    if x is None or _is_nan(x):
        return None
    if isinstance(x, str):
        t = x.strip()
        return t if t else None
    t = str(x).strip()
    return t if t else None

def i(x):
    """safe int (NaN/empty/float => int or None)"""
    if x is None or _is_nan(x):
        return None
    try:
        if isinstance(x, str) and x.strip() == "":
            return None
        return int(float(x))
    except Exception:
        return None

def d(x):
    """safe date (excel date / string date)"""
    if x is None or _is_nan(x):
        return None
    try:
        dt = pd.to_datetime(x, errors="coerce")
        if pd.isna(dt):
            return None
        return dt.date()
    except Exception:
        return None

def split_th(full):
    """
    รับชื่อไทยรวม เช่น 'นาย สมชาย ใจดี' หรือ 'สมชาย ใจดี'
    """
    full = s(full)
    if not full:
        return None, None, None
    parts = re.split(r"\s+", full)
    if len(parts) == 1:
        return None, parts[0], None
    if parts[0] in TH_TITLES:
        title = parts[0]
        first = parts[1] if len(parts) >= 2 else None
        last = " ".join(parts[2:]) if len(parts) >= 3 else None
        return title, first, last
    return None, parts[0], " ".join(parts[1:])

def split_en(full):
    """
    รับชื่ออังกฤษรวม เช่น 'Mr.Masami Katsumoto' หรือ 'Masami Katsumoto'
    """
    full = s(full)
    if not full:
        return None, None, None

    # แก้เคส Mr.Masami -> Mr. Masami
    full = re.sub(r"(?i)\b(Mr|Mrs|Ms)\.", r"\1. ", full).strip()
    parts = re.split(r"\s+", full)

    title = None
    if parts:
        p0 = parts[0].lower().rstrip(".")
        if p0 in EN_TITLES:
            title = p0.capitalize()
            parts = parts[1:]

    if len(parts) == 0:
        return title, None, None
    if len(parts) == 1:
        return title, parts[0], None
    return title, parts[0], " ".join(parts[1:])

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    map หัวคอลัมน์ Excel ให้เป็นคีย์มาตรฐาน
    รองรับหัวข้อที่คุณใช้: No., Em. ID, ID Card, Name-TH, Name-EN, Position, Section, Department, Start work, Resign, Status, Degree, Major
    และรองรับหัวไทยบางแบบด้วย
    """
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    cmap = {
        "No.": "no", "No": "no", "ลำดับ": "no",

        "Em. ID": "em_id", "Em ID": "em_id", "Employee ID": "em_id", "รหัสพนักงาน": "em_id",

        "ID Card": "id_card", "IDCard": "id_card", "เลขบัตรประชาชน": "id_card", "บัตรประชาชน": "id_card",

        "Name-TH": "name_th", "Name TH": "name_th", "ชื่อไทย": "name_th", "ชื่อภาษาไทย": "name_th",
        "คำนำหน้า": "title_th",
        "ชื่อ": "first_name_th",
        "นามสกุล": "last_name_th",

        "Name-EN": "name_en", "Name EN": "name_en", "ชื่ออังกฤษ": "name_en",
        "Title (EN)": "title_en",
        "First name (EN)": "first_name_en",
        "Last name (EN)": "last_name_en",

        "Position": "position", "ตำแหน่ง": "position",
        "Section": "section",
        "Department": "department", "แผนก": "department",

        "Start work": "start_work", "Start Work": "start_work", "วันที่เริ่มงาน": "start_work",
        "Resign": "resign", "วันที่ลาออก": "resign",

        "Status": "status", "สถานะ": "status",
        "Degree": "degree", "วุฒิ": "degree",
        "Major": "major", "สาขา": "major",
    }

    df.rename(columns={c: cmap.get(c, c) for c in df.columns}, inplace=True)
    return df


# -------------------------
# Routes
# -------------------------
@app.get("/")
def root():
    return redirect(url_for("employees_list"))


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

                Employee.title_th.ilike(like),
                Employee.first_name_th.ilike(like),
                Employee.last_name_th.ilike(like),

                Employee.title_en.ilike(like),
                Employee.first_name_en.ilike(like),
                Employee.last_name_en.ilike(like),

                Employee.position.ilike(like),
                Employee.section.ilike(like),
                Employee.department.ilike(like),
                Employee.status.ilike(like),
            )
        )

    employees = query.order_by(Employee.no.asc().nullslast(), Employee.em_id.asc()).all()
    return render_template("employees.html", employees=employees, q=q, total=len(employees))


@app.route("/employees/new", methods=["GET", "POST"])
def employee_new():
    if request.method == "POST":
        em_id = s(request.form.get("em_id"))
        if not em_id:
            flash("กรุณากรอก Em. ID", "error")
            return redirect(url_for("employee_new"))

        if Employee.query.filter_by(em_id=em_id).first():
            flash("Em. ID นี้มีอยู่แล้ว (ห้ามซ้ำ)", "error")
            return redirect(url_for("employee_new"))

        emp = Employee(
            no=i(request.form.get("no")),
            em_id=em_id,
            id_card=s(request.form.get("id_card")),

            title_th=s(request.form.get("title_th")),
            first_name_th=s(request.form.get("first_name_th")),
            last_name_th=s(request.form.get("last_name_th")),

            title_en=s(request.form.get("title_en")),
            first_name_en=s(request.form.get("first_name_en")),
            last_name_en=s(request.form.get("last_name_en")),

            position=s(request.form.get("position")),
            section=s(request.form.get("section")),
            department=s(request.form.get("department")),

            start_work=d(request.form.get("start_work")),
            resign=d(request.form.get("resign")),

            status=s(request.form.get("status")),
            degree=s(request.form.get("degree")),
            major=s(request.form.get("major")),
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
        emp.no = i(request.form.get("no"))
        emp.id_card = s(request.form.get("id_card"))

        emp.title_th = s(request.form.get("title_th"))
        emp.first_name_th = s(request.form.get("first_name_th"))
        emp.last_name_th = s(request.form.get("last_name_th"))

        emp.title_en = s(request.form.get("title_en"))
        emp.first_name_en = s(request.form.get("first_name_en"))
        emp.last_name_en = s(request.form.get("last_name_en"))

        emp.position = s(request.form.get("position"))
        emp.section = s(request.form.get("section"))
        emp.department = s(request.form.get("department"))

        emp.start_work = d(request.form.get("start_work"))
        emp.resign = d(request.form.get("resign"))

        emp.status = s(request.form.get("status"))
        emp.degree = s(request.form.get("degree"))
        emp.major = s(request.form.get("major"))

        db.session.commit
