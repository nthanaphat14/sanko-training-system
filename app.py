import os
import re
from datetime import datetime, date
from typing import Tuple, Optional

from flask import Flask, request, redirect, url_for, flash, render_template_string
from flask_sqlalchemy import SQLAlchemy
import pandas as pd

app = Flask(__name__)
app.secret_key = "sanko-secret"

# ---------------- DATABASE ----------------
DB_PATH = "employees.sqlite"
app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{DB_PATH}"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
db = SQLAlchemy(app)

# ---------------- MODEL ----------------
class Employee(db.Model):
    id = db.Column(db.Integer, primary_key=True)

    no = db.Column(db.Integer)
    em_id = db.Column(db.String(50), unique=True, nullable=False)

    id_card = db.Column(db.String(50))

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
    resign = db.Column(db.Date)

    status = db.Column(db.String(30))
    degree = db.Column(db.String(50))
    major = db.Column(db.String(120))

    def th_full(self):
        return " ".join(filter(None, [self.title_th, self.first_name_th, self.last_name_th]))

    def en_full(self):
        return " ".join(filter(None, [self.title_en, self.first_name_en, self.last_name_en]))

with app.app_context():
    db.create_all()

# ---------------- HELPERS ----------------
def clean(v):
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    return str(v).strip()

def parse_date(v):
    if not v:
        return None
    if isinstance(v, date):
        return v
    try:
        return pd.to_datetime(v).date()
    except:
        return None

# ---------------- HOME ----------------
@app.route("/")
def home():
    return """
    <h1>SANKO Training System 🚀</h1>
    <ul>
        <li><a href='/employees'>Employee Data</a></li>
        <li><a href='/employees/import'>Import Employee Excel</a></li>
    </ul>
    """

# ---------------- LIST ----------------
@app.route("/employees")
def employees():
    employees = Employee.query.order_by(Employee.no.asc()).all()

    html = """
    <h1>Employee Data</h1>
    <a href='/employees/new'>+ เพิ่มพนักงาน</a> |
    <a href='/employees/import'>Import Excel</a>
    <table border=1 cellpadding=6>
        <tr>
            <th>No</th><th>Em.ID</th><th>Name-TH</th><th>Name-EN</th>
            <th>Department</th><th>Action</th>
        </tr>
    """

    for e in employees:
        html += f"""
        <tr>
            <td>{e.no or ''}</td>
            <td>{e.em_id}</td>
            <td>{e.th_full()}</td>
            <td>{e.en_full()}</td>
            <td>{e.department or ''}</td>
            <td>
                <a href='/employees/{e.id}/edit'>Edit</a> |
                <a href='/employees/{e.id}/delete'>Delete</a>
            </td>
        </tr>
        """

    html += "</table>"
    return html

# ---------------- IMPORT ----------------
@app.route("/employees/import", methods=["GET", "POST"])
def import_employees():
    if request.method == "POST":
        file = request.files["file"]
        df = pd.read_excel(file)
        df.columns = [str(c).strip() for c in df.columns]

        for _, row in df.iterrows():
            em_id = clean(row.get("Em. ID"))
            if not em_id:
                continue

            emp = Employee.query.filter_by(em_id=em_id).first()
            if not emp:
                emp = Employee(em_id=em_id)
                db.session.add(emp)

            emp.no = int(row["No."]) if "No." in df.columns and not pd.isna(row["No."]) else None
            emp.id_card = clean(row.get("ID Card"))

            # ===== ชื่อไทย (ไฟล์คุณแยก 3 คอลัมน์ติดกัน) =====
            th_cols = [c for c in df.columns if "ชื่อภาษาไทย" in c or "นามสกุล" in c]
            if len(th_cols) >= 3:
                emp.title_th = clean(row[th_cols[0]])
                emp.first_name_th = clean(row[th_cols[1]])
                emp.last_name_th = clean(row[th_cols[2]])
            else:
                full = clean(row.get("Name-TH"))
                parts = full.split()
                if parts:
                    if parts[0] in ["นาย","นาง","นางสาว"]:
                        emp.title_th = parts[0]
                        emp.first_name_th = parts[1] if len(parts)>1 else ""
                        emp.last_name_th = parts[2] if len(parts)>2 else ""
                    else:
                        emp.first_name_th = parts[0]
                        emp.last_name_th = parts[1] if len(parts)>1 else ""

            # ===== ชื่ออังกฤษ =====
            en = clean(row.get("Name-EN"))
            parts = en.split()
            if parts:
                if parts[0].lower() in ["mr.","mrs.","ms.","miss"]:
                    emp.title_en = parts[0]
                    emp.first_name_en = parts[1] if len(parts)>1 else ""
                    emp.last_name_en = parts[2] if len(parts)>2 else ""
                else:
                    emp.first_name_en = parts[0]
                    emp.last_name_en = parts[1] if len(parts)>1 else ""

            emp.position = clean(row.get("Position"))
            emp.section = clean(row.get("Section"))
            emp.department = clean(row.get("Department"))
            emp.start_work = parse_date(row.get("Start work"))
            emp.resign = parse_date(row.get("Resign"))
            emp.status = clean(row.get("Status"))
            emp.degree = clean(row.get("Degree"))
            emp.major = clean(row.get("Major"))

        db.session.commit()
        return redirect(url_for("employees"))

    return """
    <h1>Import Employee Data</h1>
    <form method='post' enctype='multipart/form-data'>
        <input type='file' name='file'>
        <button type='submit'>Upload</button>
    </form>
    """

# ---------------- EDIT ----------------
@app.route("/employees/<int:id>/edit", methods=["GET","POST"])
def edit_employee(id):
    emp = Employee.query.get_or_404(id)

    if request.method == "POST":
        emp.no = int(request.form["no"])
        emp.id_card = request.form["id_card"]
        emp.title_th = request.form["title_th"]
        emp.first_name_th = request.form["first_name_th"]
        emp.last_name_th = request.form["last_name_th"]
        emp.department = request.form["department"]
        db.session.commit()
        return redirect(url_for("employees"))

    return f"""
    <h1>แก้ไขพนักงาน: {emp.em_id}</h1>
    <form method='post'>
        No:<input name='no' value='{emp.no or ""}'><br>
        ID Card:<input name='id_card' value='{emp.id_card or ""}'><br>
        คำนำหน้า:<input name='title_th' value='{emp.title_th or ""}'><br>
        ชื่อ:<input name='first_name_th' value='{emp.first_name_th or ""}'><br>
        สกุล:<input name='last_name_th' value='{emp.last_name_th or ""}'><br>
        Department:<input name='department' value='{emp.department or ""}'><br>
        <button type='submit'>Save</button>
    </form>
    """

@app.route("/employees/<int:id>/delete")
def delete_employee(id):
    emp = Employee.query.get_or_404(id)
    db.session.delete(emp)
    db.session.commit()
    return redirect(url_for("employees"))

if __name__ == "__main__":
    app.run(debug=True)
