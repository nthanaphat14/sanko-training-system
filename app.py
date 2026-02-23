import os
import re
from datetime import datetime, date
from io import BytesIO

import pandas as pd
from flask import (
    Flask, render_template, request, redirect, url_for,
    flash, send_file
)
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import or_, func

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-key-change-me")

# SQLite (Render ใช้งานได้ แต่อย่าลืมว่า Free instance อาจรีเซ็ตได้)
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DB_PATH = os.path.join(BASE_DIR, "employees.db")
app.config["SQLALCHEMY_DATABASE_URI"] = os.environ.get("DATABASE_URL", f"sqlite:///{DB_PATH}")
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)


# =========================
# Model
# =========================
class Employee(db.Model):
    id = db.Column(db.Integer, primary_key=True)

    no = db.Column(db.Integer, nullable=True)
    em_id = db.Column(db.String(50), unique=True, nullable=False, index=True)
    id_card = db.Column(db.String(50), nullable=True)

    # TH name split
    title_th = db.Column(db.String(50), nullable=True)
    first_name_th = db.Column(db.String(100), nullable=True)
    last_name_th = db.Column(db.String(100), nullable=True)

    # EN name split
    title_en = db.Column(db.String(50), nullable=True)
    first_name_en = db.Column(db.String(100), nullable=True)
    last_name_en = db.Column(db.String(100), nullable=True)

    position = db.Column(db.String(150), nullable=True, index=True)
    section = db.Column(db.String(150), nullable=True, index=True)
    department = db.Column(db.String(150), nullable=True, index=True)

    start_work = db.Column(db.Date, nullable=True)
    resign = db.Column(db.Date, nullable=True)

    status = db.Column(db.String(50), nullable=True, index=True)  # e.g., W, Active, Resign
    degree = db.Column(db.String(100), nullable=True)
    major = db.Column(db.String(150), nullable=True)

    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    @property
    def name_th_full(self) -> str:
        parts = [p for p in [self.title_th, self.first_name_th, self.last_name_th] if p]
        return " ".join(parts).strip()

    @property
    def name_en_full(self) -> str:
        parts = [p for p in [self.title_en, self.first_name_en, self.last_name_en] if p]
        return " ".join(parts).strip()


with app.app_context():
    db.create_all()


# =========================
# Helpers
# =========================
def normalize_col(col: str) -> str:
    """normalize column names (thai/english) for mapping"""
    if col is None:
        return ""
    c = str(col).strip().lower()
    c = re.sub(r"\s+", " ", c)
    c = c.replace("\n", " ").strip()
    return c


def pick_col(df_cols, candidates):
    """
    Return first matching column name from df_cols that matches any candidate (normalized).
    candidates: list[str] normalized patterns (exact) or regex-like keywords
    """
    norm_map = {normalize_col(c): c for c in df_cols}
    for cand in candidates:
        cand_n = normalize_col(cand)
        # exact match
        if cand_n in norm_map:
            return norm_map[cand_n]
    # keyword match (contains)
    for cand in candidates:
        cand_n = normalize_col(cand)
        for nc, original in norm_map.items():
            if cand_n and cand_n in nc:
                return original
    return None


def safe_int(v):
    try:
        if pd.isna(v):
            return None
        s = str(v).strip()
        if s == "":
            return None
        return int(float(s))
    except Exception:
        return None


def safe_str(v):
    if v is None:
        return None
    if isinstance(v, float) and pd.isna(v):
        return None
    s = str(v).strip()
    return s if s != "" else None


def parse_date_any(v):
    """
    Accept:
    - pandas Timestamp
    - yyyy-mm-dd
    - dd-mm-yy / dd-mm-yyyy
    - dd/mm/yyyy
    - excel serial (numeric)
    """
    if v is None:
        return None
    if isinstance(v, (datetime, date)):
        return v.date() if isinstance(v, datetime) else v
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime().date()

    # Excel serial
    if isinstance(v, (int, float)) and not pd.isna(v):
        try:
            # pandas handles Excel serial sometimes; fallback:
            d = pd.to_datetime(v, unit="D", origin="1899-12-30")
            return d.to_pydatetime().date()
        except Exception:
            pass

    s = safe_str(v)
    if not s:
        return None

    # try parse with pandas
    for dayfirst in (True, False):
        try:
            d = pd.to_datetime(s, dayfirst=dayfirst, errors="raise")
            if pd.isna(d):
                continue
            return d.to_pydatetime().date()
        except Exception:
            continue
    return None


TH_TITLES = {"นาย", "นาง", "นางสาว", "ด.ช.", "ด.ญ.", "คุณ", "mr.", "mrs.", "ms."}
EN_TITLES = {"mr", "mrs", "ms", "miss", "dr", "prof"}


def split_th_name(full_name: str):
    """
    Best-effort split Thai name into title/first/last.
    Accept "นาย สมชาย ใจดี" or "สมชาย ใจดี"
    """
    if not full_name:
        return (None, None, None)
    s = str(full_name).strip()
    s = re.sub(r"\s+", " ", s)

    parts = s.split(" ")
    if len(parts) == 1:
        return (None, parts[0], None)

    # detect title
    title = None
    if parts[0] in TH_TITLES:
        title = parts[0]
        parts = parts[1:]

    first = parts[0] if len(parts) >= 1 else None
    last = parts[1] if len(parts) >= 2 else None

    # if more than 2 parts -> last name join
    if len(parts) > 2:
        last = " ".join(parts[1:])

    return (title, first, last)


def split_en_name(full_name: str):
    """
    Best-effort split EN name into title/first/last.
    Accept "Mr. John Smith" or "John Smith"
    """
    if not full_name:
        return (None, None, None)

    s = str(full_name).strip()
    s = re.sub(r"\s+", " ", s)

    parts = s.replace(".", "").split(" ")
    if len(parts) == 1:
        return (None, parts[0], None)

    title = None
    if parts[0].lower() in EN_TITLES:
        title = parts[0]
        parts = parts[1:]

    first = parts[0] if len(parts) >= 1 else None
    last = parts[1] if len(parts) >= 2 else None
    if len(parts) > 2:
        last = " ".join(parts[1:])

    return (title, first, last)


def build_employee_query():
    """
    Apply filters from query string:
    q, department, position, status
    """
    q = (request.args.get("q") or "").strip()
    dept = (request.args.get("department") or "").strip()
    pos = (request.args.get("position") or "").strip()
    st = (request.args.get("status") or "").strip()

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
                Employee.degree.ilike(like),
                Employee.major.ilike(like),
            )
        )

    if dept:
        query = query.filter(Employee.department == dept)

    if pos:
        query = query.filter(Employee.position == pos)

    if st:
        query = query.filter(Employee.status == st)

    # order
    query = query.order_by(Employee.no.asc().nullslast(), Employee.em_id.asc())
    return query


def distinct_values(column):
    vals = db.session.query(column).filter(column.isnot(None)).distinct().order_by(column.asc()).all()
    return [v[0] for v in vals if v[0] is not None and str(v[0]).strip() != ""]


# =========================
# Routes
# =========================
@app.get("/")
def index():
    return redirect(url_for("employees"))


@app.get("/employees")
def employees():
    query = build_employee_query()
    employees_list = query.all()

    total = Employee.query.count()

    # Active/Resign badges: พยายามเดา status
    active_count = Employee.query.filter(
        or_(
            Employee.status.ilike("%active%"),
            Employee.status.ilike("%w%"),
            Employee.status.ilike("%working%"),
            Employee.status.ilike("%work%")
        )
    ).count()

    resign_count = Employee.query.filter(
        or_(
            Employee.status.ilike("%resign%"),
            Employee.status.ilike("%ลาออก%"),
            Employee.status.ilike("%terminate%"),
        )
    ).count()

    departments = distinct_values(Employee.department)
    positions = distinct_values(Employee.position)
    statuses = distinct_values(Employee.status)

    return render_template(
        "employees.html",
        employees=employees_list,
        total=total,
        active_count=active_count,
        resign_count=resign_count,
        departments=departments,
        positions=positions,
        statuses=statuses,
        q=request.args.get("q", ""),
        department=request.args.get("department", ""),
        position=request.args.get("position", ""),
        status=request.args.get("status", "")
    )


@app.route("/employee/new", methods=["GET", "POST"])
def employee_new():
    if request.method == "POST":
        data = request.form

        em_id = (data.get("em_id") or "").strip()
        if not em_id:
            flash("กรุณากรอก Em. ID", "danger")
            return redirect(url_for("employee_new"))

        if Employee.query.filter_by(em_id=em_id).first():
            flash("Em. ID นี้มีอยู่แล้ว (ห้ามซ้ำ)", "danger")
            return redirect(url_for("employee_new"))

        e = Employee(
            no=safe_int(data.get("no")),
            em_id=em_id,
            id_card=safe_str(data.get("id_card")),

            title_th=safe_str(data.get("title_th")),
            first_name_th=safe_str(data.get("first_name_th")),
            last_name_th=safe_str(data.get("last_name_th")),

            title_en=safe_str(data.get("title_en")),
            first_name_en=safe_str(data.get("first_name_en")),
            last_name_en=safe_str(data.get("last_name_en")),

            position=safe_str(data.get("position")),
            section=safe_str(data.get("section")),
            department=safe_str(data.get("department")),

            start_work=parse_date_any(data.get("start_work")),
            resign=parse_date_any(data.get("resign")),

            status=safe_str(data.get("status")),
            degree=safe_str(data.get("degree")),
            major=safe_str(data.get("major")),
        )
        db.session.add(e)
        db.session.commit()
        flash("เพิ่มพนักงานเรียบร้อย", "success")
        return redirect(url_for("employees"))

    return render_template("employee_form.html", mode="new", employee=None)


@app.route("/employee/<int:emp_id>/edit", methods=["GET", "POST"])
def employee_edit(emp_id):
    e = Employee.query.get_or_404(emp_id)

    if request.method == "POST":
        data = request.form

        e.no = safe_int(data.get("no"))
        # em_id ห้ามแก้ (เพื่อกันข้อมูลหลุด)
        e.id_card = safe_str(data.get("id_card"))

        e.title_th = safe_str(data.get("title_th"))
        e.first_name_th = safe_str(data.get("first_name_th"))
        e.last_name_th = safe_str(data.get("last_name_th"))

        e.title_en = safe_str(data.get("title_en"))
        e.first_name_en = safe_str(data.get("first_name_en"))
        e.last_name_en = safe_str(data.get("last_name_en"))

        e.position = safe_str(data.get("position"))
        e.section = safe_str(data.get("section"))
        e.department = safe_str(data.get("department"))

        e.start_work = parse_date_any(data.get("start_work"))
        e.resign = parse_date_any(data.get("resign"))

        e.status = safe_str(data.get("status"))
        e.degree = safe_str(data.get("degree"))
        e.major = safe_str(data.get("major"))

        db.session.commit()
        flash("แก้ไขข้อมูลเรียบร้อย", "success")
        return redirect(url_for("employees"))

    return render_template("employee_form.html", mode="edit", employee=e)


@app.post("/employee/<int:emp_id>/delete")
def employee_delete(emp_id):
    e = Employee.query.get_or_404(emp_id)
    db.session.delete(e)
    db.session.commit()
    flash("ลบพนักงานเรียบร้อย", "success")
    return redirect(url_for("employees"))


@app.route("/import", methods=["GET", "POST"])
def import_employees():
    if request.method == "POST":
        f = request.files.get("file")
        if not f or f.filename.strip() == "":
            flash("กรุณาเลือกไฟล์ Excel/CSV", "danger")
            return redirect(url_for("import_employees"))

        filename = f.filename.lower()
        try:
            if filename.endswith(".csv"):
                df = pd.read_csv(f)
            else:
                # xlsx / xls
                df = pd.read_excel(f)

            if df.empty:
                flash("ไฟล์ว่าง ไม่มีข้อมูล", "danger")
                return redirect(url_for("import_employees"))

            # column mapping (รองรับหลายชื่อ)
            col_no = pick_col(df.columns, ["no.", "no", "ลำดับ", "ลำดับที่"])
            col_emid = pick_col(df.columns, ["em. id", "em id", "employee id", "รหัสพนักงาน", "em_id"])
            col_idcard = pick_col(df.columns, ["id card", "idcard", "เลขบัตรประชาชน", "บัตรประชาชน"])

            # TH names (รองรับทั้งแบบแยกและรวม)
            col_title_th = pick_col(df.columns, ["คำนำหน้า", "คำนำหน้า (th)", "title_th", "title th"])
            col_first_th = pick_col(df.columns, ["ชื่อภาษาไทย", "ชื่อ (th)", "ชื่อ", "firstname_th", "first_name_th", "ชื่อจริง"])
            col_last_th = pick_col(df.columns, ["นามสกุล", "สกุล (th)", "lastname_th", "last_name_th", "นามสกุล (th)"])
            col_full_th = pick_col(df.columns, ["name-th", "name th", "ชื่อ-ไทย", "ชื่อไทย", "ชื่อภาษาไทย(รวม)"])

            # EN names
            col_title_en = pick_col(df.columns, ["title (en)", "title_en", "title en"])
            col_first_en = pick_col(df.columns, ["first name (en)", "first_name_en", "firstname_en"])
            col_last_en = pick_col(df.columns, ["last name (en)", "last_name_en", "lastname_en"])
            col_full_en = pick_col(df.columns, ["name-en", "name en", "ชื่ออังกฤษ", "ชื่อ-อังกฤษ"])

            col_position = pick_col(df.columns, ["position", "ตำแหน่ง"])
            col_section = pick_col(df.columns, ["section", "หน่วยงาน", "ฝ่าย", "แผนกย่อย"])
            col_department = pick_col(df.columns, ["department", "แผนก", "department "])

            col_start = pick_col(df.columns, ["start work", "start_work", "เริ่มงาน", "วันที่เริ่มงาน"])
            col_resign = pick_col(df.columns, ["resign", "resign date", "ลาออก", "วันที่ลาออก"])
            col_status = pick_col(df.columns, ["status", "สถานะ"])
            col_degree = pick_col(df.columns, ["degree", "วุฒิ"])
            col_major = pick_col(df.columns, ["major", "สาขา"])

            if not col_emid:
                flash("ไม่พบคอลัมน์ Em. ID ในไฟล์ (ต้องมี Em. ID/รหัสพนักงาน)", "danger")
                return redirect(url_for("import_employees"))

            inserted = 0
            updated = 0
            skipped = 0

            for _, row in df.iterrows():
                em_id = safe_str(row.get(col_emid))
                if not em_id:
                    skipped += 1
                    continue

                # TH split priority: (title/first/last) -> (full)
                title_th = safe_str(row.get(col_title_th)) if col_title_th else None
                first_th = safe_str(row.get(col_first_th)) if col_first_th else None
                last_th = safe_str(row.get(col_last_th)) if col_last_th else None

                if (not first_th or not last_th) and col_full_th:
                    t, fth, lth = split_th_name(safe_str(row.get(col_full_th)))
                    title_th = title_th or t
                    first_th = first_th or fth
                    last_th = last_th or lth

                # EN split priority: (title/first/last) -> (full)
                title_en = safe_str(row.get(col_title_en)) if col_title_en else None
                first_en = safe_str(row.get(col_first_en)) if col_first_en else None
                last_en = safe_str(row.get(col_last_en)) if col_last_en else None

                if (not first_en or not last_en) and col_full_en:
                    ten, fen, len_ = split_en_name(safe_str(row.get(col_full_en)))
                    title_en = title_en or ten
                    first_en = first_en or fen
                    last_en = last_en or len_

                obj = Employee.query.filter_by(em_id=em_id).first()
                payload = dict(
                    no=safe_int(row.get(col_no)) if col_no else None,
                    id_card=safe_str(row.get(col_idcard)) if col_idcard else None,

                    title_th=title_th,
                    first_name_th=first_th,
                    last_name_th=last_th,

                    title_en=title_en,
                    first_name_en=first_en,
                    last_name_en=last_en,

                    position=safe_str(row.get(col_position)) if col_position else None,
                    section=safe_str(row.get(col_section)) if col_section else None,
                    department=safe_str(row.get(col_department)) if col_department else None,

                    start_work=parse_date_any(row.get(col_start)) if col_start else None,
                    resign=parse_date_any(row.get(col_resign)) if col_resign else None,

                    status=safe_str(row.get(col_status)) if col_status else None,
                    degree=safe_str(row.get(col_degree)) if col_degree else None,
                    major=safe_str(row.get(col_major)) if col_major else None,
                )

                # ถ้าชื่อไทย "เข้าไม่มา" มักเกิดจากคอลัมน์ผิด -> payload จะช่วยเติมจาก col_full_th
                # ถ้าจะบังคับให้ชื่อไทยต้องมี ก็เช็คตรงนี้ได้

                if obj:
                    for k, v in payload.items():
                        # update only if new value not None/empty
                        if v is not None and str(v).strip() != "":
                            setattr(obj, k, v)
                    updated += 1
                else:
                    obj = Employee(em_id=em_id, **payload)
                    db.session.add(obj)
                    inserted += 1

            db.session.commit()
            flash(f"Import สำเร็จ ✅ เพิ่มใหม่ {inserted} | อัปเดต {updated} | ข้าม {skipped}", "success")
            return redirect(url_for("employees"))

        except Exception as ex:
            flash(f"Import ล้มเหลว: {ex}", "danger")
            return redirect(url_for("import_employees"))

    return render_template("import.html")


@app.get("/export")
def export_excel():
    """
    Export ตามตัวกรองเดียวกับหน้า /employees:
    ?q=...&department=...&position=...&status=...
    """
    query = build_employee_query()
    rows = query.all()

    # build dataframe
    data = []
    for e in rows:
        data.append({
            "No.": e.no,
            "Em. ID": e.em_id,
            "ID Card": e.id_card,
            "Title-TH": e.title_th,
            "FirstName-TH": e.first_name_th,
            "LastName-TH": e.last_name_th,
            "Name-TH": e.name_th_full,
            "Title-EN": e.title_en,
            "FirstName-EN": e.first_name_en,
            "LastName-EN": e.last_name_en,
            "Name-EN": e.name_en_full,
            "Position": e.position,
            "Section": e.section,
            "Department": e.department,
            "Start work": e.start_work.isoformat() if e.start_work else "",
            "Resign": e.resign.isoformat() if e.resign else "",
            "Status": e.status,
            "Degree": e.degree,
            "Major": e.major,
        })

    df = pd.DataFrame(data)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Employee Data")

    output.seek(0)

    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    filename = f"employee_data_{stamp}.xlsx"
    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# =========================
# Run (local)
# =========================
if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=True)
