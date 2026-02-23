from flask import Flask, request
import sqlite3

app = Flask(__name__)

DB_NAME = "database.db"


def get_conn():
    return sqlite3.connect(DB_NAME)


def init_db():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS employees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            dept TEXT NOT NULL,
            position TEXT NOT NULL
        )
        """
    )
    conn.commit()
    conn.close()


init_db()


@app.get("/")
def home():
    return """
    <h1>SANKO Training System 🚀</h1>
    <h3>เมนูหลัก</h3>
    <ul>
        <li><a href="/employees">Employee Report</a></li>
        <li><a href="/training">Training Matrix</a></li>
    </ul>
    """


@app.route("/employees", methods=["GET", "POST"])
def employees():
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        dept = request.form.get("dept", "").strip()
        position = request.form.get("position", "").strip()

        if name and dept and position:
            conn = get_conn()
            cur = conn.cursor()
            cur.execute(
                "INSERT INTO employees (name, dept, position) VALUES (?, ?, ?)",
                (name, dept, position),
            )
            conn.commit()
            conn.close()

    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT id, name, dept, position FROM employees ORDER BY id DESC")
    rows = cur.fetchall()
    conn.close()

    table_rows = ""
    for r in rows:
        table_rows += f"""
        <tr>
            <td>{r[0]}</td>
            <td>{r[1]}</td>
            <td>{r[2]}</td>
            <td>{r[3]}</td>
        </tr>
        """

    return f"""
    <h2>Employee Report</h2>

    <form method="POST">
        ชื่อ: <input name="name" required>
        แผนก: <input name="dept" required>
        ตำแหน่ง: <input name="position" required>
        <button type="submit">เพิ่มพนักงาน</button>
    </form>

    <br><br>

    <table border="1" cellpadding="5">
        <tr>
            <th>ID</th>
            <th>ชื่อ</th>
            <th>แผนก</th>
            <th>ตำแหน่ง</th>
        </tr>
        {table_rows}
    </table>

    <br>
    <a href="/">กลับหน้าหลัก</a>
    """


@app.get("/training")
def training():
    return """
    <h2>Training Matrix</h2>
    <p>หน้านี้จะทำต่อ: ใส่ Matrix ตามตำแหน่ง + หลักสูตร</p>
    <br>
    <a href="/">กลับหน้าหลัก</a>
    """
