from flask import Flask, request
import sqlite3

app = Flask(__name__)

# เก็บข้อมูลแบบชั่วคราว (ถ้ารีสตาร์ท/รีโหลด อาจหายได้)
app = Flask(__name__)
def init_db():
    conn = sqlite3.connect("database.db")
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS employees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT,
            dept TEXT,
            position TEXT
        )
    """)
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
    @app.route("/employees", methods=["GET", "POST"])
def employees():

    if request.method == "POST":
        name = request.form["name"]
        dept = request.form["dept"]
        position = request.form["position"]

        conn = sqlite3.connect("database.db")
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO employees (name, dept, position) VALUES (?, ?, ?)",
            (name, dept, position)
        )
        conn.commit()
        conn.close()

    conn = sqlite3.connect("database.db")
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM employees")
    rows = cursor.fetchall()
    conn.close()

    table_rows = ""
    for row in rows:
        table_rows += f"""
        <tr>
            <td>{row[0]}</td>
            <td>{row[1]}</td>
            <td>{row[2]}</td>
            <td>{row[3]}</td>
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
