from flask import Flask

app = Flask(__name__)

@app.route("/")
def home():
    return """
    <h1>SANKO Training System 🚀</h1>
    <hr>
    <h2>เมนูหลัก</h2>
    <ul>
        <li><a href="/employees">Employee Report</a></li>
        <li><a href="/training">Training Matrix</a></li>
    </ul>
    """

@app.route("/employees", methods=["GET", "POST"])
def employees():
    if request.method == "POST":
        name = request.form["name"]
        dept = request.form["dept"]
        position = request.form["position"]
        employees_data.append((name, dept, position))

    table_rows = ""
    for emp in employees_data:
        table_rows += f"<tr><td>{emp[0]}</td><td>{emp[1]}</td><td>{emp[2]}</td></tr>"

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
            <th>ชื่อ</th>
            <th>แผนก</th>
            <th>ตำแหน่ง</th>
        </tr>
        {table_rows}
    </table>

    <br>
    <a href="/">กลับหน้าหลัก</a>
    """
