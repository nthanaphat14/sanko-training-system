from flask import Flask, request

app = Flask(__name__)

# เก็บข้อมูลแบบชั่วคราว (ถ้ารีสตาร์ท/รีโหลด อาจหายได้)
employees_data = []

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
    global employees_data

    if request.method == "POST":
        name = request.form.get("name", "").strip()
        dept = request.form.get("dept", "").strip()
        position = request.form.get("position", "").strip()

        if name and dept and position:
            employees_data.append((name, dept, position))

    rows = ""
    for i, (name, dept, position) in enumerate(employees_data, start=1):
        rows += f"""
        <tr>
          <td>{i}</td>
          <td>{name}</td>
          <td>{dept}</td>
          <td>{position}</td>
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

    <br>

    <table border="1" cellpadding="6">
      <tr>
        <th>#</th>
        <th>ชื่อ</th>
        <th>แผนก</th>
        <th>ตำแหน่ง</th>
      </tr>
      {rows}
    </table>

    <br>
    <a href="/">กลับหน้าหลัก</a>
    """

@app.get("/training")
def training():
    return """
    <h2>Training Matrix</h2>
    <p>หน้านี้เราจะทำต่อเป็นตารางหลักสูตรตามตำแหน่ง</p>
    <a href="/">กลับหน้าหลัก</a>
    """
