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

@app.route("/employees")
def employees():
    return """
    <h2>Employee Report</h2>
    <table border="1" cellpadding="5">
        <tr>
            <th>ชื่อ</th>
            <th>แผนก</th>
            <th>ตำแหน่ง</th>
            <th>ผ่าน OJT</th>
        </tr>
        <tr>
            <td>สมชาย</td>
            <td>Diecasting</td>
            <td>Operator</td>
            <td>✔</td>
        </tr>
    </table>
    <br>
    <a href="/">กลับหน้าหลัก</a>
    """

@app.route("/training")
def training():
    return """
    <h2>Training Matrix</h2>
    <p>Core Tools / ISO / IATF / Safety Training</p>
    <br>
    <a href="/">กลับหน้าหลัก</a>
    """

if __name__ == "__main__":
    app.run()
