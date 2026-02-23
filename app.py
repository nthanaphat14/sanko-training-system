from flask import Flask

app = Flask(__name__)

@app.get("/")
def home():
    return """
    <h2>SANKO Training System 🚀</h2>
    <p>ระบบกำลังทำงานอยู่</p>
    """

# สำคัญสำหรับ Render
if __name__ == "__main__":
    app.run()
