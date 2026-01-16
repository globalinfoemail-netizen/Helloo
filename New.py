from flask import Flask, render_template, jsonify
import sqlite3
from datetime import datetime

app = Flask(__name__)
DB = "gpse.db"

def db():
    c = sqlite3.connect(DB)
    c.row_factory = sqlite3.Row
    return c

def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M")

@app.route("/")
@app.route("/dashboard")
def dashboard():
    with db() as c:
        reports = c.execute("SELECT * FROM reports").fetchall()
    return render_template("dashboard.html", reports=reports)

@app.route("/kpi-library")
def kpi_library():
    return render_template("kpi_library.html")

@app.route("/api/departments")
def departments():
    with db() as c:
        rows = c.execute("SELECT * FROM departments").fetchall()
    return jsonify([dict(r) for r in rows])

@app.route("/api/kpis/<int:dept_id>")
def kpis_by_dept(dept_id):
    with db() as c:
        rows = c.execute(
            "SELECT * FROM kpi_master WHERE dept_id=?", (dept_id,)
        ).fetchall()
    return jsonify([dict(r) for r in rows])

@app.route("/api/kpi/<int:kpi_id>")
def kpi_detail(kpi_id):
    with db() as c:
        row = c.execute(
            "SELECT * FROM kpi_master WHERE kpi_id=?", (kpi_id,)
        ).fetchone()
    return jsonify(dict(row))
    
if __name__ == "__main__":
    app.run(debug=True)
