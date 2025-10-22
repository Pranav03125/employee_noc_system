from flask import Flask, render_template, request, redirect, url_for, send_from_directory, flash
import os
from noc_generator import EmployeeNOCGenerator

app = Flask(__name__)
app.secret_key = "super_secret_key"

# Config
TEMPLATE_PATH = "NDA-1.docx"            # Ensure this file exists in project root (or set absolute path)
OUTPUT_FOLDER = "generated_noc"
PORT = 5001

# Ensure folders exist
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Create a single generator instance (or create per-request)
try:
    generator = EmployeeNOCGenerator(TEMPLATE_PATH)
except Exception as e:
    # If you run app before placing the template, this informs you
    print(f"Error initializing NOC generator: {e}")
    raise

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/generate", methods=["POST"])
def generate():
    full_name = request.form.get("full_name", "").strip()
    job_title = request.form.get("job_title", "").strip()
    department = request.form.get("department", "").strip()

    if not full_name or not job_title or not department:
        flash("Please fill all fields: Full Name, Job Title, Department.")
        return redirect(url_for("index"))

    try:
        out_path = generator.generate_noc(full_name, job_title, department, output_dir=OUTPUT_FOLDER)
        flash(f"NOC generated for {full_name}")
        return redirect(url_for("results"))
    except Exception as e:
        flash(f"Error generating NOC: {e}")
        return redirect(url_for("index"))

@app.route("/results")
def results():
    files = [f for f in os.listdir(OUTPUT_FOLDER) if f.lower().endswith(".docx")]
    # sort files by modification time descending
    files = sorted(files, key=lambda f: os.path.getmtime(os.path.join(OUTPUT_FOLDER, f)), reverse=True)
    return render_template("results.html", files=files)

@app.route("/download/<filename>")
def download(filename):
    return send_from_directory(OUTPUT_FOLDER, filename, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=PORT, debug=True)
