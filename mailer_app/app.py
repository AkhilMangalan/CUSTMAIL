from flask import Flask, render_template, request
import os
from customer_GOX import process_and_send

app = Flask(__name__)
UPLOAD = "uploads"
os.makedirs(UPLOAD, exist_ok=True)

@app.route("/", methods=["GET","POST"])
def index():
    if request.method == "POST":
        excel = request.files["excel"]
        template = request.files["template"]

        excel_path = os.path.join(UPLOAD, excel.filename)
        template_path = os.path.join(UPLOAD, template.filename)

        excel.save(excel_path)
        template.save(template_path)

        logs = process_and_send(
            excel_path,
            template_path,
            request.form["api_key"],
            request.form["sender_email"],
            request.form["sender_name"],
            request.form["subject"],
            request.form["body"],
            request.form["cc"],
            request.form["bcc"]
        )

        return render_template("index.html", logs=logs)

    return render_template("index.html", logs=None)

app.run(debug=True)
