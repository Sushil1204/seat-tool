from flask import Flask, render_template, request, send_file, redirect, url_for
import pandas as pd
import re, os
from PyPDF2 import PdfReader

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

LATEST_FILE = None   # will store latest excel path

@app.route("/", methods=["GET", "POST"])
def index():
    global LATEST_FILE
    table_data = None

    if request.method == "POST":
        pdf = request.files["pdf"]
        path = os.path.join(UPLOAD_FOLDER, pdf.filename)
        pdf.save(path)

        reader = PdfReader(path)
        text = "\n".join(p.extract_text() for p in reader.pages if p.extract_text())

        pattern = re.compile(
            r"(\d{2}/\d{2}/\d{4})\s+\d+\s+(\d+)\s+\d+\s+([A-Z]+)(.*?)(?=MEDIUM\s+TOTAL)",
            re.S
        )

        rows = []

        for m in pattern.finditer(text):
            date, paper, subject, block = m.groups()
            seats = re.findall(r"M\d{6}", block)

            for i in range(0, len(seats), 30):
                batch_no = (i // 30) + 1
                batch = seats[i:i+30]

                for seat in batch:
                    rows.append({
                        "Room": f"Room {batch_no}",
                        "Batch": batch_no,
                        "Seat Number": seat,
                        "Date": date,
                        "Subject": subject,
                        "Paper": paper,
                        "Medium": "1"
                    })

        df = pd.DataFrame(rows)

        # ----- SAVE EXCEL IN OUTPUTS FOLDER -----
        output_file = "Room_Wise_Seat_List.xlsx"
        output_path = os.path.join(OUTPUT_FOLDER, output_file)
        df.to_excel(output_path, index=False)

        LATEST_FILE = output_path

        table_data = df.to_dict(orient="records")

    return render_template("index.html", data=table_data)


# ----- DOWNLOAD ROUTE -----
@app.route("/download")
def download():
    global LATEST_FILE
    if LATEST_FILE and os.path.exists(LATEST_FILE):
        return send_file(LATEST_FILE, as_attachment=True)
    return "No file generated yet", 404


if __name__ == "__main__":
    app.run(debug=True)
