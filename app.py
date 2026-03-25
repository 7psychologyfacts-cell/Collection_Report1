import os
import io
import pandas as pd
import numpy as np
import openpyxl
from flask import Flask, request, send_file, render_template_string
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max

HTML = open("templates/index.html").read() if os.path.exists("templates/index.html") else ""

@app.route("/")
def index():
    with open("templates/index.html") as f:
        return f.read()

@app.route("/process", methods=["POST"])
def process():
    file1 = request.files.get("main_file")
    file2 = request.files.get("lookup_file")

    if not file1 or not file2:
        return "Both files are required.", 400

    # =============================
    # 2️⃣ MAIN FILE READ (HTML या Excel)
    # =============================
    file1_bytes = file1.read()
    filename1 = secure_filename(file1.filename)

    try:
        tables = pd.read_html(io.BytesIO(file1_bytes))
        df = tables[0]
    except Exception:
        df = pd.read_excel(io.BytesIO(file1_bytes), dtype={"UTR NO": str})

    # =============================
    # 3️⃣ HEADER FIX
    # =============================
    df.columns = df.iloc[0]
    df = df[1:].reset_index(drop=True)

    # =============================
    # 4️⃣ CLEAN COLUMN NAMES
    # =============================
    df.columns = df.columns.str.strip()

    # =============================
    # 5️⃣ NUMERIC FIX
    # =============================
    numeric_cols = [
        "NET BILL AMT.", "SPONSER_AMOUNT", "CLAIM AMOUNT",
        "RECEIVED AMOUNT", "TDS AMOUNT", "WRITEOFF AMOUNT",
        "PATIENT AMOUNT", "PROCESSING FEE", "LEGITIMATE DISCOUNT",
        "DISALLOWANCE AMOUNT", "Total"
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(",", "")
            df[col] = pd.to_numeric(df[col], errors='coerce')

    # =============================
    # 6️⃣ DATE FIX
    # =============================
    date_columns = ["FILE_SUBMISSION_DT", "UTR DATE", "INVOICE DATE", "RECONCILED DATE"]
    for col in date_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce', dayfirst=True)

    # =============================
    # 7️⃣ FILTER UNIT_NAME
    # =============================
    if "UNIT_NAME" in df.columns:
        df = df[~df["UNIT_NAME"].isin(["Zynova", "---END---"])]

    # =============================
    # 8️⃣ UTR NO FINAL FIX
    # =============================
    if "UTR NO" in df.columns:
        df["UTR NO"] = df["UTR NO"].fillna("").astype(str)
        df["UTR NO"] = df["UTR NO"].replace("nan", "").str.strip()

    # =============================
    # 9️⃣ VISIT COLUMN
    # =============================
    if "VISIT_ID" in df.columns:
        df["VISIT"] = df["VISIT_ID"].astype(str).str[:2]
        df["VISIT"] = df["VISIT"].replace("ER", "OP")
        df["VISIT_ID"] = df["VISIT_ID"].astype(str).str.replace("^ER", "OP", regex=True)

    # =============================
    # 🔟 LOOKUP FILE
    # =============================
    file2_bytes = file2.read()
    lookup_df = pd.read_excel(io.BytesIO(file2_bytes))
    lookup_df.columns = lookup_df.columns.str.strip()
    lookup_df = lookup_df.iloc[:, 1:4]
    lookup_df.columns = ["SPONSOR", "Existing", "Payer"]

    payer_map    = lookup_df.drop_duplicates(subset="SPONSOR").set_index("SPONSOR")["Payer"]
    existing_map = lookup_df.drop_duplicates(subset="SPONSOR").set_index("SPONSOR")["Existing"]

    if "SPONSOR" in df.columns:
        df["Payer"]    = df["SPONSOR"].map(payer_map).fillna("NA")
        df["Existing"] = df["SPONSOR"].map(existing_map).fillna("NA")

    # =============================
    # 1️⃣1️⃣ TOTAL COLUMN
    # =============================
    amount_cols = [
        "RECEIVED AMOUNT", "TDS AMOUNT", "WRITEOFF AMOUNT",
        "PATIENT AMOUNT", "PROCESSING FEE", "LEGITIMATE DISCOUNT",
        "DISALLOWANCE AMOUNT"
    ]
    for col in amount_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    df["Total"] = df[[c for c in amount_cols if c in df.columns]].sum(axis=1)

    # =============================
    # 1️⃣2️⃣ FINAL COLUMN ORDER
    # =============================
    final_cols = [
        "UNIT_NAME", "RECONCILED DATE", "VISIT", "VISIT_ID", "ADMISSION NUMBER",
        "MRNO", "PATIENT NAME", "INVOICE_NO", "INVOICE DATE", "Payer", "Existing",
        "SPONSOR", "UTR NO", "UTR DATE", "NET BILL AMT.", "SPONSER_AMOUNT",
        "CLAIM AMOUNT", "RECEIVED AMOUNT", "TDS AMOUNT", "WRITEOFF AMOUNT",
        "PATIENT AMOUNT", "PROCESSING FEE", "LEGITIMATE DISCOUNT",
        "DISALLOWANCE AMOUNT", "Total", "REMARKS", "FILE_SUBMISSION_DT",
        "IS RESUBMISION", "ADMITTING DR.", "SPECIALITY"
    ]
    final_cols = [c for c in final_cols if c in df.columns]
    final_df = df[final_cols]

    # =============================
    # 1️⃣3️⃣ TEMP SAVE → OPENPYXL FORMATTING
    # =============================
    temp_buffer = io.BytesIO()
    final_df.to_excel(temp_buffer, index=False)
    temp_buffer.seek(0)

    wb = openpyxl.load_workbook(temp_buffer)
    ws = wb.active

    # DATE FORMAT APPLY
    for col_name in date_columns:
        if col_name in final_df.columns:
            col_index = list(final_df.columns).index(col_name) + 1
            for row in range(2, ws.max_row + 1):
                cell = ws.cell(row=row, column=col_index)
                if cell.value:
                    cell.number_format = 'DD-MM-YYYY'

    # NUMBER FORMAT FIX
    for col in range(1, ws.max_column + 1):
        for row in range(2, ws.max_row + 1):
            cell = ws.cell(row=row, column=col)
            if isinstance(cell.value, (int, float)):
                cell.number_format = '0'

    # =============================
    # 1️⃣4️⃣ FINAL SAVE + SEND
    # =============================
    output_buffer = io.BytesIO()
    wb.save(output_buffer)
    output_buffer.seek(0)

    return send_file(
        output_buffer,
        as_attachment=True,
        download_name="Final_Output.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    app.run(debug=True)
