from flask import Flask, request, render_template_string, send_file
import pandas as pd
import numpy as np
import openpyxl
from io import BytesIO
import re

app = Flask(__name__)

HTML_FORM = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel Reconciliation Tool</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            justify-content: center;
            align-items: center;
            padding: 20px;
        }
        .card {
            background: white;
            border-radius: 20px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
            padding: 40px;
            max-width: 600px;
            width: 100%;
            transition: transform 0.3s ease;
        }
        .card:hover {
            transform: translateY(-5px);
        }
        h1 {
            color: #333;
            font-size: 28px;
            margin-bottom: 10px;
            text-align: center;
        }
        .subtitle {
            text-align: center;
            color: #666;
            margin-bottom: 30px;
            font-size: 14px;
        }
        .form-group {
            margin-bottom: 25px;
        }
        label {
            display: block;
            margin-bottom: 8px;
            font-weight: 600;
            color: #444;
        }
        input[type="file"] {
            width: 100%;
            padding: 10px;
            border: 2px dashed #ccc;
            border-radius: 10px;
            background: #f9f9f9;
            cursor: pointer;
            transition: all 0.3s;
        }
        input[type="file"]:hover {
            border-color: #667eea;
            background: #f0f0ff;
        }
        button {
            width: 100%;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border: none;
            padding: 14px;
            border-radius: 10px;
            font-size: 16px;
            font-weight: bold;
            cursor: pointer;
            transition: transform 0.2s, box-shadow 0.2s;
        }
        button:hover {
            transform: scale(1.02);
            box-shadow: 0 10px 20px rgba(0,0,0,0.2);
        }
        .footer {
            margin-top: 20px;
            text-align: center;
            font-size: 12px;
            color: #888;
        }
        .alert {
            padding: 12px;
            border-radius: 8px;
            margin-bottom: 20px;
            text-align: center;
        }
        .alert-error {
            background-color: #ffe6e6;
            color: #cc0000;
            border: 1px solid #ffcccc;
        }
        .alert-success {
            background-color: #e6ffe6;
            color: #2e7d32;
            border: 1px solid #ccffcc;
        }
    </style>
</head>
<body>
    <div class="card">
        <h1>📊 Excel Reconciliation Tool</h1>
        <div class="subtitle">Upload your files and get a formatted Excel report</div>

        {% if error %}
        <div class="alert alert-error">{{ error }}</div>
        {% endif %}
        {% if success %}
        <div class="alert alert-success">{{ success }}</div>
        {% endif %}

        <form method="post" enctype="multipart/form-data" action="/">
            <div class="form-group">
                <label>📄 Main File (.html / .xls / .xlsx)</label>
                <input type="file" name="main_file" accept=".html,.xls,.xlsx" required>
            </div>
            <div class="form-group">
                <label>🔍 Lookup File (.xlsx)</label>
                <input type="file" name="lookup_file" accept=".xlsx" required>
            </div>
            <button type="submit">🚀 Process & Download</button>
        </form>
        <div class="footer">
            The output file will be downloaded automatically after processing.
        </div>
    </div>
</body>
</html>
"""

def process_files(main_file, lookup_file):
    """Core processing logic - exactly as in original Colab script."""
    # 2️⃣ MAIN FILE READ (HTML or Excel)
    try:
        tables = pd.read_html(main_file)
        df = tables[0]
        print("✅ HTML format में पढ़ा गया")
    except Exception:
        df = pd.read_excel(main_file, dtype={"UTR NO": str})
        print("✅ Excel format में पढ़ा गया")

    # 3️⃣ HEADER FIX
    df.columns = df.iloc[0]
    df = df[1:].reset_index(drop=True)

    # 4️⃣ CLEAN COLUMN NAMES
    df.columns = df.columns.str.strip()

    # 5️⃣ NUMERIC FIX
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

    # 6️⃣ DATE FIX
    date_columns = ["FILE_SUBMISSION_DT", "UTR DATE", "INVOICE DATE","RECONCILED DATE"]
    for col in date_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce', dayfirst=True)

    # 7️⃣ FILTER UNIT_NAME
    if "UNIT_NAME" in df.columns:
        df = df[~df["UNIT_NAME"].isin(["Zynova", "---END---"])]

    # 8️⃣ ✅ UTR NO FINAL FIX
    if "UTR NO" in df.columns:
        df["UTR NO"] = df["UTR NO"].fillna("").astype(str)
        df["UTR NO"] = df["UTR NO"].replace("nan", "").str.strip()

    # 9️⃣ VISIT COLUMN
    if "VISIT_ID" in df.columns:
        df["VISIT"] = df["VISIT_ID"].astype(str).str[:2]
        df["VISIT"] = df["VISIT"].replace("ER", "OP")
        df["VISIT_ID"] = df["VISIT_ID"].astype(str).str.replace("^ER", "OP", regex=True)

    # 🔟 LOOKUP FILE
    lookup_df = pd.read_excel(lookup_file)
    lookup_df.columns = lookup_df.columns.str.strip()
    lookup_df = lookup_df.iloc[:, 1:4]
    lookup_df.columns = ["SPONSOR", "Existing", "Payer"]

    payer_map    = lookup_df.drop_duplicates(subset="SPONSOR").set_index("SPONSOR")["Payer"]
    existing_map = lookup_df.drop_duplicates(subset="SPONSOR").set_index("SPONSOR")["Existing"]

    if "SPONSOR" in df.columns:
        df["Payer"]    = df["SPONSOR"].map(payer_map).fillna("NA")
        df["Existing"] = df["SPONSOR"].map(existing_map).fillna("NA")

    # 1️⃣1️⃣ TOTAL COLUMN
    amount_cols = [
        "RECEIVED AMOUNT", "TDS AMOUNT", "WRITEOFF AMOUNT",
        "PATIENT AMOUNT", "PROCESSING FEE", "LEGITIMATE DISCOUNT",
        "DISALLOWANCE AMOUNT"
    ]
    for col in amount_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    df["Total"] = df[[c for c in amount_cols if c in df.columns]].sum(axis=1)

    # 1️⃣2️⃣ FINAL COLUMN ORDER
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

    # 1️⃣3️⃣ TEMP SAVE → OPENPYXL FORMATTING
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        final_df.to_excel(writer, index=False, sheet_name='Sheet1')
        # Get workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        # DATE FORMAT APPLY
        for col_name in date_columns:
            if col_name in final_df.columns:
                col_index = list(final_df.columns).index(col_name) + 1
                for row in range(2, worksheet.max_row + 1):
                    cell = worksheet.cell(row=row, column=col_index)
                    if cell.value:
                        cell.number_format = 'DD-MM-YYYY'

        # NUMBER FORMAT FIX
        for col in range(1, worksheet.max_column + 1):
            for row in range(2, worksheet.max_row + 1):
                cell = worksheet.cell(row=row, column=col)
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '0'

    output.seek(0)
    return output

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Check if files are present
        if 'main_file' not in request.files or 'lookup_file' not in request.files:
            return render_template_string(HTML_FORM, error="Please upload both files.")
        main_file = request.files['main_file']
        lookup_file = request.files['lookup_file']
        if main_file.filename == '' or lookup_file.filename == '':
            return render_template_string(HTML_FORM, error="No file selected.")

        try:
            # Process in memory
            output_excel = process_files(main_file, lookup_file)
            return send_file(
                output_excel,
                as_attachment=True,
                download_name="Final_Output.xlsx",
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            error_msg = str(e)
            # Provide a user-friendly error
            return render_template_string(HTML_FORM, error=f"Processing failed: {error_msg}")

    # GET request: show the form
    return render_template_string(HTML_FORM)

if __name__ == '__main__':
    app.run(debug=True)