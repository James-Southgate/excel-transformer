import os
import io
import tempfile
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from flask import Flask, request, render_template_string, send_file

app = Flask(__name__)

# ------------------------------------------------------------
# Your EXACT transformation functions (copied from your script)
# ------------------------------------------------------------
def normalize(val):
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return ""
    return str(val).strip()

def load_with_computed_offer(file_path, sheet_name=None):
    # Step 1 — Load Excel with formula‑aware Offer column
    read_kwargs = {"dtype": str, "keep_default_na": False}
    if sheet_name:
        read_kwargs["sheet_name"] = sheet_name
    df = pd.read_excel(file_path, **read_kwargs)

    wb = load_workbook(file_path, data_only=False)
    ws = wb.active if sheet_name is None else wb[sheet_name]

    headers = [cell.value for cell in ws[1]]
    try:
        offer_col_idx = headers.index('Offer') + 1
    except ValueError:
        return df

    offer_values = []
    for i in range(len(df)):
        excel_row = i + 2
        cell = ws.cell(row=excel_row, column=offer_col_idx)
        if cell.data_type == 'f':
            sale_str = df.loc[i, 'Sale Price']
            reg_str = df.loc[i, 'Reg Price']
            try:
                sale = float(''.join(c for c in str(sale_str) if c.isdigit() or c == '.'))
                reg = float(''.join(c for c in str(reg_str) if c.isdigit() or c == '.'))
                diff = reg - sale
                if diff == int(diff):
                    offer_text = f"Save ${int(diff)}"
                else:
                    offer_text = f"Save ${diff:.1f}".rstrip('0').rstrip('.')
            except (ValueError, TypeError):
                offer_text = None
        else:
            offer_text = cell.value if cell.value is not None else ''
        offer_values.append(normalize(offer_text))

    df['Offer'] = offer_values

    for col in ['Sale Price', 'Reg Price']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(r'[^0-9.]', '', regex=True)
            df[col] = df[col].str.strip()
    return df

def sort_by_department(df):
    return df.sort_values(by="Department", key=lambda s: s.str.lower(), kind="mergesort").reset_index(drop=True)

def insert_department_header_rows(df):
    rows = []
    previous_department = None
    for _, row in df.iterrows():
        current_department = normalize(row["Department"])
        if current_department != previous_department:
            separator = {col: "" for col in df.columns}
            separator["Department Headers"] = f"{current_department}.eps" if current_department else ".eps"
            rows.append(separator)
            previous_department = current_department
        data = row.to_dict()
        data["Department Headers"] = ""
        rows.append(data)
    return pd.DataFrame(rows, columns=df.columns)

def map_sale_type_eps(df):
    df["Sale Type"] = df["Sale Type"].apply(lambda v: f"{normalize(v)}.eps" if normalize(v) else "")
    return df

def process_offer(df):
    offer_idx = df.columns.get_loc("Offer")
    new_dollar = []
    new_offer_dollar = []
    new_offer_cents = []
    new_offer = []
    for val in df["Offer"]:
        original = normalize(val)
        if "$" in original:
            prefix, price = original.split("$", 1)
            price = price.strip()
            if price.startswith("0."):
                digits_after_decimal = price.split(".", 1)[1]
                new_offer.append(prefix)
                new_dollar.append("")
                new_offer_dollar.append(digits_after_decimal)
                new_offer_cents.append("¢")
            else:
                new_offer.append(prefix)
                new_dollar.append("$")
                if "." in price:
                    dollars, cents = price.split(".", 1)
                    new_offer_dollar.append(dollars)
                    new_offer_cents.append(cents)
                else:
                    new_offer_dollar.append(price)
                    new_offer_cents.append("")
        else:
            new_offer.append(original)
            new_dollar.append("")
            new_offer_dollar.append("")
            new_offer_cents.append("")
    df["Offer"] = new_offer
    df.insert(offer_idx + 1, "$", new_dollar)
    df.insert(offer_idx + 2, "Offer Dollar", new_offer_dollar)
    df.insert(offer_idx + 3, "Offer Cents", new_offer_cents)
    return df

def process_sale_price(df):
    sp_idx = df.columns.get_loc("Sale Price")
    new_sale_price = []
    new_sale_cents = []
    for val in df["Sale Price"]:
        s = normalize(val)
        if s == "":
            new_sale_price.append("")
            new_sale_cents.append("")
            continue
        if "." in s:
            dollars, cents = s.split(".", 1)
        else:
            dollars, cents = s, "00"
        if len(cents) == 1:
            cents = f"0{cents}"
        if dollars == "0":
            new_sale_price.append(cents)
            new_sale_cents.append("¢")
        else:
            new_sale_price.append(dollars)
            new_sale_cents.append(cents)
    df["Sale Price"] = new_sale_price
    df.insert(sp_idx + 1, "Sale Cents", new_sale_cents)
    return df

def process_reg_price(df):
    df["Reg Price"] = df["Reg Price"].apply(
        lambda v: f"Reg Price ${normalize(v)}.00 |" if "." not in normalize(v) and normalize(v) else
                  (f"Reg Price ${normalize(v)} |" if normalize(v) else "")
    )
    return df

def transform_in_memory(input_bytes):
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_in:
        tmp_in.write(input_bytes)
        tmp_in_path = tmp_in.name
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_out:
        tmp_out_path = tmp_out.name
    try:
        df = load_with_computed_offer(tmp_in_path)
        df = sort_by_department(df)
        df.insert(0, "Department Headers", "")
        df = insert_department_header_rows(df)
        df = map_sale_type_eps(df)
        df = process_offer(df)
        df = process_sale_price(df)
        df = process_reg_price(df)
        df.to_excel(tmp_out_path, index=False, engine="openpyxl")
        with open(tmp_out_path, "rb") as f:
            return f.read()
    finally:
        os.unlink(tmp_in_path)
        os.unlink(tmp_out_path)

# ------------------------------------------------------------
# Flask routes
# ------------------------------------------------------------
HTML_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <title>Excel Transformer</title>
    <meta charset="UTF-8">
    <style>
        body { font-family: sans-serif; margin: 2rem; }
        input, button { font-size: 1rem; margin: 0.5rem 0; }
        #status { margin-top: 1rem; font-style: italic; }
    </style>
</head>
<body>
    <h1>Excel Cleanup Tool</h1>
    <p>Upload an Excel file (.xlsx) – get a transformed version back.</p>
    <input type="file" id="fileInput" accept=".xlsx" />
    <button onclick="transform()">Transform & Download</button>
    <div id="status"></div>

    <script>
        async function transform() {
            const fileInput = document.getElementById('fileInput');
            const file = fileInput.files[0];
            if (!file) {
                alert("Select an Excel file first");
                return;
            }
            const statusDiv = document.getElementById('status');
            statusDiv.innerText = "Processing... (cold start may take 30-60s)";
            const formData = new FormData();
            formData.append('file', file);
            try {
                const response = await fetch('/transform', { method: 'POST', body: formData });
                if (!response.ok) {
                    const errorText = await response.text();
                    throw new Error(errorText);
                }
                const blob = await response.blob();
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'transformed.xlsx';
                document.body.appendChild(a);
                a.click();
                a.remove();
                URL.revokeObjectURL(url);
                statusDiv.innerText = "Done! File downloaded.";
            } catch (err) {
                statusDiv.innerText = "Error: " + err.message;
                console.error(err);
            }
        }
    </script>
</body>
</html>
"""

@app.route("/")
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route("/transform", methods=["POST"])
def transform():
    if "file" not in request.files:
        return "No file uploaded", 400
    file = request.files["file"]
    if file.filename == "":
        return "Empty filename", 400

    input_bytes = file.read()
    try:
        output_bytes = transform_in_memory(input_bytes)
        return send_file(
            io.BytesIO(output_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="transformed.xlsx"
        )
    except Exception as e:
        app.logger.exception("Transformation failed")
        return f"Error: {str(e)}", 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)