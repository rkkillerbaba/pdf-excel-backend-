import io
import re
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import StreamingResponse, HTMLResponse
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation

app = FastAPI(title="Exact PDF Data to Master Sheet Converter")

# -------------------------------------------------------------
# 🎨 WEB UI INTERFACE
# -------------------------------------------------------------
HTML_PAGE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PDF to Dynamic Master Sheet Converter</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #eef2f5; color: #333; display: flex; justify-content: center; align-items: center; min-height: 100vh; padding: 20px; }
        .card { background: #ffffff; width: 100%; max-width: 500px; padding: 30px; border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,0.1); text-align: center; }
        h2 { color: #1f4e79; margin-bottom: 8px; font-size: 22px; }
        p { color: #666; font-size: 13px; margin-bottom: 24px; }
        .upload-area { border: 2px dashed #007bff; border-radius: 8px; padding: 30px 20px; background: #f8faff; cursor: pointer; transition: 0.3s; margin-bottom: 20px; }
        .upload-area:hover { background: #eaf2ff; }
        .file-input { display: none; }
        .icon { font-size: 40px; color: #007bff; margin-bottom: 10px; }
        .btn { width: 100%; padding: 12px; background: #28a745; color: white; border: none; border-radius: 6px; font-size: 15px; font-weight: bold; cursor: pointer; transition: 0.3s; }
        .btn:hover { background: #218838; }
        .btn:disabled { background: #ccc; cursor: not-allowed; }
        #file-name { margin-top: 10px; font-size: 12px; font-weight: bold; color: #007bff; }
        .spinner { display: none; margin: 15px auto 0; border: 4px solid #f3f3f3; border-top: 4px solid #007bff; border-radius: 50%; width: 30px; height: 30px; animation: spin 1s linear infinite; }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
    </style>
</head>
<body>
    <div class="card">
        <h2>📄 Bulk PDF Data Extractor</h2>
        <p>PDF Upload Karein & Master Excel Template Download Karein</p>
        <form id="uploadForm">
            <div class="upload-area" onclick="document.getElementById('pdfFile').click()">
                <div class="icon">📁</div>
                <div style="font-weight: bold; font-size: 14px;">Select PDF File</div>
                <div id="file-name">No file selected</div>
            </div>
            <input type="file" id="pdfFile" class="file-input" accept=".pdf" onchange="showFileName()" required>
            <button type="submit" id="submitBtn" class="btn">⚡ Convert PDF to Excel</button>
        </form>
        <div id="spinner" class="spinner"></div>
    </div>

    <script>
        function showFileName() {
            const input = document.getElementById('pdfFile');
            const fileNameDiv = document.getElementById('file-name');
            if (input.files.length > 0) {
                fileNameDiv.innerText = "Selected: " + input.files[0].name;
            }
        }

        document.getElementById('uploadForm').addEventListener('submit', async function(e) {
            e.preventDefault();
            const fileInput = document.getElementById('pdfFile');
            if (fileInput.files.length === 0) return;

            const submitBtn = document.getElementById('submitBtn');
            const spinner = document.getElementById('spinner');

            submitBtn.disabled = true;
            submitBtn.innerText = "Extracting Exact Data...";
            spinner.style.display = "block";

            const formData = new FormData();
            formData.append("file", fileInput.files[0]);

            try {
                const response = await fetch("/convert-pdf/", { method: "POST", body: formData });
                if (!response.ok) throw new Error("Conversion failed!");

                const blob = await response.blob();
                const downloadUrl = window.URL.createObjectURL(blob);
                const a = document.createElement("a");
                a.href = downloadUrl;
                a.download = fileInput.files[0].name.replace(".pdf", "_Extracted.xlsx");
                document.body.appendChild(a);
                a.click();
                a.remove();
            } catch (err) {
                alert("Error extracting PDF data.");
            } finally {
                submitBtn.disabled = false;
                submitBtn.innerText = "⚡ Convert PDF to Excel";
                spinner.style.display = "none";
            }
        });
    </script>
</body>
</html>
"""

@app.get("/", response_class=HTMLResponse)
def serve_ui():
    return HTML_PAGE

# -------------------------------------------------------------
# ⚙️ EXTRACTION & HELPER FUNCTIONS
# -------------------------------------------------------------
def extract_clean_number(val_str: str) -> float:
    """Extracts numeric float values while stripping formulas/symbols (e.g. '₹1,500 (3 × ₹500)' -> 1500.0)"""
    if not val_str:
        return 0.0
    main_part = val_str.split('(')[0]
    clean_str = re.sub(r"[^\d.]", "", main_part)
    try:
        return float(clean_str)
    except:
        return 0.0

def parse_exact_pdf_page(page, page_num: int) -> dict:
    row_data = {
        "Page_No": page_num,
        "Customer_Name": "",
        "Product": "",
        "Quantity": 0,
        "Unit_Price": 0.0,
        "Base_Amount": 0.0,
        "Discount": "",
        "Taxable_Amount": 0.0,
        "GST": 0.0,
        "Net_Total": 0.0,
        "Description": ""
    }

    tables = page.extract_tables()
    
    if tables:
        for table in tables:
            for row in table:
                if len(row) >= 2 and row[0] and row[1]:
                    field = str(row[0]).replace("\n", " ").strip().lower()
                    val = str(row[1]).replace("\n", " ").strip()
                    
                    if "customer" in field:
                        row_data["Customer_Name"] = val
                    elif "product" in field:
                        row_data["Product"] = val
                    elif "quantity" in field:
                        try: row_data["Quantity"] = int(re.sub(r"[^\d]", "", val))
                        except: pass
                    elif "unit price" in field:
                        row_data["Unit_Price"] = extract_clean_number(val)
                    elif "base amount" in field:
                        row_data["Base_Amount"] = extract_clean_number(val)
                    elif "discount" in field:
                        row_data["Discount"] = val
                    elif "taxable" in field:
                        row_data["Taxable_Amount"] = extract_clean_number(val)
                    elif "gst" in field:
                        row_data["GST"] = extract_clean_number(val)
                    elif "money received" in field or "net total" in field:
                        row_data["Net_Total"] = extract_clean_number(val)
                    elif "description" in field:
                        row_data["Description"] = val
    else:
        text = page.extract_text() or ""
        lines = text.split("\n")
        for line in lines:
            if "|" in line or ":" in line:
                parts = re.split(r"[:|]", line, 1)
                if len(parts) == 2:
                    field = parts[0].strip().lower()
                    val = parts[1].strip()
                    
                    if "customer" in field: row_data["Customer_Name"] = val
                    elif "product" in field: row_data["Product"] = val
                    elif "quantity" in field:
                        try: row_data["Quantity"] = int(re.sub(r"[^\d]", "", val))
                        except: pass
                    elif "unit price" in field: row_data["Unit_Price"] = extract_clean_number(val)
                    elif "description" in field: row_data["Description"] = val

    return row_data

# -------------------------------------------------------------
# 🚀 FASTAPI CONVERT ROUTE
# -------------------------------------------------------------
@app.post("/convert-pdf/")
async def convert_pdf(file: UploadFile = File(...)):
    if not file.filename.endswith('.pdf'):
        raise HTTPException(status_code=400, detail="Only PDF files allowed.")
    
    try:
        contents = await file.read()
        pdf_file = io.BytesIO(contents)
        pages_data = []
        
        with pdfplumber.open(pdf_file) as pdf:
            for i, page in enumerate(pdf.pages, start=1):
                pages_data.append(parse_exact_pdf_page(page, i))
        
        wb = openpyxl.Workbook()
        
        # 1. PageViewer Sheet (Interactive Dashboard)
        ws_view = wb.active
        ws_view.title = "PageViewer"
        ws_view.views.sheetView[0].showGridLines = True

        header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        white_title = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
        label_font = Font(name="Calibri", size=11, bold=True, color="333333")
        formula_font = Font(name="Calibri", size=11, bold=False, color="000000")
        accent_font = Font(name="Calibri", size=12, bold=True, color="28A745")

        ws_view.merge_cells("B1:D1")
        ws_view["B1"] = "ACTIVE PAGE DATA DASHBOARD"
        ws_view["B1"].font = white_title
        ws_view["B1"].fill = header_fill
        ws_view["B1"].alignment = Alignment(horizontal="center", vertical="center")

        ws_view["B3"] = "Select Page No:"
        ws_view["B3"].font = label_font
        ws_view["C3"] = 1
        ws_view["C3"].font = Font(size=12, bold=True, color="007BFF")

        # Dynamic Dropdown Binding (Fix for Cell C3)
        total_rows = len(pages_data) + 1
        dv = DataValidation(
            type="list", 
            formula1=f"=DataSheet!$A$2:$A${total_rows}", 
            allow_blank=False
        )
        ws_view.add_data_validation(dv)
        dv.add(ws_view["C3"])

        calc_rows = [
            ("Customer Name", '=XLOOKUP(C3, DataSheet!A:A, DataSheet!B:B, "N/A")'),
            ("Product Name", '=XLOOKUP(C3, DataSheet!A:A, DataSheet!C:C, "N/A")'),
            ("Quantity", '=XLOOKUP(C3, DataSheet!A:A, DataSheet!D:D, 0)'),
            ("Unit Price (₹)", '=XLOOKUP(C3, DataSheet!A:A, DataSheet!E:E, 0)'),
            ("Base Amount (₹)", '=XLOOKUP(C3, DataSheet!A:A, DataSheet!F:F, 0)'),
            ("Discount", '=XLOOKUP(C3, DataSheet!A:A, DataSheet!G:G, "0")'),
            ("Taxable Amount (₹)", '=XLOOKUP(C3, DataSheet!A:A, DataSheet!H:H, 0)'),
            ("GST Amount (₹)", '=XLOOKUP(C3, DataSheet!A:A, DataSheet!I:I, 0)'),
            ("FINAL PAYABLE (₹)", '=XLOOKUP(C3, DataSheet!A:A, DataSheet!J:J, 0)'),
            ("Description", '=XLOOKUP(C3, DataSheet!A:A, DataSheet!K:K, "")')
        ]

        row_idx = 5
        for label, formula in calc_rows:
            ws_view[f"B{row_idx}"] = label
            ws_view[f"B{row_idx}"].font = label_font
            ws_view[f"C{row_idx}"] = formula
            ws_view[f"C{row_idx}"].font = accent_font if "FINAL PAYABLE" in label else formula_font
            row_idx += 1

        # 2. DataSheet (Exact Extracted Raw Data)
        ws_data = wb.create_sheet(title="DataSheet")
        headers = [
            "Page_No", "Customer_Name", "Product", "Quantity", "Unit_Price",
            "Base_Amount", "Discount", "Taxable_Amount", "GST_Amount",
            "Money_Received", "Description"
        ]
        ws_data.append(headers)

        for idx, item in enumerate(pages_data, start=2):
            ws_data.append([
                item["Page_No"],
                item["Customer_Name"],
                item["Product"],
                item["Quantity"],
                item["Unit_Price"],
                item["Base_Amount"],
                item["Discount"],
                item["Taxable_Amount"],
                item["GST"],
                item["Net_Total"],
                item["Description"]
            ])

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={'Content-Disposition': f'attachment; filename="{file.filename.replace(".pdf", "_Extracted.xlsx")}"'}
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
