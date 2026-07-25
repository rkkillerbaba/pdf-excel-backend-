import io
import re
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import StreamingResponse, HTMLResponse
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation

app = FastAPI(title="PDF to Master Sheet Converter")

# -------------------------------------------------------------
# 🎨 WEB UI INTERFACE (HTML/CSS)
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
        <h2>📄 PDF to Master Sheet</h2>
        <p>100-Page PDF Upload Karein & Dynamic Excel Template Download Karein</p>

        <form id="uploadForm">
            <div class="upload-area" onclick="document.getElementById('pdfFile').click()">
                <div class="icon">📁</div>
                <div style="font-weight: bold; font-size: 14px;">Select or Drag PDF File Here</div>
                <div id="file-name">No file selected</div>
            </div>
            <input type="file" id="pdfFile" class="file-input" accept=".pdf" onchange="showFileName()" required>
            
            <button type="submit" id="submitBtn" class="btn">⚡ Convert & Download Excel</button>
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
            submitBtn.innerText = "Processing PDF...";
            spinner.style.display = "block";

            const formData = new FormData();
            formData.append("file", fileInput.files[0]);

            try {
                const response = await fetch("/convert-pdf/", {
                    method: "POST",
                    body: formData
                });

                if (!response.ok) {
                    throw new Error("Conversion failed!");
                }

                const blob = await response.blob();
                const downloadUrl = window.URL.createObjectURL(blob);
                const a = document.createElement("a");
                a.href = downloadUrl;
                a.download = fileInput.files[0].name.replace(".pdf", "_MasterDashboard.xlsx");
                document.body.appendChild(a);
                a.click();
                a.remove();

            } catch (err) {
                alert("Error processing PDF file. Please try again.");
            } finally {
                submitBtn.disabled = false;
                submitBtn.innerText = "⚡ Convert & Download Excel";
                spinner.style.display = "none";
            }
        });
    </script>
</body>
</html>
"""

@app.get("/", response_class=HTMLResponse)
def serve_ui():
    """Serves the Web UI for uploading PDFs directly from the browser."""
    return HTML_PAGE

# -------------------------------------------------------------
# ⚙️ PDF PARSING & EXCEL ENGINE
# -------------------------------------------------------------
def parse_pdf_text(text: str, page_num: int) -> dict:
    row_data = {"Page_No": page_num}
    patterns = {
        "Customer_Name": r"Customer Name:\s*(.*)",
        "Product": r"Product:\s*(.*)",
        "Quantity": r"Quantity:\s*(\d+)",
        "Unit_Price": r"Unit Price:\s*(?:Rs\.|₹)?\s*([\d\.]+)",
        "Discount_Pct": r"Discount:\s*([\d\.]+)\%?",
        "GST_Pct": r"GST Rate:\s*([\d\.]+)\%?",
        "Description": r"Description:\s*(.*)"
    }
    
    for key, pattern in patterns.items():
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            val = match.group(1).strip()
            if key == "Quantity":
                row_data[key] = int(val)
            elif key in ["Unit_Price", "Discount_Pct", "GST_Pct"]:
                row_data[key] = float(val)
            else:
                row_data[key] = val
        else:
            row_data[key] = 0 if key in ["Quantity", "Unit_Price", "Discount_Pct", "GST_Pct"] else ""
            
    return row_data

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
                text = page.extract_text() or ""
                pages_data.append(parse_pdf_text(text, i))
        
        wb = openpyxl.Workbook()
        
        # 1. PageViewer Sheet (Frontend Dynamic Dashboard)
        ws_view = wb.active
        ws_view.title = "PageViewer"
        ws_view.views.sheetView[0].showGridLines = True

        header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        white_title = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
        label_font = Font(name="Calibri", size=11, bold=True, color="333333")
        formula_font = Font(name="Calibri", size=11, bold=False, color="000000")
        accent_font = Font(name="Calibri", size=12, bold=True, color="28A745")

        ws_view.merge_cells("B1:D1")
        ws_view["B1"] = "ACTIVE PAGE DATA & CALCULATION TEMPLATE"
        ws_view["B1"].font = white_title
        ws_view["B1"].fill = header_fill
        ws_view["B1"].alignment = Alignment(horizontal="center", vertical="center")

        ws_view["B3"] = "Select Active Page:"
        ws_view["B3"].font = label_font
        ws_view["C3"] = 1
        ws_view["C3"].font = Font(size=12, bold=True, color="007BFF")

        max_p = len(pages_data)
        dv = DataValidation(type="list", formula1=f'"{",".join(str(p) for p in range(1, max_p + 1))}"', allow_blank=False)
        ws_view.add_data_validation(dv)
        dv.add("C3")

        calc_rows = [
            ("Customer Name", '=XLOOKUP(C3, DataSheet!A:A, DataSheet!B:B, "N/A")'),
            ("Product Name", '=XLOOKUP(C3, DataSheet!A:A, DataSheet!C:C, "N/A")'),
            ("Quantity", '=XLOOKUP(C3, DataSheet!A:A, DataSheet!D:D, 0)'),
            ("Unit Price (₹)", '=XLOOKUP(C3, DataSheet!A:A, DataSheet!E:E, 0)'),
            ("Base Amount (₹)", '=C7*C8'),
            ("Discount Rate (%)", '=XLOOKUP(C3, DataSheet!A:A, DataSheet!G:G, 0)'),
            ("Discount Amount (₹)", '=C9*(C10/100)'),
            ("Taxable Amount (₹)", '=C9-C11'),
            ("CGST Amount (9%) (₹)", '=C12*0.09'),
            ("SGST Amount (9%) (₹)", '=C12*0.09'),
            ("Total GST Amount (₹)", '=C13+C14'),
            ("Extra Shipping (₹)", 0),
            ("Gross Total (₹)", '=C12+C15+C16'),
            ("Round Off Adjustment", '=ROUND(C17,0)-C17'),
            ("FINAL PAYABLE (₹)", '=ROUND(C17,0)'),
            ("Description", '=XLOOKUP(C3, DataSheet!A:A, DataSheet!M:M, "")')
        ]

        row_idx = 5
        for label, formula in calc_rows:
            ws_view[f"B{row_idx}"] = label
            ws_view[f"B{row_idx}"].font = label_font
            ws_view[f"C{row_idx}"] = formula
            ws_view[f"C{row_idx}"].font = accent_font if "FINAL PAYABLE" in label else formula_font
            row_idx += 1

        # 2. DataSheet (Backend Master Dataset)
        ws_data = wb.create_sheet(title="DataSheet")
        headers = ["Page_No", "Customer_Name", "Product", "Quantity", "Unit_Price", "Base_Amount", "Discount_Pct", "Discount_Amount", "Taxable_Amount", "GST_Pct", "GST_Amount", "Money_Received", "Description"]
        ws_data.append(headers)

        for idx, item in enumerate(pages_data, start=2):
            base_amt = f"=D{idx}*E{idx}"
            disc_amt = f"=F{idx}*(G{idx}/100)"
            taxable_amt = f"=F{idx}-H{idx}"
            gst_amt = f"=I{idx}*(J{idx}/100)"
            net_amt = f"=I{idx}+K{idx}"

            ws_data.append([
                item["Page_No"], item["Customer_Name"], item["Product"],
                item["Quantity"], item["Unit_Price"], base_amt,
                item["Discount_Pct"], disc_amt, taxable_amt,
                item["GST_Pct"], gst_amt, net_amt, item["Description"]
            ])

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={'Content-Disposition': f'attachment; filename="{file.filename.replace(".pdf", "_CalculatedMaster.xlsx")}"'}
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
