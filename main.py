import io
import re
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import StreamingResponse
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation

app = FastAPI(title="PDF to Dynamic Excel Converter")

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

@app.get("/")
def home():
    return {"status": "Online", "message": "PDF to Master Sheet API is Active"}

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
