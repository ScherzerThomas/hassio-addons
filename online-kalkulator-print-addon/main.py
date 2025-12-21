import subprocess
import tempfile
import os
from fastapi import FastAPI
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from typing import List
from openpyxl import Workbook
from io import BytesIO

app = FastAPI()

class Item(BaseModel):
    name: str
    price: float
    qty: float

class Payload(BaseModel):
    items: List[Item]

@app.post("/generate-excel-pdf")
def generate_excel_pdf(payload: Payload):
    # -----------------------------
    # 1. CREATE EXCEL WORKBOOK
    # -----------------------------
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"

    ws.append(["Name", "Price", "Qty", "Total"])

    for idx, item in enumerate(payload.items, start=2):
        ws[f"A{idx}"] = item.name
        ws[f"B{idx}"] = item.price
        ws[f"C{idx}"] = item.qty
        ws[f"D{idx}"] = f"=B{idx}*C{idx}"

    # Save Excel to temporary file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_excel:
        wb.save(tmp_excel.name)
        excel_path = tmp_excel.name

    # -----------------------------
    # 2. CONVERT EXCEL → PDF USING LIBREOFFICE
    # -----------------------------
    output_dir = tempfile.mkdtemp()
    subprocess.run(
        [
            "libreoffice",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", output_dir,
            excel_path
        ],
        check=True
    )

    pdf_path = os.path.join(
        output_dir,
        os.path.basename(excel_path).replace(".xlsx", ".pdf")
    )

    # -----------------------------
    # 3. STREAM PDF BACK
    # -----------------------------
    pdf_file = open(pdf_path, "rb")

    # Cleanup Excel file (PDF cleaned after streaming)
    os.remove(excel_path)

    return StreamingResponse(
        pdf_file,
        media_type="application/pdf",
        headers={"Content-Disposition": "attachment; filename=output.pdf"}
    )
