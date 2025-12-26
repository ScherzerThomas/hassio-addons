import subprocess
import os
import zipfile
from datetime import datetime
from fastapi import FastAPI
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from typing import List
from openpyxl import Workbook
from io import BytesIO

app = FastAPI()

WORKDIR = "/app/workdir"
os.makedirs(WORKDIR, exist_ok=True)

class Item(BaseModel):
    name: str
    price: float
    qty: float

class Payload(BaseModel):
    items: List[Item]

@app.post("/generate-excel-pdf")
def generate_excel_pdf(payload: Payload):
    # -----------------------------
    # 0. GENERATE TIMESTAMPED NAME
    # -----------------------------
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_name = f"REPORT_{timestamp}"

    excel_path = os.path.join(WORKDIR, f"{base_name}.xlsx")
    pdf_path   = os.path.join(WORKDIR, f"{base_name}.pdf")
    recalc_path = os.path.join(WORKDIR, f"{base_name}_recalc.xlsx")

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

    wb.save(excel_path)

    # -----------------------------
    # 2. RECALCULATE FORMULAS AND REWRITE XLSX
    # -----------------------------
    subprocess.run(
        [
            "libreoffice",
            "--headless",
            "--calc",
            "--convert-to", "xlsx",
            "--outdir", WORKDIR,
            excel_path
        ],
        check=True
    )

    # Replace original Excel with recalculated one
    os.replace(recalc_path, excel_path)

    # -----------------------------
    # 3. CONVERT RECALCULATED XLSX → PDF
    # -----------------------------
    subprocess.run(
        [
            "libreoffice",
            "--headless",
            "--calc",
            "--convert-to", "pdf",
            "--outdir", WORKDIR,
            excel_path
        ],
        check=True
    )

    # -----------------------------
    # 4. PACKAGE BOTH FILES INTO ZIP
    # -----------------------------
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        zipf.write(excel_path, arcname=f"{base_name}.xlsx")
        zipf.write(pdf_path,   arcname=f"{base_name}.pdf")

    zip_buffer.seek(0)
    
    os.remove(excel_path)
    os.remove(pdf_path)

    return StreamingResponse(
        zip_buffer,
        media_type="application/zip",
        headers={"Content-Disposition": f"attachment; filename={base_name}.zip"}
    )
