import subprocess
import platform
import tempfile
import os
import zipfile
from datetime import datetime
from fastapi import FastAPI
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from typing import List
from openpyxl import Workbook

app = FastAPI()


if platform.system() == "Windows":
    LIBREOFFICE = r"C:\Program Files\LibreOffice\program\soffice.exe"
    WORKDIR = r"C:\tmp\workdir"
else:
    LIBREOFFICE = "libreoffice"
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


    # excel_path = path to your original file
    excel_dir = os.path.dirname(excel_path)

    # Create a temp directory inside the same folder
    with tempfile.TemporaryDirectory(dir=excel_dir) as tmpdir:

        # Run LibreOffice conversion inside the temp dir
        subprocess.run(
            [
                LIBREOFFICE,
                "--headless",
                "--convert-to", "xlsx",
                "--outdir", tmpdir,
                excel_path
            ],
            check=True
        )

        # Find the converted file
        converted = [
            f for f in os.listdir(tmpdir)
            if f.endswith(".xlsx")
        ][0]

        converted_path = os.path.join(tmpdir, converted)

        # Replace original file with recalculated one
        os.replace(converted_path, excel_path)

        # Temp directory is automatically deleted here


    if platform.system() == "Windows":
        # The ':Zone.Identifier' suffix targets the Alternate Data Stream
        with open(f"{excel_path}:Zone.Identifier", "w") as f:
            f.write("[ZoneTransfer]\nZoneId=3")

    # -----------------------------
    # 3. CONVERT RECALCULATED XLSX → PDF
    # -----------------------------
    subprocess.run(
        [
            LIBREOFFICE,
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

    zip_path = os.path.join(WORKDIR, f"{base_name}.zip") 
    with open(zip_path, "wb") as file_stream: 
        with zipfile.ZipFile(file_stream, "w", zipfile.ZIP_DEFLATED) as zipf:
            zipf.write(excel_path, arcname=f"{base_name}.xlsx")
            zipf.write(pdf_path,   arcname=f"{base_name}.pdf")

    if platform.system() != "Windows":
        os.remove(excel_path)
        os.remove(pdf_path)

    return StreamingResponse(
        open(zip_path, "rb"),
        media_type="application/zip",
        headers={"Content-Disposition": f"attachment; filename={base_name}.zip"}
    )
