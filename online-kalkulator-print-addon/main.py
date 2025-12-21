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
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"

    ws.append(["Name", "Price", "Qty", "Total"])

    for idx, item in enumerate(payload.items, start=2):
        ws[f"A{idx}"] = item.name
        ws[f"B{idx}"] = item.price
        ws[f"C{idx}"] = item.qty
        ws[f"D{idx}"] = f"=B{idx}*C{idx}"

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_excel:
        wb.save(tmp_excel.name)
        excel_path = tmp_excel.name

    pdf_path = excel_path.replace(".xlsx", ".pdf")

    subprocess.run(
        ["unoconv", "-f", "pdf", excel_path],
        check=True
    )

    pdf_file = open(pdf_path, "rb")

    os.remove(excel_path)

    return StreamingResponse(
        pdf_file,
        media_type="application/pdf",
        headers={"Content-Disposition": "attachment; filename=output.pdf"}
    )
