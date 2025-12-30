import subprocess
import platform
import tempfile
import os
import json
import zipfile
from datetime import datetime
from fastapi import FastAPI
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from typing import List
from openpyxl import Workbook, load_workbook
from openpyxl.workbook.defined_name import DefinedName

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

def set_named_value(workbook, name_symbol, new_value):
    """
    Updates the value of a symbolic name in an openpyxl workbook.
    Handles both cell-based named ranges and named constants.
    """
    if name_symbol not in workbook.defined_names:
        raise ValueError(f"Named symbol '{name_symbol}' not found in workbook.")

    defn = workbook.defined_names[name_symbol]
    dests = list(defn.destinations)

    if dests:
        # Case 1: The name refers to a cell or range (Named Cell)
        for sheet_title, cell_address in dests:
            ws = workbook[sheet_title]
            # If it's a single cell reference like $A$1
            if ":" not in cell_address:
                ws[cell_address] = new_value
            else:
                # If it's a range like $A$1:$B$2, updates all cells in that range
                for row in ws[cell_address]:
                    for cell in row:
                        cell.value = new_value
    else:
        # Case 2: The name is a constant or formula (Named Constant)
        # Directly overwrite the definition with the new numeric value
        workbook.defined_names[name_symbol] = DefinedName(name_symbol, attr_text=str(new_value))


@app.post("/generate-excel-pdf")
def generate_excel_pdf(payload: Payload):
    # -----------------------------
    # 0. GENERATE TIMESTAMPED NAME
    # -----------------------------
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_name = f"REPORT_{timestamp}"

    data_path = os.path.join(WORKDIR, f"data.json")
    template_path = os.path.join(WORKDIR, f"Template.xlsx")
    excel_path = os.path.join(WORKDIR, f"{base_name}.xlsx")
    pdf_path   = os.path.join(WORKDIR, f"{base_name}.pdf")
    recalc_path = os.path.join(WORKDIR, f"{base_name}_recalc.xlsx")

    # -----------------------------
    # 1. CREATE EXCEL WORKBOOK
    # -----------------------------
    wb = load_workbook(template_path)
    
    # 1. Load your nested JSON data
    with open(data_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

        groups = ["GM", "M1", "M2", "GK"]
        for group in groups:
            for key, value in data[group].items():
                if key == "Aktiv": continue
                set_named_value(wb, f"{group}_{key}", value)
            
    #set_named_value(wb, "GM_Anschaffungspreis", 443322)


    #ws.append(["Name", "Price", "Qty", "Total"])
#
    #for idx, item in enumerate(payload.items, start=2):
    #    ws[f"A{idx}"] = item.name
    #    ws[f"B{idx}"] = item.price
    #    ws[f"C{idx}"] = item.qty
    #    ws[f"D{idx}"] = f"=B{idx}*C{idx}"

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
