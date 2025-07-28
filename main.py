from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel
from openpyxl import load_workbook
import os
import uuid
from typing import Dict
import re

app = FastAPI()

@app.get("/")
def read_root():
    return {"message": "Welcome to FastAPI!"}

class TemplateData(BaseModel):
    replacements: Dict[str, str]

def replace_placeholders_in_sheet(sheet, replacements):
    
    placeholder_pattern = re.compile(r'\(\d+\)')
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                matches = placeholder_pattern.findall(cell.value)
                if matches:
                    new_value = cell.value
                    for match in matches:
                        placeholder_key = match[1:-1]
                        if placeholder_key in replacements:
                            new_value = new_value.replace(match, replacements[placeholder_key])
                    cell.value = new_value

@app.post("/generate-excel/")
async def generate_excel(data: TemplateData):
    try:
        template_path = "Copy of Custom_Doc-DOCUMENTS_5855_TEMPLATE.xlsx"
        if not os.path.exists(template_path):
            raise HTTPException(status_code=404, detail="Template file not found")

        wb = load_workbook(template_path)
        
        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
            replace_placeholders_in_sheet(sheet, data.replacements)
        
        output_filename = f"output_{uuid.uuid4()}.xlsx"
        output_path = os.path.join("temp", output_filename)
        os.makedirs("temp", exist_ok=True)
        wb.save(output_path)
        
        return FileResponse(
            path=output_path,
            filename=output_filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error generating Excel file: {str(e)}")

