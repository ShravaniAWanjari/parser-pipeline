from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse, RedirectResponse
from fastapi.middleware.cors import CORSMiddleware
from pathlib import Path
import shutil
import os
import json
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

from sheet_insights.parser import extract_markdown, get_sheet_names
from sheet_insights.insights import get_insights
from sheet_insights.general_summary import generate_general_insights
from sheet_insights.kpi_dashboard import get_all_supplier_kpi_json

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

UPLOAD_DIR = Path('uploads')
MARKDOWN_DIR = Path("results/markdown_output")
INSIGHTS_FILE = Path('results/insights.json')
GENERAL_INSIGHTS_FILE = Path('results/general-info.json')
ADDITIONAL_INSIGHTS_FILE = Path('results/additional-insights.json')
RESULTS_DIR = Path('results')
OUTPUT_JSON = Path("results/final_supplier_kpis.json")

for folder in [UPLOAD_DIR, MARKDOWN_DIR, RESULTS_DIR]:
    folder.mkdir(parents=True, exist_ok=True)

def normalize_filename(sheet_name: str) -> str:
    return sheet_name.strip().replace(" ", "_") + ".md"

@app.get("/")
def read_root():
    return RedirectResponse(url='/docs')

@app.post("/upload_excel/")
async def upload_excel(file: UploadFile = File(...)):
    if not file.filename.endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Only .xlsx files are supported.")

    file_path = UPLOAD_DIR / file.filename
    with open(file_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    print(f"Uploaded file saved: {file_path.name}")

    try:
        all_sheet_names = get_sheet_names(str(file_path))
        if not all_sheet_names:
            raise HTTPException(status_code=400, detail="No sheets found in the Excel file")

        print(f"Sheets found in Excel: {all_sheet_names}")

        excluded_sheets = ['Average Summary', 'Analysis SUMMARY']
        sheets_to_process = [sheet.strip() for sheet in all_sheet_names if sheet.strip() not in excluded_sheets]

        print(f"Sheets selected for processing: {sheets_to_process}")

        existing_md_files = [
            sheet for sheet in sheets_to_process
            if (MARKDOWN_DIR / normalize_filename(sheet)).exists()
        ]
        missing_md_files = [
            sheet for sheet in sheets_to_process
            if sheet not in existing_md_files
        ]

        if missing_md_files:
            print(f"Generating markdown files for sheets: {missing_md_files}")
            markdown_paths, name_mapping = extract_markdown(
                str(file_path),
                MARKDOWN_DIR,
                sheets_to_process=missing_md_files,
                skip_first_sheet=False,
            )
            markdown_paths = [path for path in markdown_paths if path.exists()]
            print(f"Newly generated markdown files: {[p.name for p in markdown_paths]}")
        else:
            print("All markdown files already exist. Skipping generation.")
            markdown_paths = [
                MARKDOWN_DIR / normalize_filename(sheet)
                for sheet in sheets_to_process
            ]
            name_mapping = {sheet: sheet for sheet in sheets_to_process}

        print(f"Total markdown files available: {len(markdown_paths)}")
        print(f"Sheet-to-name mapping: {name_mapping}")

        if not markdown_paths:
            raise HTTPException(status_code=400, detail="No markdown files available for processing.")

        supplier_kpi_info = get_all_supplier_kpi_json()
        print("Created final_supplier_kpis.json")

        insights = get_insights()
        with open(INSIGHTS_FILE, "w", encoding='utf-8') as f:
            json.dump(insights, f, indent=2, ensure_ascii=False)

        print(f"Saved insights to: {INSIGHTS_FILE}")

        print("Generating general insights...")
        general = generate_general_insights()

        with open(INSIGHTS_FILE, "r", encoding="utf-8") as f:
            insights_content = json.load(f)

        print("Processing completed successfully.")

        return {
            "insights": insights_content,
            "general-insights": general,
            "Supplier-KPIs": supplier_kpi_info
        }

    except Exception as e:
        print(f"Error during processing: {e}")
        raise HTTPException(status_code=500, detail=f"Processing failed: {str(e)}")

if __name__ == "__main__":
    import uvicorn

    print("Starting FastAPI server on port 8001...")
    print("Server will be available at: http://localhost:8001")
    print("API documentation available at: http://localhost:8001/docs")
    print("Press Ctrl+C to stop the server")

    uvicorn.run(
        "app:app",
        host="0.0.0.0",
        port=8001,
        reload=True,
        log_level="info"
    )