from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import RedirectResponse
from fastapi.middleware.cors import CORSMiddleware
from pathlib import Path
import shutil
import os
import json
import re

from sheet_insights.parser import extract_csv, get_sheet_names, normalize_sheet_name
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
CSV_DIR = Path("results/csv_output")
INSIGHTS_FILE = Path('results/insights.json')
RESULTS_DIR = Path('results')
OUTPUT_JSON = Path("results/final_supplier_kpis.json")

for folder in [UPLOAD_DIR, CSV_DIR, RESULTS_DIR]:
    folder.mkdir(parents=True, exist_ok=True)

def normalize_filename(sheet_name: str) -> str:
    """Use consistent normalization with parser.py"""
    return normalize_sheet_name(sheet_name) + ".csv"

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

    print(f"📁 Uploaded file saved: {file_path.name}")

    try:
        all_sheet_names = get_sheet_names(str(file_path))
        if not all_sheet_names:
            raise HTTPException(status_code=400, detail="No sheets found in the Excel file")

        print(f"📋 All sheets found in Excel: {all_sheet_names}")

        # Define excluded sheets
        excluded_sheets = ['Average Summary', 'Analysis SUMMARY']
        sheets_to_process = [sheet for sheet in all_sheet_names if sheet.strip() not in excluded_sheets]

        print(f"📌 Sheets selected for processing: {sheets_to_process}")
        
        # Check if we have any sheets to process
        if not sheets_to_process:
            print("⚠️ WARNING: No sheets available for processing after exclusions")
            print(f"   All sheets: {all_sheet_names}")
            print(f"   Excluded: {excluded_sheets}")
            raise HTTPException(
                status_code=400, 
                detail=f"No processable sheets found. Available sheets: {all_sheet_names}. Excluded: {excluded_sheets}"
            )

        # Check for existing CSV files with proper debugging
        existing_csv_files = []
        missing_csv_files = []
        
        for sheet in sheets_to_process:
            csv_path = CSV_DIR / normalize_filename(sheet)
            if csv_path.exists():
                existing_csv_files.append(sheet)
                print(f"✅ Found existing CSV: {csv_path}")
            else:
                missing_csv_files.append(sheet)
                print(f"❌ Missing CSV: {csv_path}")

        print(f"📊 Existing CSV files: {len(existing_csv_files)}")
        print(f"🔨 Missing CSV files: {len(missing_csv_files)}")

        # Generate missing CSV files
        if missing_csv_files:
            print(f"🛠️ Generating CSV files for sheets: {missing_csv_files}")
            csv_paths, name_mapping = extract_csv(
                str(file_path),
                CSV_DIR,
                sheets_to_process=missing_csv_files,
                skip_first_sheet=False,
            )
            
            # Verify CSV files were actually created
            actual_csv_paths = [path for path in csv_paths if path.exists()]
            print(f"✅ Successfully generated CSV files: {[p.name for p in actual_csv_paths]}")
            
            if len(actual_csv_paths) == 0:
                print("❌ ERROR: No CSV files were generated despite processing sheets")
                raise HTTPException(
                    status_code=500, 
                    detail="CSV generation failed - no files were created. This could be due to sheets having no meaningful content."
                )
        else:
            print("⏩ All CSV files already exist. Skipping generation.")
            actual_csv_paths = []
            name_mapping = {}

        # Collect all available CSV files
        all_csv_paths = []
        for sheet in sheets_to_process:
            csv_path = CSV_DIR / normalize_filename(sheet)
            if csv_path.exists():
                all_csv_paths.append(csv_path)

        print(f"📦 Total CSV files available for processing: {len(all_csv_paths)}")
        print(f"📂 CSV files: {[p.name for p in all_csv_paths]}")

        if not all_csv_paths:
            # Provide detailed error information
            error_details = {
                "total_sheets": len(all_sheet_names),
                "excluded_sheets": excluded_sheets,
                "sheets_to_process": sheets_to_process,
                "csv_directory": str(CSV_DIR),
                "expected_files": [normalize_filename(sheet) for sheet in sheets_to_process]
            }
            print(f"❌ DETAILED ERROR INFO: {error_details}")
            
            raise HTTPException(
                status_code=400, 
                detail=f"No CSV files available for processing. Details: {error_details}"
            )

        # Process KPIs and insights
        print("📊 Generating supplier KPI data...")
        supplier_kpi_info = get_all_supplier_kpi_json()
        print("✅ Created final_supplier_kpis.json")

        print("🧠 Generating insights...")
        insights = get_insights()
        with open(INSIGHTS_FILE, "w", encoding='utf-8') as f:
            json.dump(insights, f, indent=2, ensure_ascii=False)
        print(f"✅ Saved insights to: {INSIGHTS_FILE}")

        print("📈 Generating general insights...")
        general = generate_general_insights()

        # Load insights for response
        with open(INSIGHTS_FILE, "r", encoding="utf-8") as f:
            insights_content = json.load(f)

        print("🎉 Processing completed successfully.")

        return {
            "insights": insights_content,
            "general-insights": general,
            "Supplier-KPIs": supplier_kpi_info
        }

    except HTTPException:
        # Re-raise HTTP exceptions as-is
        raise
    except Exception as e:
        print(f"❌ Unexpected error during processing: {e}")
        print(f"❌ Error type: {type(e).__name__}")
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Processing failed: {str(e)}")

if __name__ == "__main__":
    import uvicorn

    print("🚀 Starting FastAPI server on port 8001...")
    uvicorn.run(
        "app:app",
        host="0.0.0.0",
        port=8001,
        reload=True,
        log_level="info"
    )