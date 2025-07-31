from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse, RedirectResponse
from fastapi.middleware.cors import CORSMiddleware
from pathlib import Path
import shutil
import os
import json
from concurrent.futures import ThreadPoolExecutor, as_completed
import openpyxl
import time
from sheet_insights.parser import extract_markdown, get_sheet_names
from sheet_insights.insights import get_insights
from sheet_insights.general_summary import generate_general_insights
from sheet_insights.additional_insights import generate_additional_insights
from sheet_insights.kpi_dashboard import get_all_supplier_kpi_json

app = FastAPI()

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # In production, replace with specific origins
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
uploaded_file = next(Path(UPLOAD_DIR).glob("*.xlsx"))


for folder in [UPLOAD_DIR, MARKDOWN_DIR, RESULTS_DIR]:
    folder.mkdir(parents=True, exist_ok=True)

@app.get("/")
def read_root():
    return RedirectResponse(url='/docs')



@app.post("/upload_excel/")
async def upload_excel(file: UploadFile = File(...)):
    if not file.filename.endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Only .xlsx files are supported.")

    # Save uploaded file
    file_path = UPLOAD_DIR / file.filename
    with open(file_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    print(f"📁 Uploaded file: {file.filename}")


    try:
        # Get all sheet names for validation
        all_sheet_names = get_sheet_names(str(file_path))
        if not all_sheet_names:
            raise HTTPException(status_code=400, detail="No sheets found in the Excel file")
        
        print(f"📋 Found sheets: {all_sheet_names}")

        # Skip first two sheets: "Average Summary" and "Analysis SUMMARY"
        if len(all_sheet_names) > 2:
            print(f"📋 Skipping first two sheets: '{all_sheet_names[0]}' and '{all_sheet_names[1]}'")
            sheets_to_process = all_sheet_names[2:]
        elif len(all_sheet_names) > 1:
            print(f"📋 Skipping first sheet: '{all_sheet_names[0]}'")
            sheets_to_process = all_sheet_names[1:]
        else:
            print(f"📋 Processing single sheet: '{all_sheet_names[0]}'")
            sheets_to_process = all_sheet_names

        print(f"🔄 Sheets to process: {sheets_to_process}")

        markdown_paths, name_mapping = extract_markdown(
            str(file_path),
            MARKDOWN_DIR,
            sheets_to_process=sheets_to_process,
            skip_first_sheet=False,
            max_empty_rows=3  # ← New parameter to control empty row limit
        )

        if not markdown_paths:
            raise HTTPException(status_code=400, detail="No sheets could be processed")

        print(f"📝 Generated {len(markdown_paths)} markdown files")
        print(f"🗺️ Name mapping: {name_mapping}")

        get_all_supplier_kpi_json()
        # Process each sheet for insights with optimized parallel processing
        def process_sheet(markdown_file):
            try:
                start_time = time.time()

                with open(markdown_file, "r", encoding="utf-8") as f:
                    text = f.read()

                # Get the original sheet name from mapping
                clean_name = markdown_file.stem
                original_sheet_name = name_mapping.get(clean_name, clean_name)

                print(f"🔍 Processing insights for: '{original_sheet_name}' (file: {markdown_file.name})")

                # Generate insights using original sheet name
                insight = get_insights(text, original_sheet_name)

                processing_time = time.time() - start_time
                print(f"⏱️ Processed '{original_sheet_name}' in {processing_time:.2f}s")

                return original_sheet_name, insight

            except Exception as e:
                print(f"❌ Error processing {markdown_file.name}: {e}")
                return name_mapping.get(markdown_file.stem, markdown_file.stem), None

        insights = {}
        print(f"🚀 Starting insight generation for {len(markdown_paths)} sheets...")

        optimal_workers = min(4, len(markdown_paths)) 

        print(f"🔧 Using {optimal_workers} parallel workers for processing")

        start_time = time.time()

        with ThreadPoolExecutor(max_workers=optimal_workers) as executor:
            # Submit all tasks and track progress
            future_to_file = {executor.submit(process_sheet, file): file for file in markdown_paths}
            results = []

            # Process completed tasks as they finish
            for i, future in enumerate(as_completed(future_to_file), 1):
                try:
                    result = future.result()
                    results.append(result)
                    print(f"📈 Progress: {i}/{len(markdown_paths)} sheets processed")
                except Exception as e:
                    file = future_to_file[future]
                    print(f"❌ Error processing {file.name}: {e}")
                    results.append((name_mapping.get(file.stem, file.stem), None))

        total_time = time.time() - start_time
        print(f"⏱️ Total processing time: {total_time:.2f}s")

        # Collect results
        processed_count = 0
        for sheet_name, insight in results:
            if insight:
                insights[sheet_name] = insight
                processed_count += 1
                print(f"✅ Generated insights for: '{sheet_name}'")
            else:
                print(f"❌ Failed to generate insights for: '{sheet_name}'")

        print(f"📊 Successfully generated insights for {processed_count}/{len(results)} sheets")

        # Save insights to file
        with open(INSIGHTS_FILE, "w", encoding='utf-8') as f:
            json.dump(insights, f, indent=2, ensure_ascii=False)

        print(f"💾 Saved insights to: {INSIGHTS_FILE}")

        # Generate general insights
        print(f"🔄 Generating general insights...")
        general = generate_general_insights(str(INSIGHTS_FILE), str(GENERAL_INSIGHTS_FILE))

        # Load insights for response
        with open(INSIGHTS_FILE, "r", encoding="utf-8") as f:
            insights_content = json.load(f)


        print(f"🎉 Processing completed successfully!")

        return {
            "message": f"Successfully processed {len(sheets_to_process)} sheets",
            "processed_sheets": list(insights.keys()),
            "insights": insights_content,
            "general-insights": general
        }

    except Exception as e:
        print(f"❌ Error during processing: {e}")
        raise HTTPException(status_code=500, detail=f"Processing failed: {str(e)}")

@app.get('/download/supplier_kpi_file')
def download_general():
    if not OUTPUT_JSON.exists():
        raise HTTPException(status_code=404, detail='KPI file not generated')
    return FileResponse(path=OUTPUT_JSON, filename='supplier_kpi.json', media_type='application/json')



if __name__ == "__main__":
    import uvicorn

    print("🚀 Starting FastAPI server on port 8001...")
    print("📍 Server will be available at: http://localhost:8001")
    print("📖 API documentation available at: http://localhost:8001/docs")
    print("🔄 Press Ctrl+C to stop the server")

    # Run the server on port 8001
    uvicorn.run(
        "app:app",  # Use import string for reload to work properly
        host="0.0.0.0",
        port=8001,
        reload=True,  # Enable auto-reload for development
        log_level="info"
    )

