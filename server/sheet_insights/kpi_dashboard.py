import os
import json
import time
from pathlib import Path
from openai import AzureOpenAI
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Azure OpenAI Client
client = AzureOpenAI(
    api_key=os.getenv("AZURE_OPENAI_API_KEY"),
    api_version="2023-07-01-preview",
    azure_endpoint=os.getenv("AZURE_ENDPOINT")
)

# Define standard KPI mapping
kpi_map = {
    "Safety- Accident data": "accidents",
    "Production loss due to Material shortage": "productionLossHrs",
    "OK delivery cycles- as per delivery calculation sheet of ACMA (%)": "okDeliveryPercent",
    "Number of trips / month": "trips",
    "Qty Shipped / month": "quantityShipped",
    "No of Parts/ Trip": "partsPerTrip",
    "Vehicle turnaround time": "vehicleTAT",
    "Machin break down Hrs": "machineDowntimeHrs",
    "No of Machines breakdown": "machineBreakdowns"
}

# Main KPI extraction function
def get_all_supplier_kpi_json(markdown_folder: Path = Path("results/markdown_output"), output_path: Path = Path("results/final_supplier_kpis.json")):
    final_output = {
        "generatedOn": "2025-07-30",
        "kpiMetadata": {
            "unitDescriptions": {
                "accidents": "Number of safety incidents reported",
                "productionLossHrs": "Production hours lost due to supplier-caused material shortage",
                "okDeliveryPercent": "Percentage of OK deliveries based on ACMA standards",
                "trips": "Number of shipment trips completed per month",
                "quantityShipped": "Number of parts shipped by the supplier",
                "partsPerTrip": "Efficiency metric showing avg. parts shipped per trip",
                "vehicleTAT": "Average vehicle turnaround time at the plant (in hours)",
                "machineDowntimeHrs": "Machine breakdown time (in hours)",
                "machineBreakdowns": "Number of machine breakdowns"
            }
        }
    }

    deployment_name = os.getenv("AZURE_OPENAI_DEPLOYMENT")
    if not deployment_name:
        print("❌ AZURE_OPENAI_DEPLOYMENT environment variable not set")
        return None

    markdown_files = list(markdown_folder.glob("*.md"))

    for md_path in markdown_files:
        supplier_name = md_path.stem.replace("- Supplier Partner Performance Matrix", "").strip()
        markdown_content = md_path.read_text(encoding="utf-8")

        user_prompt = f"""
Given the markdown content below, extract KPIs and convert into this JSON structure:

{{
  "supplier": "{supplier_name}",
  "kpis": {{
    "accidents": {{"Jan": 0, "Feb": 0, "Mar": 0, "Apr": 0}},
    "productionLossHrs": {{"Jan": 0, ...}},
    ...
  }}
}}

Use these keys only: {json.dumps(kpi_map)}
Use null for missing values.
Markdown:
{markdown_content}
"""

        try:
            start_time = time.time()
            response = client.chat.completions.create(
                model=deployment_name,
                messages=[
                    {"role": "system", "content": "You are a data extractor AI. Convert supplier markdown tables into KPI-wise JSON."},
                    {"role": "user", "content": user_prompt}
                ],
                temperature=0
            )
            elapsed = time.time() - start_time
            print(f"✅ Processed {supplier_name} in {elapsed:.2f}s")
        except Exception as e:
            print(f"❌ Failed to process {supplier_name}: {e}")
            continue

        try:
            ai_json = json.loads(response.choices[0].message.content)
        except Exception as e:
            print(f"❌ Failed to parse JSON for {supplier_name}: {e}")
            continue

        for kpi, values in ai_json.get("kpis", {}).items():
            if kpi not in final_output:
                final_output[kpi] = {}
            final_output[kpi][supplier_name] = values

    # Write final JSON
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(final_output, f, indent=2)

    print(f"✅ All supplier KPIs aggregated to {output_path}")
