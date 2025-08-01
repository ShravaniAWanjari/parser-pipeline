import os
import json
import time
import csv
from pathlib import Path
from openai import AzureOpenAI
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

client = AzureOpenAI(
    api_key=os.getenv("AZURE_OPENAI_API_KEY"),
    api_version="2023-07-01-preview",
    azure_endpoint=os.getenv("AZURE_ENDPOINT")
)

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

ALL_MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

def find_month_data_columns(csv_rows):
    """
    Find the exact column positions where month data should be mapped
    The data starts immediately after the unit column and includes empty cells as null months
    """
    for row in csv_rows:
        if len(row) > 10:  # Ensure row has enough columns
            # Find unit column (nos, Hrs, %)
            unit_col_index = -1
            for i, cell in enumerate(row):
                if cell.strip().lower() in ['nos', 'hrs', '%']:
                    unit_col_index = i
                    break
            
            if unit_col_index >= 0:
                # Data starts immediately after unit column (including empty cells for null months)
                data_start_col = unit_col_index + 1
                return data_start_col
    
    return 6  # Default fallback

def parse_monthly_data_from_row(row, data_start_col):
    """
    Extract monthly data from a row, preserving null values for empty cells
    """
    monthly_data = {}
    
    # Extract up to 12 months of data starting from data_start_col
    for i, month in enumerate(ALL_MONTHS):
        col_index = data_start_col + i
        
        if col_index < len(row):
            cell_value = row[col_index].strip()
            
            # Handle empty/null values
            if not cell_value or cell_value == '':
                monthly_data[month] = None
            # Handle Excel errors
            elif cell_value.startswith('#') or cell_value in ['#DIV/0!', '#N/A', '#VALUE!']:
                monthly_data[month] = None
            else:
                try:
                    # Try to convert to number
                    if '.' in cell_value:
                        monthly_data[month] = float(cell_value)
                    else:
                        monthly_data[month] = int(cell_value)
                except ValueError:
                    monthly_data[month] = None
        else:
            monthly_data[month] = None
    
    return monthly_data

def get_all_supplier_kpi_json(csv_folder: Path = Path("results/csv_output"), output_path: Path = Path("results/final_supplier_kpis.json")):
    print(f"🔍 Looking for CSV files in: {csv_folder}")
    print(f"📁 CSV folder exists: {csv_folder.exists()}")
    
    if not csv_folder.exists():
        print(f"❌ CSV folder does not exist: {csv_folder}")
        return None
    
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

    csv_files = list(csv_folder.glob("*.csv"))
    print(f"📊 Found {len(csv_files)} CSV files: {[f.name for f in csv_files]}")
    
    if not csv_files:
        print("⚠️ WARNING: No CSV files found for KPI processing")
        return final_output

    processed_count = 0
    for csv_path in csv_files:
        supplier_name = csv_path.stem.replace("- Supplier Partner Performance Matrix", "").strip()
        print(f"🔄 Processing supplier: {supplier_name} from {csv_path.name}")

        try:
            with open(csv_path, newline='', encoding='utf-8') as f:
                reader = csv.reader(f)
                csv_rows = list(reader)
                
            # Check if CSV has meaningful content
            if len(csv_rows) < 2:
                print(f"⚠️ CSV file appears to have minimal content: {csv_path.name}")
                continue
                
        except Exception as e:
            print(f"❌ Failed to read CSV for {supplier_name}: {e}")
            continue

        # Find where monthly data columns start
        data_start_col = find_month_data_columns(csv_rows)
        print(f"📅 Data starts at column {data_start_col} for {supplier_name}")

        # Process each KPI row
        supplier_kpis = {}
        
        for row in csv_rows:
            if len(row) < 5:  # Skip rows that are too short
                continue
                
            # Check if this row contains a KPI we're interested in
            kpi_name = None
            for i in range(min(3, len(row))):  # Check first 3 columns for KPI name
                cell_content = row[i].strip()
                if cell_content in kpi_map:
                    kpi_name = kpi_map[cell_content]
                    break
            
            if kpi_name:
                # Extract monthly data for this KPI
                monthly_data = parse_monthly_data_from_row(row, data_start_col)
                supplier_kpis[kpi_name] = monthly_data
                print(f"  ✓ Extracted {kpi_name}: {sum(1 for v in monthly_data.values() if v is not None)} months of data")

        # Add supplier KPIs to final output
        for kpi_name, monthly_data in supplier_kpis.items():
            if kpi_name not in final_output:
                final_output[kpi_name] = {}
            final_output[kpi_name][supplier_name] = monthly_data

        processed_count += 1

    print(f"📊 Successfully processed {processed_count} suppliers")
    
    # Save the output
    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(final_output, f, indent=2)
        print(f"✅ Saved KPI data to: {output_path}")
    except Exception as e:
        print(f"❌ Failed to save KPI data: {e}")
        return None

    print(f"🎉 KPI processing completed - {processed_count} suppliers processed")
    return final_output