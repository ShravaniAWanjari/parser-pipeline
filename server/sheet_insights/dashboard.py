import openpyxl
import json
from pathlib import Path
from collections import defaultdict

def extract_insight_data(file_path, output_path="results/dashboard.json"):
    """
    Extract insights data from Excel file and save to JSON format.
    
    Args:
        file_path: Path to the Excel file
        output_path: Path where the JSON output should be saved (default: results/dashboard.json)
    """
    try:
        # Load workbook
        wb = openpyxl.load_workbook(file_path, read_only=True)
        sheet_names = wb.sheetnames
        
        # Skip first sheet (usually summary sheet)
        if len(sheet_names) > 1:
            sheet_names = sheet_names[1:]
        
        insights = defaultdict(dict)  # {metric: {company: {month: value}}}

        print(f"📊 Processing {len(sheet_names)} sheets from Excel file...")

        for sheet_name in sheet_names:
            try:
                sheet = wb[sheet_name]
                company = sheet_name.strip()
                
                print(f"🔄 Processing sheet: {company}")

                # Find the header row (row 6), then extract month names from header
                header_row = None
                for row in sheet.iter_rows(min_row=1, max_row=10, values_only=True):
                    if row and any(str(cell).strip().title() in ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'] for cell in row if cell):
                        header_row = row
                        break
                
                if not header_row:
                    print(f"⚠️ No month headers found in sheet: {company}")
                    continue
                
                # Convert header to strings and find month indices
                header = [str(cell).strip().title() if cell else "" for cell in header_row]
                months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
                month_indices = []
                
                for i, h in enumerate(header):
                    if h in months:
                        month_indices.append((i, h))
                
                if not month_indices:
                    print(f"⚠️ No valid month columns found in sheet: {company}")
                    continue

                print(f"📅 Found months: {[month for _, month in month_indices]}")

                # Find where the header row is located
                header_row_num = 6  # Default
                for row_num, row in enumerate(sheet.iter_rows(min_row=1, max_row=10, values_only=True), 1):
                    if row and row == header_row:
                        header_row_num = row_num
                        break

                # Process data rows below the header
                data_rows_processed = 0
                for row in sheet.iter_rows(min_row=header_row_num + 1, values_only=True):
                    if not row or len(row) < 2:  # Skip empty or too short rows
                        continue
                    
                    # Get metric name from column B (index 1)
                    metric_cell = row[1] if len(row) > 1 else None
                    if not metric_cell:
                        continue
                    
                    metric = str(metric_cell).strip()
                    if not metric or metric.lower() in ['', 'none', 'null']:
                        continue

                    values = {}
                    
                    # Extract values for each month
                    for idx, month in month_indices:
                        if idx < len(row):
                            val = row[idx]
                            if val is not None:
                                try:
                                    # Try to convert to float
                                    if isinstance(val, (int, float)):
                                        values[month] = float(val)
                                    elif isinstance(val, str):
                                        # Remove any currency symbols, commas, etc.
                                        clean_val = val.replace(',', '').replace('$', '').replace('%', '').strip()
                                        if clean_val and clean_val != '-':
                                            values[month] = float(clean_val)
                                except (ValueError, TypeError):
                                    # Skip non-numeric values
                                    continue

                    # Only add metrics that have at least one valid value
                    if values:
                        insights[metric][company] = values
                        data_rows_processed += 1

                print(f"✅ Processed {data_rows_processed} data rows for {company}")

            except Exception as e:
                print(f"❌ Error processing sheet '{sheet_name}': {e}")
                continue

        # Ensure output directory exists
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Convert defaultdict to regular dict for JSON serialization
        insights_dict = dict(insights)
        for metric in insights_dict:
            insights_dict[metric] = dict(insights_dict[metric])

        # Save output
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(insights_dict, f, indent=2, ensure_ascii=False)

        print(f"✅ Successfully saved {len(insights_dict)} metrics to {output_path}")
        print(f"📈 Metrics extracted: {list(insights_dict.keys())}")
        
        return insights_dict

    except Exception as e:
        print(f"❌ Error processing Excel file: {e}")
        raise

if __name__ == "__main__":
    # Test the function
    import sys
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
        output_path = sys.argv[2] if len(sys.argv) > 2 else "results/dashboard.json"
        extract_insight_data(file_path, output_path)
    else:
        print("Usage: python dashboard.py <excel_file_path> [output_path]")