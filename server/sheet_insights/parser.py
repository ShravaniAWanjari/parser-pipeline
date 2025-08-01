import openpyxl
import csv
import re
from pathlib import Path

def normalize_sheet_name(sheet_name: str) -> str:
    """Consistent sheet name normalization used across the application"""
    clean_name = re.sub(r'[^\w\-_]', '_', sheet_name.strip())
    clean_name = re.sub(r'_+', '_', clean_name).strip('_')
    return clean_name

def get_sheet_names(file_path):
    try:
        workbook = openpyxl.load_workbook(file_path, read_only=True)
        names = workbook.sheetnames
        workbook.close()
        print(f"Found {len(names)} sheets: {names}")
        return names
    except Exception as e:
        print(f"Failed to load sheet names: {e}")
        return []

def is_empty_row(row):
    return all(cell is None or str(cell).strip() == '' for cell in row)

def has_meaningful_content(row):
    non_empty_cells = [cell for cell in row if cell is not None and str(cell).strip()]
    if len(non_empty_cells) < 2:
        return False
    formula_only = all(str(cell).startswith('=') for cell in non_empty_cells)
    return not (formula_only and len(non_empty_cells) < 3)

def find_data_boundaries(sheet, start_row=6):
    rows = list(sheet.iter_rows(min_row=start_row, values_only=True))
    last_meaningful_row = -1
    for i, row in enumerate(rows):
        if has_meaningful_content(row):
            last_meaningful_row = i
    return rows[:last_meaningful_row + 4] if last_meaningful_row >= 0 else rows[:10]

def extract_csv(file_path, output_dir, sheets_to_process=None, skip_first_sheet=True):
    all_sheet_names = get_sheet_names(file_path)
    if not all_sheet_names:
        print("❌ No sheets found in Excel file")
        return [], {}

    target_sheets = sheets_to_process if sheets_to_process else (
        all_sheet_names[1:] if skip_first_sheet else all_sheet_names
    )

    print(f"Processing {len(target_sheets)} sheet(s): {target_sheets}")

    workbook = openpyxl.load_workbook(file_path, read_only=True)
    csv_paths = []
    name_mapping = {}

    # Create a mapping of exact sheet names for better matching
    exact_sheet_names = {sheet.strip(): sheet for sheet in workbook.sheetnames}
    
    for sheet_name in target_sheets:
        try:
            print(f"🔄 Processing sheet: '{sheet_name}'")
            
            # Try to find the exact sheet in workbook
            actual_sheet_name = None
            if sheet_name in workbook.sheetnames:
                actual_sheet_name = sheet_name
            elif sheet_name.strip() in exact_sheet_names:
                actual_sheet_name = exact_sheet_names[sheet_name.strip()]
            else:
                # Try to find by stripped comparison
                for wb_sheet in workbook.sheetnames:
                    if wb_sheet.strip() == sheet_name.strip():
                        actual_sheet_name = wb_sheet
                        break
            
            if not actual_sheet_name:
                print(f"❌ Sheet '{sheet_name}' not found in workbook")
                print(f"   Available sheets: {workbook.sheetnames}")
                continue
                
            print(f"   Using actual sheet name: '{actual_sheet_name}'")
            sheet = workbook[actual_sheet_name]
            
            
            rows = find_data_boundaries(sheet, start_row=6)
            non_empty_rows = [row for row in rows if not is_empty_row(row)]
            
            if not non_empty_rows:
                print(f"⚠️ No meaningful content found in sheet: {sheet_name}")
                continue

            max_cols = max(len(row) for row in non_empty_rows)
            clean_name = normalize_sheet_name(sheet_name)
            csv_path = output_dir / f"{clean_name}.csv"

            with open(csv_path, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                for row in rows:
                    padded_row = list(row) + [None] * (max_cols - len(row))
                    cleaned = [
                        '' if cell is None or str(cell).strip() == '' else str(cell).strip()
                        for cell in padded_row
                    ]
                    writer.writerow(cleaned)

            csv_paths.append(csv_path)
            name_mapping[clean_name] = actual_sheet_name  # Use the actual sheet name found
            print(f"✅ Created CSV: {csv_path}")
            
        except KeyError as ke:
            print(f"❌ Sheet '{sheet_name}' not found in workbook")
            print(f"   Available sheets: {workbook.sheetnames}")
            # Try to find a close match
            available_sheets = workbook.sheetnames
            close_matches = [s for s in available_sheets if sheet_name.strip() in s or s in sheet_name.strip()]
            if close_matches:
                print(f"   Possible matches: {close_matches}")
            continue
        except Exception as e:
            print(f"❌ Failed to process sheet {sheet_name}: {e}")
            continue

    workbook.close()
    print(f"📁 Successfully saved {len(csv_paths)} CSV files")
    
    if not csv_paths:
        print("⚠️ WARNING: No CSV files were generated. This could be due to:")
        print("   - All sheets have no meaningful content")
        print("   - All target sheets were excluded or not found")
        print("   - Processing errors occurred for all sheets")
    
    return csv_paths, name_mapping