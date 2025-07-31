import re
from pathlib import Path
import openpyxl

def get_sheet_names(file_path):
    """Extract all sheet names from Excel file"""
    try:
        workbook = openpyxl.load_workbook(file_path, read_only=True)
        names = workbook.sheetnames
        workbook.close()
        print(f"📋 Found {len(names)} sheets: {names}")
        return names
    except Exception as e:
        print(f"❌ Failed to load sheet names: {e}")
        return []

def format_metadata(sheet):
    """Extract and format metadata from the top few rows of the Excel sheet"""
    metadata = []
    for row in sheet.iter_rows(min_row=1, max_row=5, values_only=True):
        line = ' '.join([str(cell).strip() for cell in row if cell])
        if line:
            metadata.append(line)
    return metadata

def is_empty_row(row):
    """Check if a row is essentially empty (all None, empty strings, or whitespace)"""
    return all(cell is None or str(cell).strip() == '' for cell in row)

def has_meaningful_content(row):
    """Check if row has meaningful content (not just formulas or minimal data)"""
    non_empty_cells = [cell for cell in row if cell is not None and str(cell).strip()]
    
    # If less than 2 non-empty cells, consider it not meaningful
    if len(non_empty_cells) < 2:
        return False
    
    # Check if it's just formulas without actual data
    formula_only = all(str(cell).startswith('=') for cell in non_empty_cells)
    if formula_only and len(non_empty_cells) < 3:
        return False
    
    return True

def find_data_boundaries(sheet, start_row=6):
    """Find the actual data boundaries to avoid processing too many empty rows"""
    rows = list(sheet.iter_rows(min_row=start_row, values_only=True))
    
    # Find last row with meaningful content
    last_meaningful_row = -1
    for i, row in enumerate(rows):
        if has_meaningful_content(row):
            last_meaningful_row = i
    
    # If we found meaningful content, include a few extra rows for spacing
    if last_meaningful_row >= 0:
        return rows[:last_meaningful_row + 4]  # +4 to allow 3-4 empty rows after data
    else:
        # If no meaningful content found, return first 10 rows as fallback
        return rows[:10]

def extract_table(sheet):
    """Extract and format the core table starting from a specific row"""
    start_row = 6  # Adjust based on where your actual data starts
    
    # Get bounded rows instead of all rows
    rows = find_data_boundaries(sheet, start_row)
    
    if not rows:
        return []
    
    # Find max number of columns from non-empty rows
    non_empty_rows = [row for row in rows if not is_empty_row(row)]
    if not non_empty_rows:
        return []
    
    max_cols = max(len(row) for row in non_empty_rows)
    
    # Process rows and remove consecutive empty rows (keep max 3-4)
    processed_rows = []
    consecutive_empty = 0
    max_consecutive_empty = 3
    
    for row in rows:
        # Pad row to match max columns
        padded_row = list(row) + [''] * (max_cols - len(row))
        
        if is_empty_row(padded_row):
            consecutive_empty += 1
            if consecutive_empty <= max_consecutive_empty:
                processed_rows.append(padded_row)
        else:
            consecutive_empty = 0
            processed_rows.append(padded_row)
    
    # Convert to markdown
    markdown_lines = []
    for i, row in enumerate(processed_rows):
        # Clean cell content
        cells = []
        for cell in row:
            if cell is None:
                cells.append('')
            else:
                # Clean the cell content
                cell_str = str(cell).replace('\n', ' ').strip()
                # Truncate very long formula strings
                if cell_str.startswith('=') and len(cell_str) > 50:
                    cell_str = cell_str[:47] + '...'
                cells.append(cell_str)
        
        line = '| ' + ' | '.join(cells) + ' |'
        markdown_lines.append(line)
        
        # Add header separator after first row if it seems like a header
        if i == 0 and has_meaningful_content(row):
            markdown_lines.append('| ' + ' | '.join(['---'] * len(cells)) + ' |')
    
    return markdown_lines

def clean_markdown_post_process(markdown_path, max_empty_lines=3):
    """Post-process markdown file to remove excessive empty table rows"""
    try:
        with open(markdown_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        
        cleaned_lines = []
        consecutive_empty_table_rows = 0
        
        for line in lines:
            # Check if line is an empty table row (only | and spaces/empty cells)
            stripped = line.strip()
            if stripped.startswith('|') and stripped.endswith('|'):
                # Remove outer pipes and split
                content = stripped[1:-1].split('|')
                # Check if all cells are empty or just whitespace
                if all(cell.strip() == '' for cell in content):
                    consecutive_empty_table_rows += 1
                    if consecutive_empty_table_rows <= max_empty_lines:
                        cleaned_lines.append(line)
                    # Skip this line if we've exceeded the limit
                else:
                    consecutive_empty_table_rows = 0
                    cleaned_lines.append(line)
            else:
                consecutive_empty_table_rows = 0
                cleaned_lines.append(line)
        
        # Write back the cleaned content
        with open(markdown_path, 'w', encoding='utf-8') as f:
            f.writelines(cleaned_lines)
        
        print(f"🧹 Cleaned excessive empty rows from {markdown_path.name}")
        
    except Exception as e:
        print(f"⚠️ Failed to clean {markdown_path}: {e}")

def extract_markdown(file_path, output_dir, sheets_to_process=None, skip_first_sheet=True, max_empty_rows=3):
    """
    Enhanced extract_markdown function with empty row filtering
    
    Args:
        file_path: Path to Excel file
        output_dir: Directory to save markdown files
        sheets_to_process: List of specific sheet names to process (optional)
        skip_first_sheet: Whether to skip the first sheet (default True)
        max_empty_rows: Maximum consecutive empty rows to keep (default 3)
    
    Returns:
        tuple: (list of markdown file paths, name mapping dict)
    """
    all_sheet_names = get_sheet_names(file_path)
    if not all_sheet_names:
        return [], {}

    if sheets_to_process:
        target_sheets = sheets_to_process
    else:
        target_sheets = all_sheet_names[1:] if skip_first_sheet else all_sheet_names

    print(f"📋 Processing {len(target_sheets)} sheet(s): {target_sheets}")
    
    wb = openpyxl.load_workbook(file_path, read_only=True)
    markdown_paths = []
    name_mapping = {}

    for sheet_name in target_sheets:
        try:
            sheet = wb[sheet_name]
            clean_sheet_name = re.sub(r'[^\w\-_]', '_', sheet_name.strip())
            clean_sheet_name = re.sub(r'_+', '_', clean_sheet_name).strip('_')
            
            print(f"📄 Processing: {sheet_name} -> {clean_sheet_name}")

            # Check if the base markdown file already exists
            base_markdown_path = output_dir / f"{clean_sheet_name}.md"
            if base_markdown_path.exists():
                print(f"⏭️ Skipping {base_markdown_path.name} - already exists")
                markdown_paths.append(base_markdown_path)
                name_mapping[base_markdown_path.stem] = sheet_name
                continue

            # If base doesn't exist, check for numbered versions
            markdown_path = base_markdown_path
            counter = 1
            while markdown_path.exists():
                markdown_path = output_dir / f"{clean_sheet_name}_{counter}.md"
                counter += 1

            # Extract content
            metadata_lines = format_metadata(sheet)
            table_lines = extract_table(sheet)
            
            # Write the markdown file
            with open(markdown_path, "w", encoding="utf-8") as f:
                f.write(f"## {sheet_name} - Supplier Partner Performance Matrix\n\n")
                for line in metadata_lines:
                    f.write(f"- {line}\n")
                f.write("\n")
                for line in table_lines:
                    f.write(line + "\n")

            # Post-process to clean up any remaining excessive empty rows
            clean_markdown_post_process(markdown_path, max_empty_rows)
            
            markdown_paths.append(markdown_path)
            name_mapping[markdown_path.stem] = sheet_name
            print(f"✅ Saved: {markdown_path.name}")
            
        except Exception as e:
            print(f"❌ Failed to process {sheet_name}: {e}")
            continue

    wb.close()
    return markdown_paths, name_mapping