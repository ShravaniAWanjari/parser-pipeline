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

def extract_table(sheet):
    """Extract and format the core table starting from a specific row"""
    start_row = 6  # Adjust based on where your actual data starts
    rows = list(sheet.iter_rows(min_row=start_row, values_only=True))

    # Find max number of columns
    max_cols = max(len(row) for row in rows if row)
    padded_rows = [list(row) + [''] * (max_cols - len(row)) for row in rows]

    # Convert to markdown
    markdown_lines = []
    for i, row in enumerate(padded_rows):
        cells = [str(cell).replace('\n', ' ').strip() if cell is not None else '' for cell in row]
        line = '| ' + ' | '.join(cells) + ' |'
        markdown_lines.append(line)
        if i == 0:
            markdown_lines.append('| ' + ' | '.join(['---'] * len(cells)) + ' |')

    return markdown_lines

def extract_markdown(file_path, output_dir, sheets_to_process=None, skip_first_sheet=True):
    all_sheet_names = get_sheet_names(file_path)
    if not all_sheet_names:
        return [], {}

    if sheets_to_process:
        target_sheets = sheets_to_process
    else:
        target_sheets = all_sheet_names[2:] if skip_first_sheet else all_sheet_names

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

            markdown_path = output_dir / f"{clean_sheet_name}.md"
            counter = 1
            while markdown_path.exists():
                markdown_path = output_dir / f"{clean_sheet_name}_{counter}.md"
                counter += 1

            metadata_lines = format_metadata(sheet)
            table_lines = extract_table(sheet)

            with open(markdown_path, "w", encoding="utf-8") as f:
                f.write(f"## {sheet_name} - Supplier Partner Performance Matrix\n\n")
                for line in metadata_lines:
                    f.write(f"- {line}\n")
                f.write("\n")
                for line in table_lines:
                    f.write(line + "\n")

            markdown_paths.append(markdown_path)
            name_mapping[markdown_path.stem] = sheet_name

            print(f"✅ Saved: {markdown_path.name}")

        except Exception as e:
            print(f"❌ Failed to process {sheet_name}: {e}")
            continue

    wb.close()
    return markdown_paths, name_mapping
