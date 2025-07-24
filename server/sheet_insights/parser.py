from sheet_insights.config import parser
import re
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor
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

def extract_markdown(file_path, output_dir, sheets_to_process=None, skip_first_sheet=True):
    """
    Extract markdown from Excel sheets with proper name mapping

    Args:
        file_path: Path to Excel file
        output_dir: Directory to save markdown files
        sheets_to_process: List of specific sheet names to process (optional)
        skip_first_sheet: Whether to skip the first sheet (default: True)
    """
    from sheet_insights.config import parser

    # Get actual sheet names from the Excel file
    all_sheet_names = get_sheet_names(file_path)
    if not all_sheet_names:
        print("❌ No sheets found in the Excel file")
        return [], {}

    # Determine which sheets to process
    if sheets_to_process is not None:
        # Use the explicitly provided sheets
        target_sheets = sheets_to_process
        print(f"📋 Processing specified sheets: {target_sheets}")
        # Calculate how many sheets to skip from the beginning
        sheets_to_skip = len(all_sheet_names) - len(target_sheets)
        if target_sheets == all_sheet_names[2:]:
            sheets_to_skip = 2  # Skipping first two sheets
        elif target_sheets == all_sheet_names[1:]:
            sheets_to_skip = 1  # Skipping first sheet
        else:
            sheets_to_skip = 0  # Custom selection
    else:
        # Auto-determine based on skip_first_sheet parameter
        if skip_first_sheet and len(all_sheet_names) > 2:
            # Skip first two sheets: "Average Summary" and "Analysis SUMMARY"
            target_sheets = all_sheet_names[2:]  # Skip first two sheets
            sheets_to_skip = 2
            print(f"📋 Skipping first two sheets '{all_sheet_names[0]}' and '{all_sheet_names[1]}', processing: {target_sheets}")
        elif skip_first_sheet and len(all_sheet_names) > 1:
            target_sheets = all_sheet_names[1:]  # Skip first sheet only
            sheets_to_skip = 1
            print(f"📋 Skipping first sheet '{all_sheet_names[0]}', processing: {target_sheets}")
        else:
            target_sheets = all_sheet_names
            sheets_to_skip = 0
            print(f"📋 Processing all sheets: {target_sheets}")

    # Load documents using parser
    print(f"🔄 Loading documents from: {file_path}")
    documents = parser.load_data(file_path)
    print(f"📊 Parser returned {len(documents)} documents")

    # IMPORTANT: LlamaParse processes ALL sheets, so we need to align documents with target sheets
    # Skip the appropriate number of documents from the beginning
    if sheets_to_skip > 0 and len(documents) == len(all_sheet_names):
        print(f"🔧 Skipping first {sheets_to_skip} document(s) to align with target sheets")
        documents = documents[sheets_to_skip:]
    elif len(documents) != len(all_sheet_names):
        print(f"⚠️ WARNING: Document count ({len(documents)}) doesn't match all sheets count ({len(all_sheet_names)})")

    # Validate document count matches expected sheets
    if len(documents) != len(target_sheets):
        print(f"⚠️ WARNING: Document count ({len(documents)}) doesn't match target sheets count ({len(target_sheets)})")
        print(f"Documents: {len(documents)}, Target sheets: {target_sheets}")

        # Try to match by taking the first N documents where N = len(target_sheets)
        if len(documents) > len(target_sheets):
            print(f"🔧 Truncating documents to match sheet count")
            documents = documents[:len(target_sheets)]
        elif len(documents) < len(target_sheets):
            print(f"🔧 Adjusting target sheets to match document count")
            target_sheets = target_sheets[:len(documents)]

    markdown_paths = []
    name_mapping = {}

    # Process each document with corresponding sheet name
    for i, doc in enumerate(documents):
        try:
            if i < len(target_sheets):
                # Use actual sheet name
                original_sheet_name = target_sheets[i]
                # Clean the name for filename (replace spaces, special chars)
                clean_sheet_name = re.sub(r'[^\w\-_]', '_', original_sheet_name.strip())
                clean_sheet_name = re.sub(r'_+', '_', clean_sheet_name)  # Remove multiple underscores
                clean_sheet_name = clean_sheet_name.strip('_')  # Remove leading/trailing underscores
                
                print(f"📄 Processing sheet {i+1}/{len(documents)}: '{original_sheet_name}' -> '{clean_sheet_name}'")
            else:
                # Fallback for unexpected documents
                original_sheet_name = f"Unknown_Sheet_{i+1}"
                clean_sheet_name = f"Unknown_Sheet_{i+1}"
                print(f"⚠️ No sheet name for document {i}, using fallback: {clean_sheet_name}")

            # Create markdown file
            markdown_path = output_dir / f"{clean_sheet_name}.md"
            
            # Handle duplicate filenames
            counter = 1
            original_path = markdown_path
            while markdown_path.exists():
                stem = original_path.stem
                markdown_path = output_dir / f"{stem}_{counter}.md"
                counter += 1

            # Write markdown content
            with open(markdown_path, "w", encoding="utf-8") as f:
                f.write(doc.text)

            # Update document metadata
            doc.metadata["source"] = original_sheet_name
            doc.metadata["clean_name"] = clean_sheet_name
            
            markdown_paths.append(markdown_path)

            # Map filename stem to original sheet name for later reference
            name_mapping[markdown_path.stem] = original_sheet_name

            print(f"✅ Saved: {markdown_path.name} for sheet: '{original_sheet_name}'")

        except Exception as e:
            print(f"❌ Error processing document {i}: {e}")
            continue

    print(f"🎉 Successfully processed {len(markdown_paths)} sheets")
    print(f"📝 Name mapping: {name_mapping}")
    
    return markdown_paths, name_mapping