import os
import json
import re
import pdfplumber
from datetime import datetime

def extract_text_from_pdf(pdf_path):
    """Extracts text from all pages of a PDF using pdfplumber."""
    text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
    except Exception as e:
        print(f"Error reading PDF {pdf_path}: {e}")
    return text

def extract_text_from_txt(txt_path):
    """Reads text from a TXT file."""
    try:
        with open(txt_path, 'r', encoding='utf-8') as f:
            return f.read()
    except Exception as e:
        print(f"Error reading TXT {txt_path}: {e}")
        return ""

def parse_date(date_str):
    """Attempts to parse and format date to DD-MM-YYYY."""
    formats = ["%d-%m-%Y", "%d/%m/%Y", "%d.%m.%Y", "%Y-%m-%d", "%d %b %Y", "%d %B %Y"]
    for fmt in formats:
        try:
            return datetime.strptime(date_str.strip(), fmt).strftime("%d-%m-%Y")
        except ValueError:
            continue
    return date_str.strip() # Return as is if parsing fails

def extract_po_data(text):
    """
    Extracts header information and line items from the PO text using Regex.
    """
    
    po_number = None
    po_date = None
    customer_name = None
    gst_number = None
    
    # 1. Extract PO Number
    po_number_match = re.search(r'(?:PO\s*No|Purchase\s*Order\s*No|Order\s*No|PO\s*Number)[\s\.:\-]*([A-Z0-9\-\/]+)', text, re.IGNORECASE)
    if po_number_match: po_number = po_number_match.group(1).strip()

    # 2. Extract PO Date
    po_date_match = re.search(r'(?:Date|Dated|PO\s*Date)[\s\.:\-]*(\d{1,2}[\-\/\.]\d{1,2}[\-\/\.]\d{2,4}|\d{1,2}\s+[A-Za-z]+\s+\d{4})', text, re.IGNORECASE)
    if po_date_match: po_date = parse_date(po_date_match.group(1))

    # 3. Extract Customer Name
    customer_match = re.search(r'(?:Bill\s*To|Invoice\s*To|Ship\s*To|Delivery\s*To|Buyer)[\s\.:\-]*\n?([A-Za-z0-9\s\.\,\&\-]+(?:\s+[A-Za-z0-9\s\.\,\&\-]+)*)', text, re.IGNORECASE)
    if customer_match:
        lines = [line.strip() for line in customer_match.group(1).split('\n') if line.strip()]
        if lines: customer_name = lines[0]

    # 4. Extract GST Number
    gst_match = re.search(r'(?:GSTIN|GST\s*No|GST\s*Number)[\s\.:\-]*([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})', text, re.IGNORECASE)
    if gst_match: gst_number = gst_match.group(1).strip()

    line_items = []
    
    # Find the section where line items typically start
    lines = text.split('\n')
    item_section_start_index = -1
    for i, line in enumerate(lines):
        # Look for a header line to indicate the start of items
        if re.search(r'Description\s+of\s+Material\s+Quantity\s+Amount', line, re.IGNORECASE):
            item_section_start_index = i
            # Skip the separator line if it exists right after the header
            if i + 1 < len(lines) and re.search(r'^-+$', lines[i+1].strip()):
                item_section_start_index = i + 1
            break

    if item_section_start_index != -1:
        # Process lines after the header/separator
        for line in lines[item_section_start_index + 1:]:
            # Refined regex for individual line items
            # It looks for a description (non-greedy), followed by quantity and amount at the end of the line.
            # It tries to avoid matching lines that are clearly not items (e.g., empty lines, totals).
            # The description part is made more specific to avoid capturing the header.
            item_pattern = re.search(r'^(?P<desc>.+?)\s+(?P<qty>\d+(?:\.\d+)?)\s+(?P<amount>\d+(?:\.\d+)?)$', line.strip())
            
            if item_pattern:
                desc = item_pattern.group('desc').strip()
                # Further filter out lines that might be mistaken for items (e.g., summary lines)
                if desc.lower() in ['total', 'subtotal', 'tax', 'cgst', 'sgst', 'igst', 'freight', 'discount', 'grand total']:
                    continue
                
                try:
                    qty = float(item_pattern.group('qty'))
                    amount = float(item_pattern.group('amount'))
                    
                    line_items.append({
                        "PO_number": po_number,
                        "po_date": po_date,
                        "Customer_name": customer_name,
                        "gst_number": gst_number,
                        "material_code": None,
                        "Description_of_material": desc,
                        "quantity": qty,
                        "amount": amount
                    })
                except ValueError:
                    continue

    return line_items

def process_folder(folder_path):
    """Processes all PDF and TXT files in the given folder."""
    all_extracted_items = []
    
    if not os.path.exists(folder_path):
        print(f"Folder not found: {folder_path}")
        return []

    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        text = ""
        
        if filename.lower().endswith('.pdf'):
            print(f"Processing PDF: {filename}")
            text = extract_text_from_pdf(file_path)
        elif filename.lower().endswith('.txt'):
            print(f"Processing TXT: {filename}")
            text = extract_text_from_txt(file_path)
        else:
            continue
            
        if text:
            items = extract_po_data(text)
            all_extracted_items.extend(items)
            
    return all_extracted_items

if __name__ == "__main__":
    TARGET_FOLDER = "./po_files" 
    
    if not os.path.exists(TARGET_FOLDER):
        os.makedirs(TARGET_FOLDER)
        print(f"Created folder '{TARGET_FOLDER}'. Please place your PDF/TXT files there.")
    
    print(f"Starting extraction from folder: {TARGET_FOLDER}")
    extracted_data = process_folder(TARGET_FOLDER)
    
    json_output = json.dumps(extracted_data, indent=2)
    
    output_file = "extracted_po_data.json"
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(json_output)
        
    print(f"\nExtraction complete. Found {len(extracted_data)} line items.")
    print(f"Results saved to {output_file}")
    
    print("\n--- JSON OUTPUT ---")
    print(json_output)
