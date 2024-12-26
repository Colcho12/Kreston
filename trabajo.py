import re
from collections import defaultdict
import pandas as pd

def export(documents):
    # Flatten data for structured export
    data = []
    for doc, details in documents.items():
        # Add Document Row (First row with metadata)
        data.append({
            'Document': doc,
            'Posting Date': details['Posting Date'],
            'Reference': details.get('Reference', 'N/A'),
            'Type': details.get('Type', 'N/A'),
            'Item': '',
            'Item Amount': '',
            'Account': '',
            'Amount Local': '',
            'Amount Document': '',
            'D/C': ''
        })
        
        # Add Item Rows Below the Header
        for item in details['Items']:
            data.append({
                'Document': '',
                'Posting Date': '',
                'Reference': '',
                'Type': '',
                'Item': item['Item'],
                'Item Amount': item['Item Amount'],
                'Account': item['Account'],
                'Amount Local': item['Amount Local'],
                'Amount Document': item['Amount Document'],
                'D/C': item['D/C']
            })

    # Export to Excel
    output_file = 'structured_documents_with_two_headers2.xlsx'
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet('Documents')
        
        # Define Formats
        header_format = workbook.add_format({'bold': True, 'bg_color': '#D9EAD3', 'align': 'center', 'border': 1})
        subheader_format = workbook.add_format({'bold': True, 'bg_color': '#C9DAF8', 'align': 'center', 'border': 1})
        document_format = workbook.add_format({'bold': True, 'bg_color': '#FFF2CC', 'align': 'left', 'border': 1})
        item_format = workbook.add_format({'align': 'left', 'border': 1})
        
        # Write Level 1 Header (First Row)
        level1_headers = ['Document', 'Posting Date', 'Reference', 'Type', '', '', '', '', '', '']
        worksheet.write_row(0, 0, level1_headers, header_format)
        
        # Write Level 2 Header (Second Row)
        level2_headers = ['Item', 'Item Amount', 'Account', 'Amount Local', 'Amount Document', 'D/C']
        worksheet.write_row(1, 0, level2_headers, subheader_format)
        
        # Write Data
        row = 2  # Start writing data from row 3 (0-indexed)
        for record in data:
            if record['Document']:  # Document-level row
                worksheet.write(row, 0, record['Document'], document_format)
                worksheet.write(row, 1, record['Posting Date'], document_format)
                worksheet.write(row, 2, record['Reference'], document_format)
                worksheet.write(row, 3, record['Type'], document_format)
            else:  # Item-level row
                worksheet.write(row, 0, record['Item'], item_format)
                worksheet.write(row, 1, record['Item Amount'], item_format)
                worksheet.write(row, 2, record['Account'], item_format)
                worksheet.write(row, 3, record['Amount Local'], item_format)
                worksheet.write(row, 4, record['Amount Document'], item_format)
                worksheet.write(row, 5, record['D/C'], item_format)
            
            row += 1  # Move to the next row
        
        # Adjust Column Widths
        worksheet.set_column('A:A', 20)  # Document
        worksheet.set_column('B:B', 15)  # Posting Date
        worksheet.set_column('C:C', 15)  # Reference
        worksheet.set_column('D:D', 10)  # Type
        worksheet.set_column('E:E', 10)  # Item
        worksheet.set_column('F:F', 15)  # Item Amount
        worksheet.set_column('G:G', 15)  # Account
        worksheet.set_column('H:H', 15)  # Amount Local
        worksheet.set_column('I:I', 15)  # Amount Document
        worksheet.set_column('J:J', 5)   # D/C
        
        # Freeze Header Rows
        worksheet.freeze_panes(2, 0)  # Freeze top two rows for headers

    print(f"Data successfully exported to {output_file}")

file_path = 'Diarios 2018-2.txt'
import re
from collections import defaultdict

import re
from collections import defaultdict

def txt_manipulate(file_path):
    # Initialize storage
    documents = defaultdict(dict)

    # Regular expressions for matching
    doc_pattern = re.compile(r'^\s*\d+\s+(\d{9})')  # Match Document Number
    date_pattern = re.compile(r'(\d{2}\.\d{2}\.\d{4})')  # Match Posting Date / Entry Date (DD.MM.YYYY)
    reference_pattern = re.compile(r'(\d{10})')  # Match Reference Number
    type_pattern = re.compile(r'\b([A-Z]{2})\b')  # Match Type (e.g., RV)
    item_pattern = re.compile(r'^\s*(\d+)\s+(\d+)\s+[A-Z0-9]+\s+')  # Match Item, Item Amount
    account_pattern = re.compile(r'(\d{7,8})')  # Match Account Number
    dc_pattern = re.compile(r'\b[HS]\b')  # Match D/C Indicator
    amount_pattern = re.compile(r'(-?[\d,]+\.\d{2})')  # Match monetary amounts

    # Read file
    with open(file_path, 'r', encoding='ISO-8859-1') as file:
        current_doc = None
        current_date = None
        current_entry_date = None
        current_reference = None
        current_type = None
        amounts_buffer = []
        ready_for_items = False
        is_first_item = True
        entry_date_found = False  # Flag to track Entry Date detection
        skip_next_line = False  # New flag to skip the next line after "Entry Date"

        for line in file:
            line = line.strip()
            
            # Inside the main parsing loop
            if 'Entry Date' in line:
                print('yes')
                entry_date_found = True
                skip_next_line = True  # Skip the next line explicitly
                continue

            # Skip the "Crcy" line
            if skip_next_line:
                skip_next_line = False  # Reset flag and proceed to the next line
                continue

            # Capture the Entry Date on the next valid line
            if entry_date_found:
                date_match = date_pattern.search(line)
                if date_match:
                    current_entry_date = date_match.group(1)
                    entry_date_found = False  # Reset flag after capturing Entry Date
            # Check for Posting Date
            date_match = date_pattern.search(line)
            if date_match and not ready_for_items:
                current_date = date_match.group(1)
            
            # Check for Reference Number
            reference_match = reference_pattern.search(line)
            if reference_match and current_doc is None:
                current_reference = reference_match.group(1)
            
            # Check for Document Type
            type_match = type_pattern.search(line)
            if type_match and current_doc is None:
                current_type = type_match.group(1)
            
            # Check for new document number
            doc_match = doc_pattern.match(line)
            if doc_match:
                current_doc = doc_match.group(1)
                documents[current_doc] = {
                    'Posting Date': current_date,
                    'Entry Date': current_entry_date,
                    'Reference': current_reference,
                    'Type': current_type,
                    'Items': []
                }
                amounts_buffer.clear()
                current_reference = None
                current_type = None
                current_entry_date = None
                ready_for_items = True
                is_first_item = True
                continue
            
            # If inside a document, match item-level data
            if current_doc and ready_for_items:
                item_match = item_pattern.match(line)
                if item_match:
                    item = item_match.group(1)
                    item_amount = item_match.group(2)
                    
                    account_match = account_pattern.search(line)
                    account = account_match.group(1) if account_match else None
                    
                    dc = None
                    if not is_first_item:
                        dc_match = dc_pattern.search(line)
                        dc = dc_match.group(0) if dc_match else None
                    
                    local_amount = None
                    document_amount = None
                    
                    if len(amounts_buffer) == 0:
                        amounts = amount_pattern.findall(line)
                        amounts_buffer.extend(amounts)
                        if len(amounts_buffer) >= 2:
                            local_amount = float(amounts_buffer.pop(0).replace(',', ''))
                            document_amount = float(amounts_buffer.pop(0).replace(',', ''))
                    
                    if len(amounts_buffer) >= 2:
                        local_amount = float(amounts_buffer.pop(0).replace(',', ''))
                        document_amount = float(amounts_buffer.pop(0).replace(',', ''))
                    
                    documents[current_doc]['Items'].append({
                        'Item': item,
                        'Item Amount': int(item_amount),
                        'Account': account,
                        'D/C': dc,
                        'Amount Local': local_amount,
                        'Amount Document': document_amount
                    })
                    
                    is_first_item = False
            
            # Reset context at a blank line
            if line.strip() == '':
                current_doc = None
                current_date = None
                current_entry_date = None
                current_reference = None
                current_type = None
                ready_for_items = False
                amounts_buffer.clear()
                is_first_item = True
    
    return documents

documents = txt_manipulate(file_path)
# Display parsed data with Entry Date
for doc, details in list(documents.items())[:4]:  # Show first 2 documents for clarity
    print(f"Document Number: {doc}")
    print(f"  Posting Date: {details['Posting Date']}")
    print(f"  Entry Date: {details['Entry Date']}")
    print(f"  Reference: {details['Reference']}")
    print(f"  Type: {details['Type']}")
    print("  Items:")
    for item in details['Items']:
        print(f"    Item: {item['Item']}, "
              f"Item Amount: {item['Item Amount']}, "
              f"Account: {item['Account']}, "
              f"D/C: {item['D/C']}, "
              f"Amount Local: {item['Amount Local']}, "
              f"Amount Document: {item['Amount Document']}")
    print("-" * 50)

