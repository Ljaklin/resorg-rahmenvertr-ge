import fitz  # pymupdf
import os

# Directory containing the PDF templates
template_dir = "data/00.Musterordner Rahmenverträge"

# Get all PDF files
pdf_files = [f for f in os.listdir(template_dir) if f.lower().endswith('.pdf')]

for pdf_file in pdf_files:
    pdf_path = os.path.join(template_dir, pdf_file)
    print(f"\n{'='*80}")
    print(f"PDF: {pdf_file}")
    print(f"{'='*80}")
    
    try:
        doc = fitz.open(pdf_path)
        field_count = 0
        
        for page_num, page in enumerate(doc, 1):
            widgets = page.widgets()
            if widgets:
                print(f"\n--- Page {page_num} ---")
                for field in widgets:
                    field_count += 1
                    print(f"  Field Name: {field.field_name}")
                    print(f"    Field Type: {field.field_type}")
                    print(f"    Field Value: {field.field_value}")
                    print()
        
        if field_count == 0:
            print("  No form fields found in this PDF")
        else:
            print(f"\nTotal fields: {field_count}")
        
        doc.close()
    except Exception as e:
        print(f"  Error reading PDF: {e}")
