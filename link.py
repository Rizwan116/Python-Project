import fitz  # PyMuPDF
import pandas as pd

def add_links_to_highlighted_sections(pdf_path, excel_path, output_pdf_path):
    # Open the PDF file
    pdf_document = fitz.open(pdf_path)
    
    # Load the Excel file
    excel_data = pd.read_excel(excel_path)

    # Ensure that the Excel has the expected columns: "Page" and "URL"
    if not all(column in excel_data.columns for column in ["Page", "URL"]):
        print("Error: Excel file must contain 'Page' and 'URL' columns.")
        return

    # Loop through each row in the Excel file to add links to highlighted sections
    for index, row in excel_data.iterrows():
        page_number = int(row['Page']) - 1  # Page numbers are 0-indexed in PyMuPDF
        url = row['URL']

        # Load the page
        page = pdf_document.load_page(page_number)

        # Retrieve all annotations on the page
        annotations = page.annots()
        
        if annotations:
            print(f"Annotations found on page {page_number + 1}:")
            for annot in annotations:
                print(f"Annotation: {annot.info}, Rect: {annot.rect}")
                
                # Add a link over the highlighted section
                page.insert_link({
                    "kind": fitz.LINK_URI,
                    "from": annot.rect,
                    "uri": url
                })
        else:
            print(f"No annotations found on page {page_number + 1}.")

    # Save the output PDF
    pdf_document.save(output_pdf_path)
    pdf_document.close()
    print(f"Links have been successfully added to highlighted sections in {output_pdf_path}")

# Usage
pdf_path = 'lorem.pdf'  # Path to the input PDF file
excel_path = 'ex.xlsx'  # Path to the Excel file containing link details
output_pdf_path = 'lorem2.pdf'  # Path to save the output PDF

add_links_to_highlighted_sections(pdf_path, excel_path, output_pdf_path)
