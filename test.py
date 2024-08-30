import pdfplumber
import tabula

# Specify the path to your PDF file
pdf_path = '1.pdf'

# Extract tables by specifying the area (in points)
# You might need to adjust these coordinates


with pdfplumber.open(pdf_path) as pdf:
    # Iterate over each page in the PDF
    for i, page in enumerate(pdf.pages):
        # Get the width and height of the page
        width = page.width
        height = page.height
        print(f"Page {i+1}: Width = {width}, Height = {height}")

        # Define the full page area
        full_page_area = (0, 0, height, width)
        print(f"Full Page Area for Page {i+1}: {full_page_area}")

        tables = tabula.read_pdf(pdf_path, pages='all', multiple_tables=True, area=(0, 0, height, width))

        # Print each table
        for i, table in enumerate(tables):
            print(f"Table {i+1}")
            print(table)
            table.to_excel(f'table_{i + 1}.xlsx')
            print("\n")
