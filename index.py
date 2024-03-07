import pandas as pd
import tabula
from PyPDF2 import PdfReader,PdfWriter
from openpyxl import Workbook

def pdf_to_excel_with_margins(pdf_path, excel_output_path):
    margins = [2, 1.3, 1.5, 1]  # in inches
    standard_margin = 1  # standard margin for the rest of the pages, in inches
    
    writer = pd.ExcelWriter(excel_output_path, engine='openpyxl')
    
    num_pages = len(PdfReader(pdf_path).pages)
    
    for page in range(1, num_pages + 1):
        if page <= len(margins):
            margin = margins[page - 1]  # Get specific margin for this page
        else:
            margin = standard_margin  # Use standard margin for other pages
        
        top_bottom_margin = margin * 72
        
        area = [top_bottom_margin, 0, 792 - top_bottom_margin, 612]
        
        tables = tabula.read_pdf(pdf_path, pages=page, area=area, multiple_tables=True)
        
        for i, table in enumerate(tables):
            table.columns = table.columns.str.replace('^Unnamed: \d+', '', regex=True)
            table.columns = [' ' if col.isdigit() else col for col in table.columns]
            
            sheet_name = f'Page_{page}_Table_{i+1}'
            table.to_excel(writer, sheet_name=sheet_name, index=False)
    
    writer._save()


def extract_pages(input_pdf_path, output_pdf_path, initial_page, end_page):
    """
    Extracts a range of pages from a PDF and saves them to a new PDF file.

    Parameters:
    - input_pdf_path: Path to the input PDF file.
    - output_pdf_path: Path where the output PDF should be saved.
    - initial_page: The first page to include in the output (1-indexed).
    - end_page: The last page to include in the output (1-indexed).
    """
    initial_page -= 1
    end_page -= 1

    reader = PdfReader(input_pdf_path)
    writer = PdfWriter()

    for i in range(initial_page, end_page + 1):
        writer.add_page(reader.pages[i])

    with open(output_pdf_path, 'wb') as output_pdf:
        writer.write(output_pdf)

output_pdf_path = 'output.pdf'  # Replace with your desired output PDF path
input_pdf_path = input("Enter path of PDF document :")  # Replace with your input PDF path

initial_page = int(input("Enter the initial page number: "))
end_page = int(input("Enter the end page number: "))

initial_page=initial_page+1
end_page=end_page+1
excel_output_path = "output.xlsx"

extract_pages(input_pdf_path, output_pdf_path, initial_page, end_page)
pdf_to_excel_with_margins(output_pdf_path, excel_output_path)
