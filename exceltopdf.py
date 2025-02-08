import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from PyPDF2 import PdfMerger, PdfReader
import os


def create_test_case_pdf(excel_path, output_pdf):
    """Generate a PDF from an Excel file and display the test results in a table format."""
    data = pd.read_excel(excel_path)
    
    doc = SimpleDocTemplate(output_pdf, pagesize=letter)
    elements = []

    # Convert DataFrame to list format for table
    table_data = [data.columns.tolist()] + data.values.tolist()
    
    # Create table
    table = Table(table_data)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 12),
        ("GRID", (0, 0), (-1, -1), 1, colors.black)
    ]))

    elements.append(table)
    doc.build(elements)


def create_summary_pdf(test_files, output_pdf):
    """Generate the summary PDF with a list of test reports."""
    c = canvas.Canvas(output_pdf, pagesize=letter)
    c.setFont("Helvetica-Bold", 14)
    c.drawString(50, 800, "Test Automation Summary")
    c.setFont("Helvetica", 12)

    y = 750
    for i, test in enumerate(test_files):
        c.drawString(50, y, f"{i+1}. {test}")
        c.bookmarkPage(f"test_page_{i}")  # Add a bookmark
        y -= 20

    c.save()

def merge_pdfs(summary_pdf, test_pdfs, final_pdf):
    """Combine the summary and all test PDFs into one and add working bookmarks."""
    merger = PdfMerger()

    # Ensure summary.pdf exists before merging
    if not os.path.exists(summary_pdf):
        create_summary_pdf([os.path.basename(pdf) for pdf in test_pdfs], summary_pdf)

    # Add summary PDF first
    merger.append(summary_pdf)
    
    # Create bookmarks for each section
    page_num = len(PdfReader(summary_pdf).pages)  # Start after the summary pages

    for pdf in test_pdfs:
        # Append the test case PDF
        merger.append(pdf)

        # Add a bookmark that links to the correct page
        merger.add_outline_item(title=os.path.basename(pdf), pagenum=page_num, parent=None)

        # Update page number offset
        page_num += len(PdfReader(pdf).pages)

    # Write the final PDF with bookmarks
    with open(final_pdf, "wb") as f_out:
        merger.write(f_out)

    print(f"Final report generated with working bookmarks: {final_pdf}")



def main():
    test_pdfs = []
    excel_files = ["test.xlsx", "test1.xlsx"]  # Change to actual file paths

    # Generate individual test case PDFs
    for excel_file in excel_files:
        pdf_file = excel_file.replace(".xlsx", ".pdf")
        create_test_case_pdf(excel_file, pdf_file)
        test_pdfs.append(pdf_file)

    # Generate summary PDF
    summary_pdf = "summary.pdf"
    create_summary_pdf(excel_files, summary_pdf)

    # Merge PDFs
    final_pdf = "final_report.pdf"
    merge_pdfs(summary_pdf, test_pdfs, final_pdf)

    print(f"Final report generated: {final_pdf}")


if __name__ == "__main__":
    main()
