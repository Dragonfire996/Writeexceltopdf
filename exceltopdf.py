import pandas as pd
import os
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, PageBreak, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.colors import HexColor, grey, whitesmoke, black, white

class ExcelToPDFConverter:
    def __init__(self, output_pdf):
        self.output_pdf = output_pdf
        self.elements = []
        self.bookmarks = []
        self.toc_data = []
        self.current_page = 1
        self.styles = getSampleStyleSheet()
        self.setup_styles()
        self.max_columns_per_table = 7

    def setup_styles(self):
        """Setup custom styles for the document"""
        self.title_style = ParagraphStyle(
            'CustomTitle',
            parent=self.styles['Heading1'],
            fontSize=24,
            leading=30,
            alignment=TA_CENTER,
            spaceAfter=30,
            textColor=HexColor('#1B4F72'),
            fontName='Helvetica-Bold'
        )
        
        self.heading_style = ParagraphStyle(
            'CustomHeading',
            parent=self.styles['Heading2'],
            fontSize=18,
            leading=24,
            alignment=TA_LEFT,
            spaceAfter=20,
            textColor=HexColor('#2874A6'),
            fontName='Helvetica-Bold'
        )
        
        self.toc_style = ParagraphStyle(
            'TOCStyle',
            parent=self.styles['Normal'],
            fontSize=12,
            leading=20,
            alignment=TA_LEFT,
            textColor=HexColor('#2E4053'),
            fontName='Helvetica'
        )

    def create_table_of_contents(self):
        """Create table of contents with working clickable links"""
        toc_title = Paragraph("Table of Contents", self.title_style)
        
        toc_data = [[Paragraph("Content", self.heading_style), 
                     Paragraph("Page", self.heading_style)]]
        
        for item in self.toc_data:
            link_text = (f'<a href="#{item["bookmark"]}" color="#2E4053">{item["title"]}</a>')
            toc_data.append([
                Paragraph(link_text, self.toc_style),
                Paragraph(str(item['page']), self.toc_style)
            ])
        
        toc_table = Table(toc_data, colWidths=[450, 50])
        toc_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
            ('TEXTCOLOR', (0, 0), (-1, 0), HexColor('#2874A6')),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 14),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('TEXTCOLOR', (0, 1), (-1, -1), HexColor('#2E4053')),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 12),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [whitesmoke, white]),
            ('GRID', (0, 0), (-1, 0), 1, HexColor('#2874A6')),
            ('LINEBELOW', (0, 1), (-1, -1), 0.5, grey),
            ('TOPPADDING', (0, 1), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 8),
        ]))
        
        self.elements.insert(0, toc_title)
        self.elements.insert(1, Spacer(1, 30))
        self.elements.insert(2, toc_table)
        self.elements.insert(3, PageBreak())

    def add_bookmark(self, canvas, title, bookmark_name):
        """Add a PDF bookmark that can be linked to"""
        key = f'#{bookmark_name}'
        canvas.bookmarkPage(key)
        canvas.addOutlineEntry(title, key, 0, False)

    def split_dataframe(self, df):
        """Split large DataFrames into multiple smaller ones"""
        if len(df.columns) > self.max_columns_per_table:
            splits = []
            for i in range(0, len(df.columns), self.max_columns_per_table):
                split_df = df.iloc[:, i:i + self.max_columns_per_table]
                splits.append(split_df)
            return splits
        return [df]

    def process_dataframe(self, df, max_width=None):
        """Process DataFrame to fit in PDF width"""
        if max_width is None:
            max_width = landscape(letter)[0] - 60
        
        data = [df.columns.tolist()] + df.values.tolist()
        
        min_col_width = 50
        max_col_width = 200
        
        col_widths = []
        for col in range(len(data[0])):
            max_col_width_content = max(len(str(row[col])) * 7 for row in data)
            col_width = max(min_col_width, min(max_col_width, max_col_width_content))
            col_widths.append(col_width)
        
        total_width = sum(col_widths)
        if total_width > max_width:
            scale_factor = max_width / total_width
            col_widths = [width * scale_factor for width in col_widths]
        
        return data, col_widths

    def process_excel_files(self, excel_files):
        """Process multiple Excel files into a single PDF with bookmarks"""
        doc = SimpleDocTemplate(
            self.output_pdf,
            pagesize=landscape(letter),
            rightMargin=30,
            leftMargin=30,
            topMargin=30,
            bottomMargin=30
        )
        
        for excel_file in excel_files:
            file_name = os.path.basename(excel_file)
            file_bookmark = f"file_{len(self.toc_data)}"
            
            self.toc_data.append({
                'title': file_name,
                'page': self.current_page,
                'bookmark': file_bookmark
            })
            
            title = Paragraph(f'<a name="{file_bookmark}"></a>{file_name}', self.title_style)
            self.elements.append(title)
            
            try:
                excel = pd.ExcelFile(excel_file)
                
                for sheet_name in excel.sheet_names:
                    try:
                        df = pd.read_excel(excel_file, sheet_name=sheet_name)
                        sheet_bookmark = f"sheet_{len(self.toc_data)}"
                        
                        self.toc_data.append({
                            'title': f"    ↳ {sheet_name}",
                            'page': self.current_page,
                            'bookmark': sheet_bookmark
                        })
                        
                        split_dfs = self.split_dataframe(df)
                        
                        sheet_header = Paragraph(
                            f'<a name="{sheet_bookmark}"></a>Sheet: {sheet_name}',
                            self.heading_style
                        )
                        self.elements.append(sheet_header)
                        
                        for i, split_df in enumerate(split_dfs):
                            if i > 0:
                                self.elements.append(
                                    Paragraph(f"Continued ({i+1}/{len(split_dfs)})",
                                    self.heading_style)
                                )
                            
                            data, col_widths = self.process_dataframe(split_df)
                            table = Table(data, colWidths=col_widths)
                            
                            table.setStyle(TableStyle([
                                ('BACKGROUND', (0, 0), (-1, 0), HexColor('#2874A6')),
                                ('TEXTCOLOR', (0, 0), (-1, 0), whitesmoke),
                                ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                                ('FONTSIZE', (0, 0), (-1, 0), 12),
                                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                                ('TOPPADDING', (0, 0), (-1, 0), 12),
                                ('BACKGROUND', (0, 1), (-1, -1), white),
                                ('TEXTCOLOR', (0, 1), (-1, -1), black),
                                ('ALIGN', (0, 1), (-1, -1), 'LEFT'),
                                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                                ('FONTSIZE', (0, 1), (-1, -1), 10),
                                ('TOPPADDING', (0, 1), (-1, -1), 6),
                                ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
                                ('GRID', (0, 0), (-1, -1), 0.5, grey),
                                ('ROWBACKGROUNDS', (0, 1), (-1, -1), [white, whitesmoke]),
                            ]))
                            
                            self.elements.append(table)
                            self.elements.append(Spacer(1, 20))
                        
                        self.elements.append(PageBreak())
                        self.current_page += 1
                        
                    except Exception as e:
                        print(f"Error processing sheet {sheet_name} in {file_name}: {str(e)}")
                        continue
                        
            except Exception as e:
                print(f"Error processing file {file_name}: {str(e)}")
                continue
        
        self.create_table_of_contents()
        
        doc.build(
            self.elements,
            onFirstPage=self._header_footer,
            onLaterPages=self._header_footer,
        )

    def _header_footer(self, canvas, doc):
        """Add page numbers and maintain bookmarks in footer"""
        canvas.saveState()
        
        canvas.setFont('Helvetica', 9)
        page_num = canvas.getPageNumber()
        text = f"Page {page_num}"
        canvas.drawRightString(landscape(letter)[0] - 30, 20, text)
        
        current_bookmark = next(
            (item for item in self.toc_data if item['page'] == page_num),
            None
        )
        if current_bookmark:
            self.add_bookmark(
                canvas,
                current_bookmark['title'],
                current_bookmark['bookmark']
            )
        
        canvas.restoreState()

def batch_excel_to_pdf(input_directory, output_pdf):
    """Convert all Excel files in a directory to a single PDF"""
    excel_files = [
        os.path.join(input_directory, f) 
        for f in os.listdir(input_directory) 
        if f.endswith(('.xlsx', '.xls', '.xlsm'))
    ]
    
    if not excel_files:
        print(f"No Excel files found in {input_directory}")
        return
    
    print(f"Found {len(excel_files)} Excel files")
    converter = ExcelToPDFConverter(output_pdf)
    converter.process_excel_files(excel_files)
    print(f"PDF created successfully: {output_pdf}")

if __name__ == "__main__":
    # Example usage:
    # Option 1: Process specific Excel files
    excel_files = [
        "test1.xlsx",
        "test.xlsx",
        "updated.xlsx"
    ]
    converter = ExcelToPDFConverter("output.pdf")
    converter.process_excel_files(excel_files)
    
    # Option 2: Process all Excel files in a directory
    # batch_excel_to_pdf("path/to/excel/files", "output.pdf")