import os
import sys
import openpyxl
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfgen import canvas
import xml.sax.saxutils as saxutils

class LandscapeNumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_footer(num_pages)
            super().showPage()
        super().save()

    def draw_footer(self, page_count):
        self.saveState()
        self.setFont("Helvetica", 7.5)
        self.setFillColor(colors.HexColor("#718096"))
        
        # Header (pages > 1)
        if self._pageNumber > 1:
            self.drawString(36, 565, getattr(self, "_doc_title", "P052: Store Wise Sales Comparison Dashboard"))
            self.drawRightString(806, 565, "Anu Tyagi | Roll: 2500320140025 | Mentor: Mr. Dinesh")
            self.setStrokeColor(colors.HexColor("#CBD5E0"))
            self.setLineWidth(0.5)
            self.line(36, 560, 806, 560)
            
        # Footer
        page_str = f"Page {self._pageNumber} of {page_count}"
        self.drawString(36, 20, "ABES Engineering College, Ghaziabad | MCA 2nd Year (2026-2027) | Project ID: P052")
        self.drawRightString(806, 20, page_str)
        self.setStrokeColor(colors.HexColor("#CBD5E0"))
        self.setLineWidth(0.5)
        self.line(36, 28, 806, 28)
        self.restoreState()

def convert_rtm_to_pdf(xlsx_path, pdf_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active
    
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=landscape(A4),
        leftMargin=30,
        rightMargin=30,
        topMargin=36,
        bottomMargin=36
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("RTM_Title", fontName="Helvetica-Bold", fontSize=15, leading=18, textColor=colors.HexColor("#1A365D"), spaceAfter=3)
    meta_style = ParagraphStyle("RTM_Meta", fontName="Helvetica-Bold", fontSize=8.5, leading=11, textColor=colors.HexColor("#2B6CB0"), spaceAfter=10)
    th_style = ParagraphStyle("TH_Style", fontName="Helvetica-Bold", fontSize=7, leading=9, textColor=colors.white, alignment=1)
    td_style = ParagraphStyle("TD_Style", fontName="Helvetica", fontSize=6.5, leading=8.5, textColor=colors.HexColor("#2D3748"))
    td_bold = ParagraphStyle("TDB_Style", fontName="Helvetica-Bold", fontSize=6.5, leading=8.5, textColor=colors.HexColor("#1A365D"))

    story = []
    story.append(Paragraph(ws["A1"].value or "REQUIREMENTS TRACEABILITY MATRIX (RTM)", title_style))
    story.append(Paragraph(ws["A2"].value or "", meta_style))

    # Read rows starting from row 3 (headers)
    rows_data = []
    for r_idx in range(3, ws.max_row + 1):
        row_vals = [ws.cell(r_idx, c).value for c in range(1, 13)]
        if not any(row_vals):
            continue
        row_cells = []
        for c_idx, val in enumerate(row_vals):
            v_str = saxutils.escape(str(val if val is not None else ""))
            if r_idx == 3:
                row_cells.append(Paragraph(v_str, th_style))
            elif c_idx in [0, 1, 2, 5, 6]:
                row_cells.append(Paragraph(v_str, td_bold))
            else:
                row_cells.append(Paragraph(v_str, td_style))
        rows_data.append(row_cells)

    # 12 columns total
    col_widths = [45, 45, 45, 145, 60, 40, 50, 75, 75, 55, 55, 92]
    table = Table(rows_data, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1A365D")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E0")),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.HexColor("#FFFFFF"), colors.HexColor("#F7FAFC")]),
        ('PADDING', (0,0), (-1,-1), 2.5),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (0,0), (2,-1), 'CENTER'),
        ('ALIGN', (5,0), (6,-1), 'CENTER'),
    ]))
    story.append(table)

    canvas_class = LandscapeNumberedCanvas
    canvas_class._doc_title = "Step 04: Requirements Traceability Matrix (RTM)"
    doc.build(story, canvasmaker=canvas_class)
    print(f"[CONVERTED] {xlsx_path} -> {pdf_path}")

def convert_db_design_to_pdf(xlsx_path, pdf_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=landscape(A4),
        leftMargin=30,
        rightMargin=30,
        topMargin=36,
        bottomMargin=36
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("DB_Title", fontName="Helvetica-Bold", fontSize=15, leading=18, textColor=colors.HexColor("#1A365D"), spaceAfter=3)
    meta_style = ParagraphStyle("DB_Meta", fontName="Helvetica-Bold", fontSize=8.5, leading=11, textColor=colors.HexColor("#2B6CB0"), spaceAfter=10)
    sheet_title_style = ParagraphStyle("Sheet_Title", fontName="Helvetica-Bold", fontSize=11, leading=14, textColor=colors.HexColor("#1A365D"), spaceBefore=6, spaceAfter=4)
    th_style = ParagraphStyle("TH_Style", fontName="Helvetica-Bold", fontSize=7, leading=9, textColor=colors.white, alignment=1)
    td_style = ParagraphStyle("TD_Style", fontName="Helvetica", fontSize=6.5, leading=8.5, textColor=colors.HexColor("#2D3748"))
    td_bold = ParagraphStyle("TDB_Style", fontName="Helvetica-Bold", fontSize=6.5, leading=8.5, textColor=colors.HexColor("#1A365D"))

    story = []
    story.append(Paragraph("STORE WISE SALES COMPARISON DASHBOARD — DATABASE DESIGN SPECIFICATION", title_style))
    story.append(Paragraph("Project ID: P052 | Candidate: Anu Tyagi (Roll: 2500320140025) | Batch: ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 | Mentor: Mr. Dinesh", meta_style))

    for s_idx, s_name in enumerate(wb.sheetnames):
        ws = wb[s_name]
        story.append(Paragraph(f"<b>Sheet {s_idx+1}: {s_name}</b>", sheet_title_style))
        
        rows_data = []
        max_col = min(ws.max_column, 8)
        for r_idx in range(1, min(ws.max_row + 1, 40)):
            row_vals = [ws.cell(r_idx, c).value for c in range(1, max_col + 1)]
            if not any(row_vals):
                continue
            row_cells = []
            for c_idx, val in enumerate(row_vals):
                v_str = saxutils.escape(str(val if val is not None else ""))
                if r_idx == 1:
                    row_cells.append(Paragraph(v_str, th_style))
                elif c_idx == 0:
                    row_cells.append(Paragraph(v_str, td_bold))
                else:
                    row_cells.append(Paragraph(v_str, td_style))
            rows_data.append(row_cells)

        if rows_data:
            c_w = 780 / len(rows_data[0])
            table = Table(rows_data, colWidths=[c_w]*len(rows_data[0]), repeatRows=1)
            table.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1A365D")),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E0")),
                ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.HexColor("#FFFFFF"), colors.HexColor("#F7FAFC")]),
                ('PADDING', (0,0), (-1,-1), 2.5),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ]))
            story.append(table)
            story.append(Spacer(1, 8))

    canvas_class = LandscapeNumberedCanvas
    canvas_class._doc_title = "Step 08: Database Design Specification"
    doc.build(story, canvasmaker=canvas_class)
    print(f"[CONVERTED] {xlsx_path} -> {pdf_path}")

if __name__ == "__main__":
    convert_rtm_to_pdf("docs/04_Traceability_Matrix.xlsx", "docs/04_Traceability_Matrix.pdf")
    convert_db_design_to_pdf("docs/08_Database_Design.xlsx", "docs/08_Database_Design.pdf")
