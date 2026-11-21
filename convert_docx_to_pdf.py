import os
import sys
import docx
from docx.oxml.ns import qn
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle, Image
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfgen import canvas
import xml.sax.saxutils as saxutils

class DocxNumberedCanvas(canvas.Canvas):
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
            self.draw_decorations(num_pages)
            super().showPage()
        super().save()

    def draw_decorations(self, page_count):
        self.saveState()
        self.setFont("Helvetica", 7.5)
        self.setFillColor(colors.HexColor("#718096"))
        
        # Header on pages > 1
        if self._pageNumber > 1:
            self.drawString(36, 806, getattr(self, "_doc_title", "Store Wise Sales Comparison Dashboard (P052)"))
            self.drawRightString(559, 806, "Anu Tyagi | Roll: 2500320140025")
            self.setStrokeColor(colors.HexColor("#CBD5E0"))
            self.setLineWidth(0.5)
            self.line(36, 800, 559, 800)
            
        # Footer
        page_str = f"Page {self._pageNumber} of {page_count}"
        self.drawString(36, 25, "ABES Engineering College, Ghaziabad | MCA 2nd Year (2026-2027) | Mentor: Mr. Dinesh")
        self.drawRightString(559, 25, page_str)
        self.setStrokeColor(colors.HexColor("#CBD5E0"))
        self.setLineWidth(0.5)
        self.line(36, 35, 559, 35)
        self.restoreState()

def docx_to_pdf(docx_path, pdf_path, custom_title=None):
    if not os.path.exists(docx_path):
        print(f"Skipping {docx_path} (not found)")
        return False
        
    doc = docx.Document(docx_path)
    os.makedirs(os.path.dirname(pdf_path), exist_ok=True)
    
    pdf_doc = SimpleDocTemplate(
        pdf_path,
        pagesize=A4,
        leftMargin=36,
        rightMargin=36,
        topMargin=46,
        bottomMargin=46
    )

    styles = getSampleStyleSheet()

    # Define custom styles
    h1_style = ParagraphStyle(
        "H1_Style",
        fontName="Helvetica-Bold",
        fontSize=13,
        leading=16,
        textColor=colors.HexColor("#1A365D"),
        spaceBefore=10,
        spaceAfter=5
    )
    h2_style = ParagraphStyle(
        "H2_Style",
        fontName="Helvetica-Bold",
        fontSize=11,
        leading=14,
        textColor=colors.HexColor("#2B6CB0"),
        spaceBefore=8,
        spaceAfter=4
    )
    h3_style = ParagraphStyle(
        "H3_Style",
        fontName="Helvetica-Bold",
        fontSize=9.5,
        leading=12,
        textColor=colors.HexColor("#2D3748"),
        spaceBefore=6,
        spaceAfter=3
    )
    title_style = ParagraphStyle(
        "Title_Style",
        fontName="Helvetica-Bold",
        fontSize=18,
        leading=22,
        textColor=colors.HexColor("#1A365D"),
        alignment=1,
        spaceAfter=8
    )
    body_style = ParagraphStyle(
        "Body_Style",
        fontName="Helvetica",
        fontSize=8.5,
        leading=11.5,
        textColor=colors.HexColor("#2D3748"),
        spaceAfter=4
    )
    bullet_style = ParagraphStyle(
        "Bullet_Style",
        fontName="Helvetica",
        fontSize=8.5,
        leading=11.5,
        textColor=colors.HexColor("#2D3748"),
        leftIndent=15,
        spaceAfter=3
    )
    cell_style = ParagraphStyle(
        "Cell_Style",
        fontName="Helvetica",
        fontSize=7.2,
        leading=9.2,
        textColor=colors.HexColor("#2D3748")
    )
    cell_bold_style = ParagraphStyle(
        "Cell_Bold_Style",
        fontName="Helvetica-Bold",
        fontSize=7.2,
        leading=9.2,
        textColor=colors.HexColor("#1A365D")
    )
    header_cell_style = ParagraphStyle(
        "Header_Cell_Style",
        fontName="Helvetica-Bold",
        fontSize=7.5,
        leading=9.5,
        textColor=colors.white
    )

    story = []

    # Iterate through body elements in document order
    for elem in doc.element.body:
        tag = elem.tag.split("}")[-1]
        
        if tag == "p":
            p = docx.text.paragraph.Paragraph(elem, doc)
            text = p.text.strip()
            if not text:
                # check for images
                imgs = elem.xpath('.//a:blip/@r:embed')
                if not imgs:
                    continue
            
            p_style_name = p.style.name.lower() if p.style else ""
            
            # Format text with bold/italic runs
            runs_html = []
            for r in p.runs:
                r_txt = saxutils.escape(r.text)
                if r.bold:
                    r_txt = f"<b>{r_txt}</b>"
                if r.italic:
                    r_txt = f"<i>{r_txt}</i>"
                runs_html.append(r_txt)
            p_html = "".join(runs_html).strip()
            if not p_html:
                p_html = saxutils.escape(text)

            if "title" in p_style_name or ("store wise sales" in text.lower() and len(text) < 50):
                story.append(Paragraph(p_html, title_style))
            elif "heading 1" in p_style_name or (len(text) < 70 and text.startswith(tuple(f"{i}." for i in range(1, 20)))):
                story.append(Paragraph(p_html, h1_style))
            elif "heading 2" in p_style_name:
                story.append(Paragraph(p_html, h2_style))
            elif "heading 3" in p_style_name:
                story.append(Paragraph(p_html, h3_style))
            elif "list" in p_style_name or p_html.startswith(("&bull;", "-", "•")):
                clean_bullet = p_html.lstrip("-•&bull;").strip()
                story.append(Paragraph(f"&bull; {clean_bullet}", bullet_style))
            else:
                story.append(Paragraph(p_html, body_style))

        elif tag == "tbl":
            table = docx.table.Table(elem, doc)
            if not table.rows:
                continue
            
            num_cols = len(table.columns)
            if num_cols == 0:
                continue

            available_width = 523 # A4 width 595 - 72
            col_width = available_width / num_cols
            
            # Proportional col widths if known (e.g. 2 cols cover table)
            if num_cols == 2:
                widths = [150, 373]
            elif num_cols == 3:
                widths = [100, 223, 200]
            elif num_cols == 4:
                widths = [80, 150, 150, 143]
            elif num_cols == 5:
                widths = [60, 100, 120, 120, 123]
            elif num_cols == 6:
                widths = [50, 80, 100, 140, 80, 73]
            else:
                widths = [col_width] * num_cols

            table_data = []
            for r_idx, row in enumerate(table.rows):
                row_cells = []
                for c_idx, cell in enumerate(row.cells):
                    cell_txt = saxutils.escape(cell.text.strip())
                    if r_idx == 0 and num_cols > 2:
                        row_cells.append(Paragraph(cell_txt, header_cell_style))
                    elif c_idx == 0 and num_cols == 2:
                        row_cells.append(Paragraph(f"<b>{cell_txt}</b>", cell_bold_style))
                    else:
                        row_cells.append(Paragraph(cell_txt, cell_style))
                table_data.append(row_cells)

            tbl_obj = Table(table_data, colWidths=widths, repeatRows=1 if num_cols > 2 else 0)
            
            t_style = [
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E0")),
                ('PADDING', (0,0), (-1,-1), 3),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ]
            if num_cols > 2:
                t_style.append(('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1A365D")))
                t_style.append(('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.HexColor("#FFFFFF"), colors.HexColor("#F7FAFC")]))
            else:
                t_style.append(('BACKGROUND', (0,0), (0,-1), colors.HexColor("#EDF2F7")))
                t_style.append(('BACKGROUND', (1,0), (1,-1), colors.HexColor("#FFFFFF")))

            tbl_obj.setStyle(TableStyle(t_style))
            story.append(tbl_obj)
            story.append(Spacer(1, 6))

    # Build canvas with title
    canvas_class = DocxNumberedCanvas
    canvas_class._doc_title = custom_title or os.path.basename(pdf_path).replace(".pdf", "").replace("_", " ")

    # Check if ER diagram PNG can be attached if converting 09_ER_Diagram
    if "09_ER_Diagram" in docx_path and os.path.exists("docs/09_ER_Diagram.png"):
        story.append(Spacer(1, 10))
        story.append(Paragraph("<b>Figure: Entity Relationship (ER) Conceptual & Logical Architecture</b>", h2_style))
        story.append(Image("docs/09_ER_Diagram.png", width=520, height=292))
        story.append(Spacer(1, 10))

    pdf_doc.build(story, canvasmaker=canvas_class)
    print(f"[CONVERTED] {docx_path} -> {pdf_path}")
    return True

if __name__ == "__main__":
    test_file = "docs/01_Project_Understanding.docx"
    test_out = "docs/01_Project_Understanding.pdf"
    docx_to_pdf(test_file, test_out)
