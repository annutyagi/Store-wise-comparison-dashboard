import os
import sys
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle, Preformatted
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfgen import canvas

CODE_FILES = [
    # 1. Main Entrypoint
    "app.py",
    # 2. Multipage Screens
    "pages/1_Store_Comparison.py",
    "pages/2_Category_Demographics.py",
    "pages/3_Temporal_Trends.py",
    "pages/4_Target_Variance.py",
    "pages/5_Basket_Dynamics.py",
    "pages/6_Data_Export.py",
    "pages/7_Data_Quality.py",
    # 3. Source Packages
    "src/data_loader.py",
    "src/cleaning.py",
    "src/analytics.py",
    "src/ui.py",
    "src/sidebar.py",
    "src/generate_data.py",
    # 4. Tests
    "tests/test_data.py",
    "tests/test_analysis.py",
    "tests/test_app.py",
    "run_tests.py",
]

class NumberedCanvas(canvas.Canvas):
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
            self.draw_page_number(num_pages)
            super().showPage()
        super().save()

    def draw_page_number(self, page_count):
        self.saveState()
        self.setFont("Helvetica", 8)
        self.setFillColor(colors.HexColor("#718096"))
        
        # Header (pages > 1)
        if self._pageNumber > 1:
            self.drawString(36, 806, "P052: Store Wise Sales Comparison Dashboard — Complete Source Code Repository")
            self.drawRightString(559, 806, "Candidate: Anu Tyagi (Roll: 2500320140025)")
            self.setStrokeColor(colors.HexColor("#CBD5E0"))
            self.setLineWidth(0.5)
            self.line(36, 800, 559, 800)
            
        # Footer
        page_str = f"Page {self._pageNumber} of {page_count}"
        self.drawString(36, 25, "ABES Engineering College, Ghaziabad | MCA 2nd Year (2026-2027)")
        self.drawRightString(559, 25, page_str)
        self.setStrokeColor(colors.HexColor("#CBD5E0"))
        self.setLineWidth(0.5)
        self.line(36, 35, 559, 35)
        self.restoreState()

def build_source_code_pdf(output_paths):
    for out_path in output_paths:
        os.makedirs(os.path.dirname(out_path), exist_ok=True)
        doc = SimpleDocTemplate(
            out_path,
            pagesize=A4,
            leftMargin=36,
            rightMargin=36,
            topMargin=46,
            bottomMargin=46
        )

        styles = getSampleStyleSheet()
        
        # Custom Typography Styles
        title_style = ParagraphStyle(
            "CoverTitle",
            parent=styles["Title"],
            fontName="Helvetica-Bold",
            fontSize=22,
            leading=26,
            textColor=colors.HexColor("#1A365D"),
            alignment=1,
            spaceAfter=8
        )
        sub_style = ParagraphStyle(
            "CoverSub",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=11,
            leading=15,
            textColor=colors.HexColor("#4A5568"),
            alignment=1,
            spaceAfter=20
        )
        file_header_style = ParagraphStyle(
            "FileHeader",
            parent=styles["Heading2"],
            fontName="Helvetica-Bold",
            fontSize=12,
            leading=16,
            textColor=colors.white,
            spaceAfter=0
        )
        code_style = ParagraphStyle(
            "CodeBlock",
            fontName="Courier",
            fontSize=6.8,
            leading=8.6,
            textColor=colors.HexColor("#1A202C")
        )

        story = []

        # -------------------------------------------------------------
        # COVER / TITLE BLOCK
        # -------------------------------------------------------------
        story.append(Spacer(1, 20))
        story.append(Paragraph("ABES ENGINEERING COLLEGE, GHAZIABAD", ParagraphStyle("H1", fontName="Helvetica-Bold", fontSize=13, alignment=1, textColor=colors.HexColor("#2B6CB0"))))
        story.append(Paragraph("MASTER OF COMPUTER APPLICATIONS (MCA) 2026-2027", ParagraphStyle("H2", fontName="Helvetica", fontSize=10, alignment=1, textColor=colors.HexColor("#718096"), spaceAfter=15)))
        story.append(Spacer(1, 10))
        story.append(Paragraph("STORE WISE SALES COMPARISON DASHBOARD", title_style))
        story.append(Paragraph("Complete Technical Source Code Repository Deliverable (ITEM 5)", sub_style))

        # Metadata Table
        meta_data = [
            [Paragraph("<b>Project ID</b>", styles["Normal"]), Paragraph("P052", styles["Normal"])],
            [Paragraph("<b>Domain Track</b>", styles["Normal"]), Paragraph("Data Analytics", styles["Normal"])],
            [Paragraph("<b>Batch ID</b>", styles["Normal"]), Paragraph("ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01", styles["Normal"])],
            [Paragraph("<b>Candidate Name</b>", styles["Normal"]), Paragraph("Anu Tyagi", styles["Normal"])],
            [Paragraph("<b>University Roll Number</b>", styles["Normal"]), Paragraph("2500320140025", styles["Normal"])],
            [Paragraph("<b>Faculty Project Mentor</b>", styles["Normal"]), Paragraph("Mr. Dinesh", styles["Normal"])],
            [Paragraph("<b>Institutional Affiliation</b>", styles["Normal"]), Paragraph("ABES Engineering College, Ghaziabad", styles["Normal"])],
            [Paragraph("<b>Public GitHub Repository</b>", styles["Normal"]), Paragraph("https://github.com/annutyagi/Store-wise-comparison-dashboard", styles["Normal"])],
            [Paragraph("<b>Live Streamlit Cloud URL</b>", styles["Normal"]), Paragraph("https://store-wise-comparison-dashboard-tfgpqzxbwqggso9wax3wfm.streamlit.app/", styles["Normal"])],
            [Paragraph("<b>Total Source Files Included</b>", styles["Normal"]), Paragraph(f"{len(CODE_FILES)} Files (Entrypoint, Multipage UI, Core Logic, Test Suite)", styles["Normal"])],
        ]
        t = Table(meta_data, colWidths=[160, 360])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#F7FAFC")),
            ('BOX', (0,0), (-1,-1), 1, colors.HexColor("#CBD5E0")),
            ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor("#E2E8F0")),
            ('PADDING', (0,0), (-1,-1), 4),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]))
        story.append(t)
        story.append(Spacer(1, 20))

        # Table of Contents
        story.append(Paragraph("<b>Table of Source Code Files</b>", ParagraphStyle("TOC", fontName="Helvetica-Bold", fontSize=11, textColor=colors.HexColor("#1A365D"), spaceAfter=6)))
        toc_data = [["#", "Module / File Name", "Category", "Lines", "Size (Bytes)"]]
        total_lines = 0
        total_bytes = 0
        for idx, fpath in enumerate(CODE_FILES, start=1):
            if os.path.exists(fpath):
                size = os.path.getsize(fpath)
                with open(fpath, "r", encoding="utf-8", errors="ignore") as fp:
                    lines = len(fp.readlines())
                total_lines += lines
                total_bytes += size
                cat = "Main App" if fpath == "app.py" else ("Page UI" if "pages/" in fpath else ("Core Package" if "src/" in fpath else "Test Suite"))
                toc_data.append([str(idx), fpath, cat, f"{lines:,}", f"{size:,}"])
            else:
                toc_data.append([str(idx), fpath, "Missing", "0", "0"])

        toc_data.append(["", "<b>TOTAL CODEBASE VOLUME</b>", "", f"<b>{total_lines:,}</b>", f"<b>{total_bytes:,}</b>"])
        
        toc_table = Table(toc_data, colWidths=[24, 210, 100, 70, 86])
        toc_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1A365D")),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('FONTNAME', (0,0), (-1,0), "Helvetica-Bold"),
            ('FONTSIZE', (0,0), (-1,0), 8),
            ('BOTTOMPADDING', (0,0), (-1,-1), 3),
            ('TOPPADDING', (0,0), (-1,-1), 3),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E0")),
            ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor("#EDF2F7")),
            ('ALIGN', (0,0), (0,-1), 'CENTER'),
            ('ALIGN', (3,0), (-1,-1), 'RIGHT'),
        ]))
        story.append(toc_table)
        story.append(PageBreak())

        # -------------------------------------------------------------
        # APPEND EACH CODE FILE
        # -------------------------------------------------------------
        for fpath in CODE_FILES:
            if not os.path.exists(fpath):
                continue
            with open(fpath, "r", encoding="utf-8", errors="ignore") as fp:
                lines = fp.readlines()
            
            # File Banner Table
            banner_data = [[
                Paragraph(f"<b>FILE: {fpath}</b>", file_header_style),
                Paragraph(f"<font color='white'>{len(lines)} Lines | {os.path.getsize(fpath):,} Bytes</font>", ParagraphStyle("Right", fontName="Helvetica", fontSize=9, alignment=2, textColor=colors.white))
            ]]
            banner_table = Table(banner_data, colWidths=[340, 180])
            banner_table.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#1A365D")),
                ('PADDING', (0,0), (-1,-1), 5),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ]))
            story.append(banner_table)
            story.append(Spacer(1, 4))

            # Code text with line numbers
            numbered_lines = []
            for l_idx, line_str in enumerate(lines, start=1):
                clean_line = line_str.rstrip("\r\n").replace("\t", "    ")
                numbered_lines.append(f"{l_idx:4d} | {clean_line}")

            code_text = "\n".join(numbered_lines)
            story.append(Preformatted(code_text, code_style))
            story.append(Spacer(1, 14))

        doc.build(story, canvasmaker=NumberedCanvas)
        print(f"Source Code PDF successfully created at: {out_path}")

if __name__ == "__main__":
    build_source_code_pdf([
        "docs/2500320140025_Source_Code.pdf",
        "docs/Source_Code.pdf"
    ])
