import os
import sys
import shutil
from convert_docx_to_pdf import docx_to_pdf

DOCS_TO_CONVERT = [
    # Master Project Document (Item 2 & 3)
    ("docs/2500320140025_Project_Document.docx", "docs/2500320140025_Project_Document.pdf", "P052: Master Capstone Project Report"),
    
    # Step-by-Step Deliverables (Item 8)
    ("docs/01_Project_Understanding.docx", "docs/01_Project_Understanding.pdf", "Step 01: Project Understanding"),
    ("docs/02_HMW_Statements.docx", "docs/02_HMW_Statements.pdf", "Step 02: How Might We (HMW) Statements"),
    ("docs/03_Needs_Analysis.docx", "docs/03_Needs_Analysis.pdf", "Step 03: Stakeholder Needs Analysis"),
    ("docs/05_BRD.docx", "docs/05_BRD.pdf", "Step 05: Business Requirements Document (BRD)"),
    ("docs/07_TRD.docx", "docs/07_TRD.pdf", "Step 07: Technical Requirements Document (TRD)"),
    ("docs/09_ER_Diagram.docx", "docs/09_ER_Diagram.pdf", "Step 09: Entity Relationship (ER) Diagram"),
    ("docs/10_Dataset_Generation.docx", "docs/10_Data_Validation_Report.pdf", "Step 10: Data Generation & Quality Audit Report"),
    ("docs/11_User_Stories.docx", "docs/11_User_Stories.pdf", "Step 11: Agile User Stories & Sprint Plan"),
    ("docs/12_UI_Framework.docx", "docs/12_UI_Framework.pdf", "Step 12: Streamlit UI Framework Specification"),
    ("docs/13_Test_Cases.docx", "docs/13_Test_Cases.pdf", "Step 13: Functional & Technical Test Cases"),
    ("docs/14_BUILD_PLAN.docx", "docs/14_BUILD_PLAN.pdf", "Step 14: Master SDLC Build Plan"),
    ("docs/15_Build_Test_Fix.docx", "docs/15_Build_Test_Fix.pdf", "Step 15: Build, Test & Verification Log"),
    ("docs/16_Deployment_Guide.docx", "docs/16_Deployment_Guide.pdf", "Step 16: Deployment & Operations Guide"),
]

print("================================================================================")
print("  CONVERTING ALL PROJECT DELIVERABLES (DOCX -> PDF)")
print("================================================================================\n")

for docx_path, pdf_path, title in DOCS_TO_CONVERT:
    docx_to_pdf(docx_path, pdf_path, custom_title=title)

# Duplicate aliases for user convenience
if os.path.exists("docs/2500320140025_Project_Document.pdf"):
    shutil.copyfile("docs/2500320140025_Project_Document.pdf", "docs/Project_Report.pdf")
    shutil.copyfile("docs/2500320140025_Project_Document.pdf", "docs/17_Project_Document.pdf")
    print("[COPIED] docs/Project_Report.pdf")

print("\n>>> ALL PROJECT DOCUMENTS CONVERTED TO HIGH-RESOLUTION PDFS! <<<")
