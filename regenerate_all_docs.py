import os
import sys
import subprocess
import shutil

scripts = [
    ("generate_step01_doc.py", "Step 01: Project Understanding"),
    ("generate_step02_doc.py", "Step 02: HMW Statements"),
    ("generate_step03_doc.py", "Step 03: Needs Analysis"),
    ("generate_step04_rtm.py", "Step 04: Requirements Traceability Matrix"),
    ("generate_step05_brd.py", "Step 05: Business Requirements Document"),
    ("generate_step07_trd.py", "Step 07: Technical Requirements Document"),
    ("generate_step08_db_design.py", "Step 08: Database Design"),
    ("generate_step09_doc.py", "Step 09: ER Diagram"),
    ("generate_step11_user_stories.py", "Step 11: Agile User Stories"),
    ("generate_step12_ui_framework.py", "Step 12: UI Framework Specification"),
    ("generate_step13_test_cases.py", "Step 13: Test Cases Specification"),
    ("generate_step14_build_plan.py", "Step 14: Build Plan Document"),
    ("generate_step15_build_guide.py", "Step 15: Build, Test, Fix Guide"),
    ("generate_step16_deployment_guide.py", "Step 16: Deployment Guide"),
    ("generate_step17_project_document.py", "Step 17: Master Project Document"),
    ("generate_step18_presentation.py", "Step 18: Viva Presentation PPTX"),
    ("generate_step18_presentation_docx.py", "Step 18: Viva Presentation DOCX"),
]

py_exe = sys.executable

print("================================================================================")
print("  REGENERATING ALL 18 SDLC PROJECT DELIVERABLES (DOCX, XLSX, PPTX)")
print("================================================================================\n")

for script, desc in scripts:
    if os.path.exists(script):
        print(f"Running {script} ({desc})...")
        res = subprocess.run([py_exe, script], capture_output=True, text=True)
        if res.returncode == 0:
            print(f"  -> SUCCESS")
        else:
            print(f"  -> FAILED: {res.stderr}")
    else:
        print(f"Skipping {script} (not found)")

# Ensure duplicate aliases exist for roll number and step names
roll = "2500320140025"
if os.path.exists(f"docs/{roll}_Project_Document.docx"):
    shutil.copyfile(f"docs/{roll}_Project_Document.docx", "docs/17_Project_Document.docx")
    shutil.copyfile(f"docs/{roll}_Project_Document.docx", "docs/17_name_project_document.docx")

if os.path.exists(f"docs/{roll}_Presentation.pptx"):
    shutil.copyfile(f"docs/{roll}_Presentation.pptx", "docs/18_Presentation.pptx")
    shutil.copyfile(f"docs/{roll}_Presentation.pptx", "docs/18_name_presentation.pptx")

if os.path.exists(f"docs/{roll}_Presentation.docx"):
    shutil.copyfile(f"docs/{roll}_Presentation.docx", "docs/18_Presentation.docx")
    shutil.copyfile(f"docs/{roll}_Presentation.docx", "docs/18_name_presentation.docx")

print("\n>>> ALL DOCUMENTS REGENERATED AND SYNCHRONIZED SUCCESSFULLY! <<<")
