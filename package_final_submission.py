import os
import shutil
import zipfile

target_dir = "docs/Final_Submission_PDFs"
os.makedirs(target_dir, exist_ok=True)

files_to_copy = [
    # Step 1 through Step 16
    ("docs/01_Project_Understanding.pdf", "01_Project_Understanding.pdf"),
    ("docs/02_HMW_Statements.pdf", "02_HMW_Statements.pdf"),
    ("docs/03_Needs_Analysis.pdf", "03_Needs_Analysis.pdf"),
    ("docs/04_Traceability_Matrix.pdf", "04_Traceability_Matrix.pdf"),
    ("docs/04_Traceability_Matrix.xlsx", "04_Traceability_Matrix.xlsx"),
    ("docs/05_BRD.pdf", "05_BRD.pdf"),
    ("docs/07_TRD.pdf", "07_TRD.pdf"),
    ("docs/08_Database_Design.pdf", "08_Database_Design.pdf"),
    ("docs/08_Database_Design.xlsx", "08_Database_Design.xlsx"),
    ("docs/09_ER_Diagram.pdf", "09_ER_Diagram.pdf"),
    ("docs/10_Data_Validation_Report.pdf", "10_Data_Validation_Report.pdf"),
    ("docs/11_User_Stories.pdf", "11_User_Stories.pdf"),
    ("docs/12_UI_Framework.pdf", "12_UI_Framework.pdf"),
    ("docs/13_Test_Cases.pdf", "13_Test_Cases.pdf"),
    ("docs/14_BUILD_PLAN.pdf", "14_BUILD_PLAN.pdf"),
    ("docs/15_Build_Test_Fix.pdf", "15_Build_Test_Fix.pdf"),
    ("docs/16_Deployment_Guide.pdf", "16_Deployment_Guide.pdf"),
    
    # Master Project Document (Step 17)
    ("docs/2500320140025_Project_Document.pdf", "17_Master_Project_Document.pdf"),
    ("docs/2500320140025_Project_Document.docx", "17_Master_Project_Document.docx"),
    
    # Project Presentation (Step 18)
    ("docs/2500320140025_Presentation.pdf", "18_Project_Presentation.pdf"),
    ("docs/2500320140025_Presentation.pptx", "18_Project_Presentation.pptx"),
    
    # Complete Source Code PDF
    ("docs/2500320140025_Source_Code.pdf", "Complete_Source_Code.pdf"),
    
    # GitHub Report PDF
    ("docs/GitHub_Report.pdf", "GitHub_Repository_Report.pdf"),
    
    # Project Database in XLSX
    ("docs/Project_Database.xlsx", "Project_Database.xlsx")
]

print("================================================================================")
print("  PACKAGING FINAL SUBMISSION DELIVERABLES INTO docs/Final_Submission_PDFs/")
print("================================================================================\n")

for src, dst in files_to_copy:
    if os.path.exists(src):
        dst_path = os.path.join(target_dir, dst)
        shutil.copyfile(src, dst_path)
        print(f"  [COPIED] {dst} ({os.path.getsize(dst_path):,} bytes)")
    else:
        print(f"  [MISSING] {src}")

# Build ZIP archive for one-click download
zip_path = "docs/Final_Submission_Package.zip"
with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
    for root, dirs, files in os.walk(target_dir):
        for f in files:
            full_p = os.path.join(root, f)
            arc_name = os.path.relpath(full_p, target_dir)
            z.write(full_p, arc_name)

print(f"\n[ZIP CREATED] {zip_path} ({os.path.getsize(zip_path):,} bytes)")
print(">>> FINAL SUBMISSION PACKAGING COMPLETE! <<<")
