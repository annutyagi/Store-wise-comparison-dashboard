import os
import re
import sys
import subprocess
import shutil

NEW_REPO_URL = "https://github.com/annutyagi/Store-wise-comparison-dashboard"
NEW_REPO_SHORT = "github.com/annutyagi/Store-wise-comparison-dashboard"
NEW_STREAMLIT_URL = "https://store-wise-comparison-dashboard.streamlit.app"
NEW_STREAMLIT_SHORT = "store-wise-comparison-dashboard.streamlit.app"
ACADEMIC_YEAR = "2026-2027"

print("================================================================================")
print("  APPLYING USER SPECIFIC UPDATES ACROSS ENTIRE PROJECT")
print(f"  1. Academic Year: 2026-2027")
print(f"  2. GitHub Repo:   {NEW_REPO_URL}")
print(f"  3. Mentor Name:   Mr. Dinesh (removed 'Department of Computer Applications')")
print(f"  4. Institution:   ABES Engineering College, Ghaziabad (removed 'AKTU Code: 032')")
print("================================================================================\n")

def replace_in_file(path, replacements):
    if not os.path.exists(path):
        return
    with open(path, "r", encoding="utf-8") as f:
        content = f.read()
    
    modified = content
    for pattern, repl in replacements:
        modified = re.sub(pattern, repl, modified)
    
    if modified != content:
        with open(path, "w", encoding="utf-8") as f:
            f.write(modified)
        print(f"[UPDATED] {path}")

# 1. Update markdown files
replace_in_file("README.md", [
    (r"https://github\.com/[^\s)]+", NEW_REPO_URL),
    (r"\*\*Faculty Project Mentor:\*\* Mr\. Dinesh.*", "**Faculty Project Mentor:** Mr. Dinesh  "),
    (r"\(AKTU Code: 032\)", ""),
    (r", AKTU", ", Ghaziabad"),
    (r"2024–2026", ACADEMIC_YEAR),
    (r"2024-2026", ACADEMIC_YEAR),
])

replace_in_file("CLAUDE.md", [
    (r"https://github\.com/[^\s)]+", NEW_REPO_URL),
    (r"\(AKTU Code: 032\)", ""),
    (r", AKTU", ", Ghaziabad"),
    (r"2024–2026", ACADEMIC_YEAR),
    (r"2024-2026", ACADEMIC_YEAR),
])

replace_in_file("BUILD_PLAN.md", [
    (r"https://github\.com/[^\s)]+", NEW_REPO_URL),
    (r"\(AKTU Code: 032\)", ""),
    (r", AKTU", ", Ghaziabad"),
    (r"2024–2026", ACADEMIC_YEAR),
    (r"2024-2026", ACADEMIC_YEAR),
])

# 2. Update all generator scripts
all_py = [f for f in os.listdir(".") if f.endswith(".py") and f.startswith("generate_step")] + ["validate_dataset.py", "verify_metadata.py"]

for py_file in all_py:
    replace_in_file(py_file, [
        (r"Mr\. Dinesh \(Department of Computer Applications\)", "Mr. Dinesh"),
        (r"Mr\. Dinesh \(Dept of CA, ABES\)", "Mr. Dinesh"),
        (r"Mr\. Dinesh \(Assistant Professor, ABES\)", "Mr. Dinesh"),
        (r"Mr\. Dinesh \(Faculty Mentor, ABES\)", "Mr. Dinesh"),
        (r"ABES Engineering College, Ghaziabad \(AKTU Code: 032\)", "ABES Engineering College, Ghaziabad"),
        (r"ABES Engineering College, AKTU", "ABES Engineering College, Ghaziabad"),
        (r"\(AKTU Code: 032\)", ""),
        (r"2024–2026", ACADEMIC_YEAR),
        (r"2024-2026", ACADEMIC_YEAR),
        (r"https://github\.com/anutyagi25/2500320140025-store-wise-sales", NEW_REPO_URL),
        (r"github\.com/anutyagi25/2500320140025-store-wise-sales", NEW_REPO_SHORT),
        (r"https://2500320140025-store-wise-sales\.streamlit\.app", NEW_STREAMLIT_URL),
        (r"2500320140025-store-wise-sales\.streamlit\.app", NEW_STREAMLIT_SHORT),
    ])

print("\n--- All code and generator scripts successfully updated. Regenerating documents... ---")

# Run regeneration
scripts = [
    "generate_step01_doc.py",
    "generate_step02_doc.py",
    "generate_step03_doc.py",
    "generate_step04_rtm.py",
    "generate_step05_brd.py",
    "generate_step07_trd.py",
    "generate_step08_db_design.py",
    "generate_step09_doc.py",
    "generate_step11_user_stories.py",
    "generate_step12_ui_framework.py",
    "generate_step13_test_cases.py",
    "generate_step14_build_plan.py",
    "generate_step15_build_guide.py",
    "generate_step16_deployment_guide.py",
    "generate_step17_project_document.py",
    "generate_step18_presentation.py",
    "generate_step18_presentation_docx.py",
]

py_exe = sys.executable
for s in scripts:
    res = subprocess.run([py_exe, s], capture_output=True, text=True)
    if res.returncode == 0:
        print(f"[REGENERATED] {s}")
    else:
        print(f"[ERROR] {s}: {res.stderr}")

# Copy aliases
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
