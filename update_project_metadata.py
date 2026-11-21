import os
import re
import sys
import subprocess

BATCH_ID = "ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01"
PROJECT_ID = "P052"
ROLL_NO = "2500320140025"
MENTOR = "Mr. Dinesh"
DOMAIN = "Data Analytics"

print(f"--- Updating Project Metadata Across Repository ---")
print(f"Roll No:        {ROLL_NO}")
print(f"Batch ID:       {BATCH_ID}")
print(f"Project ID:     {PROJECT_ID}")
print(f"Project Mentor: {MENTOR}")
print(f"Domain:         {DOMAIN}")

# 1. Update text files directly
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
        print(f"Updated: {path}")

# CLAUDE.md
replace_in_file("CLAUDE.md", [
    (r"Prof\. Kalpna Dwivedi", MENTOR),
    (r"Kalpna Dwivedi", MENTOR),
    (r"Kalpna", MENTOR),
])

# README.md
replace_in_file("README.md", [
    (r"\*\*Faculty Project Mentor:\*\*.*", f"**Faculty Project Mentor:** {MENTOR} (Department of Computer Applications)  "),
    (r"Prof\. Kalpna Dwivedi", MENTOR),
    (r"Kalpna Dwivedi", MENTOR),
    (r"\*\*Course / Track:\*\*.*", f"**Course / Track:** MCA 2nd Year ({DOMAIN})  \n**Batch ID:** {BATCH_ID}  \n**Project ID:** {PROJECT_ID}  "),
])

# BUILD_PLAN.md
replace_in_file("BUILD_PLAN.md", [
    (r"\*\*Academic Mentor:\*\*.*", f"**Academic Mentor:** {MENTOR}  \n**Project ID:** {PROJECT_ID}  \n**Batch ID:** {BATCH_ID}  \n**Domain:** {DOMAIN}  "),
    (r"Prof\. Kalpna Dwivedi", MENTOR),
    (r"Kalpna Dwivedi", MENTOR),
])

# run_tests.py
replace_in_file("run_tests.py", [
    (r"Author: Anu Tyagi \|.*", f"Project ID: {PROJECT_ID} | Author: Anu Tyagi | Roll: {ROLL_NO} | Batch: {BATCH_ID} | Mentor: {MENTOR}"),
])

# src/ui.py
replace_in_file("src/ui.py", [
    (r"Prof\. Kalpna Dwivedi", MENTOR),
    (r"Kalpna Dwivedi", MENTOR),
    (r'<span class="text-xs font-medium text-slate-400">ABES MCA Capstone Project</span>', 
     f'<span class="text-xs font-medium text-slate-400">ABES MCA Capstone | {PROJECT_ID} | {BATCH_ID}</span>'),
])

# Update all generator scripts
gen_files = [
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
    "validate_dataset.py"
]

for gf in gen_files:
    replace_in_file(gf, [
        (r"Prof\. Kalpna Dwivedi", MENTOR),
        (r"Kalpna Dwivedi", MENTOR),
        (r"Kalpna", MENTOR),
    ])

print("Code and generator files text replacements completed.")
