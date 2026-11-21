import os
import re
import sys
import subprocess

BATCH_ID = "ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01"
PROJECT_ID = "P052"
ROLL_NO = "2500320140025"
MENTOR = "Mr. Dinesh"
DOMAIN = "Data Analytics"

print("================================================================================")
print("  EXECUTING COMPREHENSIVE REPOSITORY METADATA REVISION")
print(f"  1. Roll No        : {ROLL_NO}")
print(f"  2. Batch ID       : {BATCH_ID}")
print(f"  3. Project ID     : {PROJECT_ID}")
print(f"  4. Project Mentor : {MENTOR}")
print(f"  5. Domain         : {DOMAIN}")
print("================================================================================\n")

# 1. Update CLAUDE.md
with open("CLAUDE.md", "r", encoding="utf-8") as f:
    claude_txt = f.read()
claude_txt = re.sub(r"Prof\. Kalpna Dwivedi", MENTOR, claude_txt)
claude_txt = re.sub(r"Kalpna Dwivedi", MENTOR, claude_txt)
claude_txt = re.sub(r"Kalpna", MENTOR, claude_txt)
with open("CLAUDE.md", "w", encoding="utf-8") as f:
    f.write(claude_txt)
print("[OK] CLAUDE.md updated.")

# 2. Update README.md
with open("README.md", "r", encoding="utf-8") as f:
    readme_txt = f.read()
readme_txt = re.sub(r"Prof\. Kalpna Dwivedi", MENTOR, readme_txt)
readme_txt = re.sub(r"Kalpna Dwivedi", MENTOR, readme_txt)
readme_txt = re.sub(r"\*\*Faculty Project Mentor:\*\*.*", f"**Faculty Project Mentor:** {MENTOR} (Department of Computer Applications)  ", readme_txt)
readme_txt = re.sub(r"\*\*Domain:\*\*.*", f"**Domain:** {DOMAIN}  \n**Project ID:** {PROJECT_ID}  \n**Batch ID:** {BATCH_ID}  ", readme_txt)
with open("README.md", "w", encoding="utf-8") as f:
    f.write(readme_txt)
print("[OK] README.md updated.")

# 3. Update BUILD_PLAN.md
with open("BUILD_PLAN.md", "r", encoding="utf-8") as f:
    bp_txt = f.read()
bp_txt = re.sub(r"Prof\. Kalpna Dwivedi", MENTOR, bp_txt)
bp_txt = re.sub(r"Kalpna Dwivedi", MENTOR, bp_txt)
bp_txt = re.sub(r"\*\*Academic Mentor:\*\*.*", f"**Academic Mentor:** {MENTOR}  \n**Project ID:** {PROJECT_ID}  \n**Batch ID:** {BATCH_ID}  \n**Domain:** {DOMAIN}  ", bp_txt)
with open("BUILD_PLAN.md", "w", encoding="utf-8") as f:
    f.write(bp_txt)
print("[OK] BUILD_PLAN.md updated.")

# 4. Update src/ui.py
with open("src/ui.py", "r", encoding="utf-8") as f:
    ui_txt = f.read()
ui_txt = re.sub(r"Prof\. Kalpna Dwivedi", MENTOR, ui_txt)
ui_txt = re.sub(r"Kalpna Dwivedi", MENTOR, ui_txt)
ui_txt = ui_txt.replace(
    '<span class="text-xs font-medium text-slate-400">ABES MCA Capstone Project</span>',
    f'<span class="text-xs font-medium text-slate-400">ABES MCA Capstone | Project ID: {PROJECT_ID} | {BATCH_ID}</span>'
)
with open("src/ui.py", "w", encoding="utf-8") as f:
    f.write(ui_txt)
print("[OK] src/ui.py updated.")

# 5. Update run_tests.py
with open("run_tests.py", "r", encoding="utf-8") as f:
    rt_txt = f.read()
rt_txt = re.sub(r"Author: Anu Tyagi \|.*", f"Project ID: {PROJECT_ID} | Author: Anu Tyagi | Roll: {ROLL_NO} | Batch: {BATCH_ID} | Mentor: {MENTOR}", rt_txt)
with open("run_tests.py", "w", encoding="utf-8") as f:
    f.write(rt_txt)
print("[OK] run_tests.py updated.")

# 6. Update generate_step04_rtm.py
with open("generate_step04_rtm.py", "r", encoding="utf-8") as f:
    rtm_code = f.read()
rtm_code = re.sub(r'ws1\["A2"\] = .*', f'ws1["A2"] = "Project ID: {PROJECT_ID} | Student: Anu Tyagi | Roll: {ROLL_NO} | Batch: {BATCH_ID} | Domain: {DOMAIN} | Mentor: {MENTOR} | Updated: Step 13"', rtm_code)
with open("generate_step04_rtm.py", "w", encoding="utf-8") as f:
    f.write(rtm_code)
print("[OK] generate_step04_rtm.py updated.")

# 7. Update generate_step17_project_document.py
with open("generate_step17_project_document.py", "r", encoding="utf-8") as f:
    pdoc = f.read()
pdoc = re.sub(r"Prof\. Kalpna Dwivedi", MENTOR, pdoc)
pdoc = re.sub(r"Kalpna Dwivedi", MENTOR, pdoc)
old_cov_block = """    cov_table = doc.add_table(rows=8, cols=2)
    cov_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cov_data = [
        ("Candidate Name", "Anu Tyagi"),
        ("University Roll Number", "2500320140025"),
        ("Academic Program & Year", "Master of Computer Applications (MCA), 2nd Year"),
        ("Specialization Track", "Data Analytics & Business Intelligence"),
        ("Faculty Project Mentor", "Prof. Kalpna Dwivedi (Department of Computer Applications)"),
        ("Institutional Affiliation", "ABES Engineering College, Ghaziabad (AKTU Code: 032)"),
        ("Public GitHub Repository", "https://github.com/anutyagi25/2500320140025-store-wise-sales"),
        ("Live Streamlit Cloud URL", "https://2500320140025-store-wise-sales.streamlit.app")
    ]"""

new_cov_block = f"""    cov_table = doc.add_table(rows=10, cols=2)
    cov_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cov_data = [
        ("Project ID", "{PROJECT_ID}"),
        ("Domain", "{DOMAIN}"),
        ("Batch ID", "{BATCH_ID}"),
        ("Candidate Name", "Anu Tyagi"),
        ("University Roll Number", "{ROLL_NO}"),
        ("Academic Program & Year", "Master of Computer Applications (MCA), 2nd Year"),
        ("Faculty Project Mentor", "{MENTOR} (Department of Computer Applications)"),
        ("Institutional Affiliation", "ABES Engineering College, Ghaziabad (AKTU Code: 032)"),
        ("Public GitHub Repository", "https://github.com/anutyagi25/2500320140025-store-wise-sales"),
        ("Live Streamlit Cloud URL", "https://2500320140025-store-wise-sales.streamlit.app")
    ]"""
if old_cov_block in pdoc:
    pdoc = pdoc.replace(old_cov_block, new_cov_block)
else:
    # generic regex replace for cov_data
    pdoc = re.sub(r'\("Specialization Track", "[^"]*"\)', f'("Domain", "{DOMAIN}"),\n        ("Project ID", "{PROJECT_ID}"),\n        ("Batch ID", "{BATCH_ID}")', pdoc)
    pdoc = re.sub(r'rows=8, cols=2', 'rows=10, cols=2', pdoc)

with open("generate_step17_project_document.py", "w", encoding="utf-8") as f:
    f.write(pdoc)
print("[OK] generate_step17_project_document.py updated.")

# 8. Update generate_step18_presentation.py and docx
for p_script in ["generate_step18_presentation.py", "generate_step18_presentation_docx.py"]:
    with open(p_script, "r", encoding="utf-8") as f:
        pres_txt = f.read()
    pres_txt = re.sub(r"Prof\. Kalpna Dwivedi", MENTOR, pres_txt)
    pres_txt = re.sub(r"Kalpna Dwivedi", MENTOR, pres_txt)
    pres_txt = re.sub(r"Academic Track: Master of Computer Applications.*", f"Project ID: {PROJECT_ID} | Domain: {DOMAIN} | Batch: {BATCH_ID}\",", pres_txt)
    with open(p_script, "w", encoding="utf-8") as f:
        f.write(pres_txt)
    print(f"[OK] {p_script} updated.")

# 9. Update all other generator files
all_gen = [
    "generate_step01_doc.py",
    "generate_step02_doc.py",
    "generate_step03_doc.py",
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
    "validate_dataset.py"
]

for gf in all_gen:
    if os.path.exists(gf):
        with open(gf, "r", encoding="utf-8") as f:
            t = f.read()
        t = re.sub(r"Prof\. Kalpna Dwivedi", MENTOR, t)
        t = re.sub(r"Kalpna Dwivedi", MENTOR, t)
        t = re.sub(r'\(("Course & Branch:"|"Course / Branch:"), "[^"]*"\)', f'("Batch ID:", "{BATCH_ID}")', t)
        t = re.sub(r'\(("Domain:"|"Domain"), "[^"]*"\)', f'("Domain:", "{DOMAIN}")', t)
        t = re.sub(r'\(("Project ID:"|"Project ID"), "[^"]*"\)', f'("Project ID:", "{PROJECT_ID}")', t)
        with open(gf, "w", encoding="utf-8") as f:
            f.write(t)
        print(f"[OK] {gf} updated.")

print("\n--- All Source Code & Generator Scripts Successfully Updated ---")
