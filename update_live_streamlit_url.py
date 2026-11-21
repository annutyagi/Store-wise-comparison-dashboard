import os
import re
import sys
import subprocess

NEW_URL = "https://store-wise-comparison-dashboard-tfgpqzxbwqggso9wax3wfm.streamlit.app/"
NEW_URL_NO_SLASH = "https://store-wise-comparison-dashboard-tfgpqzxbwqggso9wax3wfm.streamlit.app"
NEW_SHORT = "store-wise-comparison-dashboard-tfgpqzxbwqggso9wax3wfm.streamlit.app"

files_to_update = [
    "generate_step17_project_document.py",
    "generate_step18_presentation.py",
    "generate_step18_presentation_docx.py",
    "generate_step16_deployment_guide.py",
    "generate_source_code_pdf.py",
    "generate_presentation_pdf.py",
    "generate_github_report_pdf.py",
    "generate_github_streamlit_details_pdf.py",
    "README.md",
    "BUILD_PLAN.md",
    "CLAUDE.md",
    "verify_metadata.py"
]

print("================================================================================")
print("  UPDATING LIVE STREAMLIT CLOUD URL ACROSS ALL PROJECT ASSETS")
print(f"  New Target URL: {NEW_URL}")
print("================================================================================\n")

for fpath in files_to_update:
    if os.path.exists(fpath):
        with open(fpath, "r", encoding="utf-8") as f:
            content = f.read()
        
        # Replace variations
        mod = re.sub(r"https://store-wise-comparison-dashboard\.streamlit\.app/?", NEW_URL, content)
        mod = re.sub(r"https://2500320140025-store-wise-sales\.streamlit\.app/?", NEW_URL, mod)
        mod = re.sub(r"store-wise-comparison-dashboard\.streamlit\.app", NEW_SHORT, mod)
        mod = re.sub(r"2500320140025-store-wise-sales\.streamlit\.app", NEW_SHORT, mod)
        
        if mod != content:
            with open(fpath, "w", encoding="utf-8") as f:
                f.write(mod)
            print(f"[UPDATED] {fpath}")

print("\n--- Regenerating Documents with New Live Streamlit URL ---")
py_exe = sys.executable

regen_scripts = [
    "generate_step16_deployment_guide.py",
    "generate_step17_project_document.py",
    "generate_step18_presentation.py",
    "generate_step18_presentation_docx.py",
    "convert_all_project_docs_to_pdf.py",
    "generate_presentation_pdf.py",
    "generate_source_code_pdf.py",
    "generate_github_report_pdf.py",
    "generate_github_streamlit_details_pdf.py",
    "package_final_submission.py"
]

for s in regen_scripts:
    res = subprocess.run([py_exe, s], capture_output=True, text=True)
    if res.returncode == 0:
        print(f"[REGENERATED] {s}")
    else:
        print(f"[ERROR] {s}: {res.stderr}")

print("\n>>> LIVE STREAMLIT URL SUCCESSFULLY SYNCHRONIZED ACROSS ALL ASSETS! <<<")
