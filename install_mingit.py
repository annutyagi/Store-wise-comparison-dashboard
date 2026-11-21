import os
import sys
import urllib.request
import zipfile
import shutil

mingit_url = "https://github.com/git-for-windows/git/releases/download/v2.56.1.windows.1/MinGit-2.56.1-64-bit.zip"
target_dir = os.path.expanduser(r"~\AppData\Local\Programs\MinGit")
zip_path = os.path.expanduser(r"~\AppData\Local\Temp\MinGit.zip")

print(f"Downloading MinGit from {mingit_url}...")
req = urllib.request.Request(mingit_url, headers={"User-Agent": "Mozilla/5.0"})
with urllib.request.urlopen(req) as resp, open(zip_path, "wb") as out_file:
    shutil.copyfileobj(resp, out_file)

print(f"Extracting MinGit to {target_dir}...")
os.makedirs(target_dir, exist_ok=True)
with zipfile.ZipFile(zip_path, "r") as z:
    z.extractall(target_dir)

if os.path.exists(zip_path):
    os.remove(zip_path)

git_exe = os.path.join(target_dir, "cmd", "git.exe")
print(f"MinGit successfully installed at: {git_exe}")
