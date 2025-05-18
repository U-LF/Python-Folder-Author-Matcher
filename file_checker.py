import os
import re
import zipfile
import rarfile
from docx import Document

rarfile.UNRAR_TOOL = r"C:\Program Files\WinRAR\UnRAR.exe"

def get_author_from_docx(file_path):
    try:
        doc = Document(file_path)
        author = doc.core_properties.author
        return author.strip().upper() if author else None
    except Exception as e:
        print(f"[Error] Could not read '{file_path}': {e}")
        return None

def extract_zip_in_place(zip_path, target_folder):
    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(target_folder)
    except Exception as e:
        print(f"[Error] Extracting ZIP '{zip_path}': {e}")

def extract_rar_in_place(rar_path, target_folder):
    try:
        with rarfile.RarFile(rar_path, 'r') as rar_ref:
            rar_ref.extractall(target_folder)
    except Exception as e:
        print(f"[Error] Extracting RAR '{rar_path}': {e}")

def normalize_name(name):
    if not name:
        return ""
    return re.sub(r"\s+", " ", name).strip().upper()

def check_docx_in_folder(folder_path, folder_name):
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith(".docx"):
                file_path = os.path.join(root, file)
                author = get_author_from_docx(file_path)
                if not author:
                    print(f"[No Author] File: '{file}' in folder '{folder_name}'")
                    continue
                normalized_folder = normalize_name(folder_name)
                if author not in normalized_folder:
                    print(f"[Mismatch] Folder: '{folder_name}' | File: '{file}' | Author: '{author}'")
                # else:
                    # print(f"[Match] Folder: '{folder_name}' | File: '{file}' | Author: '{author}'")

def process_submissions(root_dir):
    for folder in os.listdir(root_dir):
        folder_path = os.path.join(root_dir, folder)
        if not os.path.isdir(folder_path):
            continue

        # Extract ZIPs and RARs in-place
        for file in os.listdir(folder_path):
            file_path = os.path.join(folder_path, file)
            if file.lower().endswith(".zip"):
                extract_zip_in_place(file_path, folder_path)
            elif file.lower().endswith(".rar"):
                extract_rar_in_place(file_path, folder_path)

        # Check all .docx files (including extracted ones)
        check_docx_in_folder(folder_path, folder)

# ======== USAGE ========
if __name__ == "__main__":
    root_directory = r"D:\OOP Mid W7"  # <- Update as needed
    process_submissions(root_directory)
