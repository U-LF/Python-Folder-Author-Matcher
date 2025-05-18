import os
import re
from docx import Document

def get_docx_author(file_path):
    try:
        doc = Document(file_path)
        author = doc.core_properties.author
        return normalize_name(author) if author else None
    except Exception as e:
        print(f"Error reading '{file_path}': {e}")
        return None

def normalize_name(name):
    if not name:
        return ""
    return re.sub(r"\s+", " ", name).strip().upper()

def check_author_in_folder_name(base_path):
    for folder in os.listdir(base_path):
        folder_path = os.path.join(base_path, folder)
        if os.path.isdir(folder_path):
            normalized_folder = normalize_name(folder)
            for file in os.listdir(folder_path):
                if file.lower().endswith(".docx"):
                    file_path = os.path.join(folder_path, file)
                    author = get_docx_author(file_path)

                    if not author:
                        print(f"[No Author] File: '{file}' in folder '{folder}'")
                    elif author not in normalized_folder:
                        print(f"[Mismatch] Folder: '{folder}' | File: '{file}' | Author: '{author}'")
                    # Optional: Uncomment to see correct matches
                    # else:
                    #     print(f"[Match] Folder: '{folder}' | File: '{file}' | Author: '{author}'")

# Example usage
if __name__ == "__main__":
    base_directory = r"D:\OOP Mid W7"  # Replace with your actual path
    check_author_in_folder_name(base_directory)
