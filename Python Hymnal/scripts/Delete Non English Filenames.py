import os
import shutil
from langdetect import detect

# Specify the directory to search in
search_directory = r"C:\Users\Admin\Desktop\With Audio"

# Specify the backup directory for non-English files
backup_directory = r"C:\Users\Admin\Desktop\asd"

def is_english(filename):
    try:
        language = detect(filename)
        return language == 'en'
    except:
        # Handle exceptions, e.g., when the filename is too short for detection
        return False

def process_files(directory):
    for root, dirs, files in os.walk(directory):
        for file in files:
            # Check if the filename is likely English
            if not is_english(file):
                file_path = os.path.join(root, file)
                # Move the non-English filename to the backup directory
                shutil.move(file_path, os.path.join(backup_directory, file))

if __name__ == "__main__":
    if not os.path.exists(backup_directory):
        os.makedirs(backup_directory)
    
    process_files(search_directory)
    print("Files checked, and non-English files moved to the backup directory.")
