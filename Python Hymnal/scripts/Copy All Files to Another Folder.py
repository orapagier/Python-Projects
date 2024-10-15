import os
import shutil

# Specify the source directory (the folder containing files you want to copy)
source_directory = r"C:\Users\Admin\Desktop\With Audio"

# Specify the destination directory (where you want to copy the files)
destination_directory = r"C:\Users\Admin\Desktop\With Audio English"

def copy_files(src_dir, dst_dir):
    for root, _, files in os.walk(src_dir):
        for file in files:
            src_file = os.path.join(root, file)
            # Copy the file to the destination directory
            shutil.copy2(src_file, dst_dir)

if __name__ == "__main__":
    copy_files(source_directory, destination_directory)
    print("Files copied to the destination directory.")
