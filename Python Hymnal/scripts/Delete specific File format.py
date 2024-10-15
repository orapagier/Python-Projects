import os

# Define the folder path where you want to delete .ppt files
folder_path = r"D:\Jelmar Orapa\NEMA HYMNS\nemahymns v1.5\Data\3 SDA HYMNAL"

# Iterate through all files in the folder
for filename in os.listdir(folder_path):
    file_path = os.path.join(folder_path, filename)

    # Check if it's a file and ends with ".ppt"
    if os.path.isfile(file_path) and filename.lower().endswith(".pptx"):
        try:
            # Delete the .ppt file
            os.remove(file_path)
            print(f"Deleted: {filename}")
        except Exception as e:
            print(f"Error deleting {filename}: {e}")
    else:
        # Keep other file formats
        print(f"Kept: {filename}")

print("Deletion process completed.")
