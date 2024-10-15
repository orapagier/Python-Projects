import os
import re

pptx_folder = r"C:\Users\Admin\Desktop\White BG 1"
txt_file = r"C:\Users\Admin\Desktop\Hymn Titles.txt"

with open(txt_file) as f:
    titles = f.readlines()
titles = [t.strip() for t in titles]

pptx_files = os.listdir(pptx_folder)

for fname in pptx_files:
    num = re.search(r'^\d+', fname).group()

    for title in titles:
        if title.startswith(num):
            title = re.sub(r'[^a-zA-Z0-9\s]+', '', title[len(num):])

            # Remove dashes from title
            title = title.replace("-", "")

            new_name = num + title + '.pptx'  # Keep the same numbering

            # Avoid name conflicts
            if os.path.exists(os.path.join(pptx_folder, new_name)):
                continue

            os.rename(os.path.join(pptx_folder, fname), os.path.join(pptx_folder, new_name))
