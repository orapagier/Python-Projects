import os
from pptx import Presentation

# Set folder path 
folder = r"C:\Users\Admin\Desktop\PPS"

# Loop through all files in the folder
for filename in os.listdir(folder):
  if filename.endswith('.pptx'):
    
    # Load Presentation and save it in .ppt format
    pres = Presentation(os.path.join(folder,filename))
    ppt_name = os.path.splitext(filename)[0] + '.pps'
    pres.save(os.path.join(folder,ppt_name))
    
    print(filename + ' converted successfully!')
