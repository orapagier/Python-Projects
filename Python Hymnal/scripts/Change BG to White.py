import os
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor

folder = r"C:\Users\Admin\Desktop\New folder" 

for filename in os.listdir(folder):
  if filename.endswith('.pptx'):

    prs = Presentation(os.path.join(folder, filename))

    for slide in prs.slides:
      slide.background.fill.solid()
      slide.background.fill.fore_color.rgb = RGBColor(255, 255, 255)

    prs.save(filename)
    
print("Background color changed to white")
