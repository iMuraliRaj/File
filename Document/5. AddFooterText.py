#adding word in footer
from docx import  Document

import Utility

document = Document()
section = document.sections[0]
footer = section.footer
footer_para = footer.paragraphs[0]

# Adding the left zoned footer
footer_para.text ="Confidential\nAGARAM Technologies \t\tPage"
#\t - Centre alignment
#\t\t - Right alignment

document.save(Utility.projectDirectory()+"footer.docx")


