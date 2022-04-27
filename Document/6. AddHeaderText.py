#adding word in footer
from docx import  Document

import Utility

document = Document()
section = document.sections[0]
footer = section.header
footer_para = footer.paragraphs[0]

# Adding the left zoned footer
footer_para.text ="Header Name\n Header second line\t\tHeader right page"

document.save(Utility.projectDirectory()+"Header.docx")


