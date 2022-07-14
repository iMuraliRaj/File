# Import docx NOT python-docx
import shutil
from docx.shared import Pt, RGBColor
import docx
from docx.shared import Inches, Cm
doc = docx.Document()

sections = doc.sections
for section in sections:
    section.top_margin = Cm(1.27)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(2.06)
    section.right_margin = Cm(2.54)

para = doc.add_paragraph().add_run('\n\n\n\n\n         Qualis LIMS (ReactJS) 1.0')
# Increasing size of the font
para.font.size = Pt(28)
para.bold=True
para.alignment=1
para.font.name = 'Arial'

para = doc.add_paragraph().add_run('          Installation Qualification\n')
# Increasing size of the font
para.font.size = Pt(28)
para.bold=True
para.alignment=1
para.font.name = 'Arial'


para = doc.add_paragraph().add_run('\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n     This information is believed to be complete and accurate as of publication, and is subject to change without notice.')
# Increasing size of the font
para.font.size = Pt(9)
para.alignment=1
para.font.name = 'Arial'

para = doc.add_paragraph().add_run('                                  Copyright © 2000—2010 Agaram Technologies. All rights reserved.\n')
# Increasing size of the font
para.font.size = Pt(9)
para.font.name = 'Arial'

doc.add_page_break()

# Adding a paragraph
contentsHeading=doc.add_paragraph().add_run('Contents')
contentsHeading.font.size = Pt(14)
contentsHeading.bold=True
contentsHeading.font.color.rgb = RGBColor(6, 4, 255)
contentsHeading.font.name = 'Cambria'

purpose=doc.add_paragraph('  Purpose\n',  style='List Number')
doc.add_paragraph('  Purpose',  style='List Number 2')
doc.add_paragraph('  Scope\n',  style='List Number')
doc.add_paragraph('  Validation Methodology\n',  style='List Number')
doc.add_paragraph('  Acronyms\n',  style='List Number')
doc.add_paragraph('  System Description\n',  style='List Number')
doc.add_paragraph('  Introduction\n',  style='List Number 2')
doc.add_paragraph('  Work Flow\n',  style='List Number 2')
doc.add_paragraph('  Responsibilities\n',  style='List Number')
doc.add_paragraph('  Test Plan\n',  style='List Number')
doc.add_paragraph('  Prerequisites Review\n',  style='List Number 2')
doc.add_paragraph('  Computer System Specification and Power Supply\n',  style='List Number 2')
doc.add_paragraph('  Software Review\n',  style='List Number 2')
doc.add_paragraph('  Qualis Installation Verification\n',  style='List Number 2')
doc.add_paragraph('  Verifying components of the Qualis\n',  style='List Number 2')
doc.add_paragraph('  Qualification Support Environment screen shots\n',  style='List Number')
doc.add_paragraph('  Deficiency and Change Request Log\n',  style='List Number')
doc.add_paragraph('  Document Approval',  style='List Number')

doc.add_page_break()
section = doc.sections[0]
header = section.header
header_para = header.paragraphs[0]

# Adding the left zoned footer
header_para.text ="\tINSTALLATION QUALIFICATION"

table = header.add_table(1, 3, Inches(10))
row = table.add_row().cells
row[0].text = "INSTALL"


def headerParagraph(content):
    # Adding a paragraph
    contentsHeading = doc.add_paragraph().add_run(content)
    contentsHeading.font.size = Pt(14)
    contentsHeading.bold = True
    contentsHeading.font.color.rgb = RGBColor(6, 4, 255)
    contentsHeading.font.name = 'Cambria'



source='D:\\iMuraliRaj\\GitHub\\File\\IQ\\Installation Qulification.docx'
doc.save(source)
dest = shutil.copyfile(source, "C:\\Users\\Murali.R\\Desktop\\IQ.docx")