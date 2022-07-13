#Author - Murali.r
#ID - #5

# import docx NOT python-docx
import docx

# create an instance of a word document
from docx import Document

doc = Document()

# add a heading of level 0 (largest heading)
header=doc.add_heading("Table of Content", 0)
# It will add the header to the document

header.alignment=2

# It will set the header allignment in center
# center - 1
# left - 0
# Right - 2
# Justify - 3

doc.save("E:\\docx\\file.docx")