#Author - Murali.r
#ID - #5

# import docx NOT python-docx
import docx

# create an instance of a word document
doc = docx.Document()

# add a heading of level 0 (largest heading)
header=doc.add_heading('Heading for the document', 0)
# It will add the header to the document

header.alignment=1
# It will set the header allignment in center
# center - 1
# left - 0
# Right - 2
# Justify - 3

doc.save("3.AddHeaderToTheDocument.docx")