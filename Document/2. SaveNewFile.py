#Author - Murali.r
#ID - #4

# The programe is used to open the existing file and save to the new name.

# Import docx NOT python-docx
import docx

# Opening a previously created document
doc = docx.Document('newFile.docx')

# Now save the document to a location
doc.save('newFile001.docx')