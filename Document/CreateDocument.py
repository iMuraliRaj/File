# Import docx NOT python-docx
import docx

# Create an instance of a word document
doc = docx.Document()

# Now save the document to a location
doc.save('newFile.docx')