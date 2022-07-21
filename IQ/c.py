from docx import Document
f = 'E:\\iphone\\IQ.docx'
document = Document(f)
pn=1
import re
for p in document.paragraphs:
    r=re.match('abc+',p.text)
    if r:
        print(r.group(),pn)
    for run in p.runs:
        if 'w:br' in run._element.xml and 'type="page"' in run._element.xml:
            pn+=1
            print('!!','='*50,pn)