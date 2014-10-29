# -*- coding: utf-8 -*-
"""
Created on Fri Sep 26 11:47:17 2014

@author: sameer
"""

from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfpage import PDFTextExtractionNotAllowed
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams, LTTextBox, LTTextLine, LTFigure, LTImage


# Open a PDF file.
fp = open('/home/sameer/2008.pdf', 'rb')
# Create a PDF parser object associated with the file object.
parser = PDFParser(fp)
# Create a PDF document object that stores the document structure.
# Supply the password for initialization.
document = PDFDocument(parser)
# Check if the document allows text extraction. If not, abort.
if not document.is_extractable:
    raise PDFTextExtractionNotAllowed
# Create a PDF resource manager object that stores shared resources.
rsrcmgr = PDFResourceManager()
# Create a PDF device object.
laparams = LAParams()
device = PDFPageAggregator(rsrcmgr, laparams=laparams)
# Create a PDF interpreter object.
interpreter = PDFPageInterpreter(rsrcmgr, device)
# Process each page contained in the document.

def parse_lt_objs (lt_objs, text=[]):
    """Iterate through the list of LT* objects and capture the text or image data contained in each"""
    text_content = [] 

    for lt_obj in lt_objs:
        if isinstance(lt_obj, LTTextBox) or isinstance(lt_obj, LTTextLine):
            # text
            text_content.append(lt_obj.get_text())
        elif isinstance(lt_obj, LTImage):
            # an image, so save it to the designated folder, and note it's place in the text 
            print 'lol'
        elif isinstance(lt_obj, LTFigure):
            # LTFigure objects are containers for other LT* objects, so recurse through the children
            text_content.append(parse_lt_objs(lt_obj, text_content))

    return '\n'.join(text_content)

pdflist = []
i = 1
for page in PDFPage.create_pages(document):
    interpreter.process_page(page)
    layout = device.get_result()
    pdflist.append(parse_lt_objs(layout))
    print i,
    i += 1

with open('pilist.txt','w') as pilist:
    pilist.write(str(pdflist))
    
