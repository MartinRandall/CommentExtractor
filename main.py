from excel import write_to_excel
from extract import get_doc_comments, Paragraph
import os

def process_document(document):
    paragraphs = get_doc_comments(document)
    if len(paragraphs) > 0:
        write_to_excel(document, paragraphs)


for root, dirs, files in os.walk("."):
    for file in files:
        if file.endswith(".docx"):
            doc = os.path.join(root, file)
            print("Processing: " + doc)
            process_document(doc)

# document="test2.docx"  #filepath for the input document
# process_document(document)
