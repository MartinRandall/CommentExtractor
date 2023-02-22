from docx import Document
from lxml import etree
import zipfile
import datetime

class Paragraph:
  def __init__(self, text, comments):
    self.text = text
    self.comments = comments

class Comment:
  def __init__(self, id, author, date, initials, text):
    self.id = id
    self.author = author
    self.date = date
    self.initials = initials
    self.text = text
    
  def setHighlightText(self, text):
    self.highlightText = text

ooXMLns = {'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

#Function to extract all the comments of document(Same as accepted answer)
#Returns a dictionary with comment id as key and comment string as value
def get_document_comments(docxFileName):
    comments_dict={}
    docxZip = zipfile.ZipFile(docxFileName)
    names = docxZip.namelist()
    if 'word/comments.xml' not in names:
        return comments_dict
    commentsXML = docxZip.read('word/comments.xml')
    et = etree.XML(commentsXML)
    comments = et.xpath('//w:comment',namespaces=ooXMLns)
    for c in comments:
        comment=c.xpath('string(.)',namespaces=ooXMLns)
        comment_id=c.xpath('@w:id',namespaces=ooXMLns)[0]
        comment_initials=c.xpath('@w:initials',namespaces=ooXMLns)[0]
        comment_date = datetime.datetime.strptime(c.xpath('@w:date',namespaces=ooXMLns)[0], "%Y-%m-%dT%H:%M:%SZ")
        comment_author=c.xpath('@w:author',namespaces=ooXMLns)[0]
        comments_dict[comment_id]=Comment(comment_id,comment_author,comment_date,comment_initials,comment)
    return comments_dict
  
def get_comment_text(commentRangeStart, commentRangeEnd):
    commentText = ""
    if commentRangeStart is not None:
        for sibling in commentRangeStart.itersiblings():
            if sibling is commentRangeEnd:
                break
            if sibling.text and sibling.text is not None:
                commentText += sibling.text
    return commentText
  
#Function to fetch all the comments in a paragraph
def paragraph_comments(paragraph,comments_dict):
    comments=[]
    for run in paragraph.runs:
        comment_reference=run._r.xpath("./w:commentReference")
        if comment_reference:
            start = run._r.xpath('preceding-sibling::w:commentRangeStart[1]')
            end = run._r.xpath('preceding-sibling::w:commentRangeEnd[1]')
            commentText = ""
            if len(start) > 0 and len(end) > 0:
                commentText = get_comment_text(start[0], end[0])
            comment_id=comment_reference[0].xpath('@w:id',namespaces=ooXMLns)[0]
            comment=comments_dict[comment_id]
            comment.setHighlightText(commentText)
            comments.append(comment)
    return comments
  
#Function to fetch all comments with their referenced paragraph
def get_doc_comments(docxFileName) -> list[Paragraph]:
    document = Document(docxFileName)
    comments_dict=get_document_comments(docxFileName)
    paragraphs=[]
    for paragraph in document.paragraphs:  
        if comments_dict: 
            comments=paragraph_comments(paragraph,comments_dict)  
            if comments:
                paragraphs.append(Paragraph(paragraph.text, comments))
    return paragraphs
  
