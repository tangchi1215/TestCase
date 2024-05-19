from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

class DocumentManager:
    @staticmethod
    def create_document():
        doc = Document()
        style = doc.styles['Normal']
        style.font.name = '標楷體'
        style.font.size = Pt(12)
        style._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
        doc._body.clear_content()
        return doc

    @staticmethod
    def save_document(doc, path):
        doc.save(path)
        print("產製完成~~!")
