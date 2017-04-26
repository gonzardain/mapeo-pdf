from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph
import StringIO
from pyPdf import PdfFileWriter, PdfFileReader

#----------------------------------------------------------------------
def settingParagraphs():
    global packet
    packet = StringIO.StringIO()
    """
    Demo to show how to use fonts in Paragraphs
    """
    p_font = 12
    c = canvas.Canvas(packet, pagesize=letter)
 
    ptext = """<font name=HELVETICA size=%s>Welcome to Reportlab! (helvetica)</font>
    """ % p_font
    createPDF(c, ptext, 20, 750)
 
    ptext = """<font name=courier size=%s>Welcome to Reportlab! (courier)</font>
    """ % p_font
    createPDF(c, ptext, 20, 730)
 
    ptext = """<font name=times-roman size=%s>Welcome to Reportlab! (times-roman)</font>
    """ % p_font
    createPDF(c, ptext, 20, 710)
 
    c.save()
 
#----------------------------------------------------------------------
def createPDF(c, text, x, y):
    """"""
    style = getSampleStyleSheet()
    width, height = letter
    p = Paragraph(text, style=style["Normal"])
    p.wrapOn(c, width, height)
    p.drawOn(c, x, y, mm)

def mergePDF():
    global packet
    packet.seek(0)
    new_pdf = PdfFileReader(packet)
    existing_pdf = PdfFileReader(file("PDFCODELAB.pdf", "rb"))
    output = PdfFileWriter()
    page = existing_pdf.getPage(0)
    page.mergePage(new_pdf.getPage(0))
    output.addPage(page)
    outputStream = file("newpdf.pdf", "wb")
    output.write(outputStream)
    outputStream.close()

 
if __name__ == "__main__":
    settingParagraphs(r"/path/to/barcode.pdf")

