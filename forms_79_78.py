from pyPdf import PdfFileWriter, PdfFileReader
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm, cm, inch
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.platypus import Paragraph
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.graphics.barcode import code128, code39
from reportlab.platypus import Paragraph, Frame, KeepInFrame
import cStringIO
from openpyxl import load_workbook
from reportlab.lib.enums import TA_RIGHT
import math, os, checksum

##################################################
#		Ubicacion de las fuentes                 #
##################################################

base = os.path.dirname(os.path.abspath(__file__))
fonts_path = os.path.join(base, "fonts/")

KeepCalmTTF = "KeepCalm-Medium.ttf"
AvenirRegularTTF = "AvenirNextLTPro-Regular.ttf"
AvenirHeavyTTF = "AvenirNextLTPro-HeavyCn.ttf"
AvenirBoldTTF = "AvenirNextLTPro-Bold.ttf"
AvenirItaTTF = "AvenirNextLTPro-It.ttf"

font_KeepCalm = fonts_path + KeepCalmTTF
pdfmetrics.registerFont(TTFont("Keep Calm", font_KeepCalm))

font_AvenirRegular = fonts_path + AvenirRegularTTF
pdfmetrics.registerFont(TTFont("Avenir Regular", font_AvenirRegular))

font_AvenirHeavy = fonts_path + AvenirHeavyTTF
pdfmetrics.registerFont(TTFont("Avenir Heavy", font_AvenirHeavy))

font_AvenirBold = fonts_path + AvenirBoldTTF
pdfmetrics.registerFont(TTFont("Avenir Bold", font_AvenirBold))

font_AvenirItalic = fonts_path + AvenirItaTTF
pdfmetrics.registerFont(TTFont("Avenir Italic", font_AvenirItalic))

def load_wb():
	data_dict = []
	wb = load_workbook('data.xlsx', data_only=False)
	sheet_ranges = wb[wb.get_sheet_names()[0]]
	ws = wb.active
	row_count = ws.max_row
	column_count = ws.max_column
	for row in range(2, row_count+1):
		for column in range (1, column_count+1):
			data = ws.cell(column=column, row=row)
			data = data.value
			create_array(data, column, data_dict)

def create_array(data, column, data_dict):
	data_dict.append(data)
	if column == 36:
		do_stuff(data_dict)


def do_stuff(data_dict):
    #print data_dict
	stringBuffer = cStringIO.StringIO()
	writePage1(stringBuffer, data_dict)
	del data_dict[:]

if __name__ == "__main__":
	global pdf_counter, pdf_final, c
	c = 0
	pdf_final= 1
	pdf_counter = 1
    #load_wb()
