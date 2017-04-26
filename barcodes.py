from reportlab.graphics.barcode import code128
from reportlab.lib.units import mm, inch
from reportlab.pdfgen import canvas

c = canvas.Canvas("test.pdf")
c.setPageSize((57*mm,32*mm))
#barcode = code128.Code128("123456789",barHeight=.9*inch,barWidth = 1.2)
barcode = code128.Code128("123456789")
barcode.drawOn(c, 2*mm, 20*mm)
c.showPage()
c.save()