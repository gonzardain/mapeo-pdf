from reportlab.lib.units import mm
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen.canvas import Canvas
from reportlab.platypus.flowables import KeepInFrame, XBox

canv = Canvas("sample.pdf", pagesize=A4)
x, y, w, h = 10*mm, 10*mm, 100*mm, 50*mm
canv.rect(x, y, w, h)

box = XBox(w + 20*mm, 10*mm, "XBox")
# in next line only "truncate" seems to work as expected
border = KeepInFrame(w, h, box, mode="shrink")

# the following three lines are needed to avoid
# AttributeErrors in KeepInFrame
border.width = w
border.height = h
border._content = [box]

border.drawOn(canv, x, y)

canv.showPage()
canv.save()