from svglib.svglib import svg2rlg
from reportlab.graphics import renderPDF

drawing = svg2rlg("prueba.svg")
renderPDF.drawToFile(drawing, "file.pdf")