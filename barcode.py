from reportlab.graphics.barcode import code39, code128, code93
from reportlab.graphics.barcode import eanbc, qr, usps
from reportlab.graphics.shapes import Drawing
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.graphics import renderPDF

#----------------------------------------------------------------------
def createBarCodes():
    """
    Create barcode examples and embed in a PDF
    """
    c = canvas.Canvas("barcodes.pdf", pagesize=letter)

    barcode_value = "1234567890"

    barcode39 = code39.Extended39("CYF010123456789123456P")
    barcode39Std = code39.Standard39("CYF020123456789123456P", barHeight=20, stop=1)
    barcode39Std2 = code39.Standard39("CYF0301234567891234P", barHeight=20, stop=1)
    # code93 also has an Extended and MultiWidth version
    #barcode93 = code93.Standard93("12345678903")

    #barcode128 = code128.Code128("12345678904")
    # the multiwidth barcode appears to be broken
    #barcode128Multi = code128.MultiWidthBarcode(barcode_value)

    #barcode_usps = usps.POSTNET("50158-9999")

    codes = [barcode39, barcode39Std2, barcode39Std]

    x = 1 * mm
    y = 285 * mm
    x1 = 6.4 * mm

    for code in codes:
        code.drawOn(c, x, y)
        y = y - 15 * mm

    # draw the eanbc8 code
    #barcode_eanbc8 = eanbc.Ean8BarcodeWidget("12345678905")
    #bounds = barcode_eanbc8.getBounds()
    #width = bounds[2] - bounds[0]
    #height = bounds[3] - bounds[1]
    #d = Drawing(50, 10)
    #d.add(barcode_eanbc8)
    #renderPDF.draw(d, c, 15, 555)

    # draw the eanbc13 code
    #barcode_eanbc13 = eanbc.Ean13BarcodeWidget("12345678906")
    #bounds = barcode_eanbc13.getBounds()
    #width = bounds[2] - bounds[0]
    #height = bounds[3] - bounds[1]
    #d = Drawing(50, 10)
    #d.add(barcode_eanbc13)
    #renderPDF.draw(d, c, 15, 465)

    # draw a QR code
    #qr_code = qr.QrCodeWidget('www.mousevspython.com')
    #bounds = qr_code.getBounds()
    #width = bounds[2] - bounds[0]
    #height = bounds[3] - bounds[1]
    #d = Drawing(45, 45, transform=[45./width,0,0,45./height,0,0])
    #d.add(qr_code)
    #renderPDF.draw(d, c, 15, 405)

    c.save()

if __name__ == "__main__":
    createBarCodes()
