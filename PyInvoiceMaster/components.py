from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, Table, TableStyle
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib import colors


class CodeSnippet(Paragraph):
    style = ParagraphStyle(
        name='CodeSnippet',
        parent=getSampleStyleSheet()['Code'],
        backColor=colors.lightgrey, leftIndent=0,
        borderPadding=(5, 5, 5, 5)
    )

    def __init__(self, code):
        Paragraph.__init__(self, code, self.style)


class SimpleTable(Table):
    def __init__(self, data, horizontal_align=None):
        Table.__init__(self, data, hAlign=horizontal_align)


class TableWithHeader(Table):
    def __init__(self, data, horizontal_align=None, style=None):
        Table.__init__(self, data, hAlign=horizontal_align)

        default_style = [
            ('INNERGRID', (0, 0), (-1, -1), .25, colors.black),
            ('BOX', (0, 0), (-1, -1), .25, colors.black),
            ('BACKGROUND', (0, 0), (-1, -len(data)), colors.lightgrey),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')
        ]

        if style and isinstance(style, list):
            default_style.extend(style)

        self.setStyle(TableStyle(default_style))


class PaidStamp(object):
    def __init__(self, x, y):
        self.x = x
        self.y = y

    def __call__(self, canvas, doc):
        # "PAID"
        canvas.saveState()
        # canvas.setFontSize(50)
        # canvas.setFillColor(colors.red)
        # canvas.setStrokeColor(colors.red)
        # canvas.rotate(45)
        # canvas.drawString(self.x, self.y, 'PAID')
        # canvas.setLineWidth(4)
        # canvas.setLineJoin(1)  # Round join
        # canvas.rect(self.x - .25 * inch, self.y - .25 *
        #             inch, width=2*inch, height=inch)
        canvas.drawInlineImage("/Users/jongregis/Python/JobAutomation/PyInvoiceMaster/FINAL NEW LOGO FOR ECA.png",
                               self.x-5.75 * inch, self.y + 3.5 * inch, width=2*inch, height=.75*inch)
        canvas.restoreState()


class PaidStampCoolCreek(object):
    def __init__(self, x, y):
        self.x = x
        self.y = y

    def __call__(self, canvas, doc):
        # "PAID"
        canvas.saveState()
        # canvas.setFontSize(50)
        # canvas.setFillColor(colors.red)
        # canvas.setStrokeColor(colors.red)
        # canvas.rotate(45)
        # canvas.drawString(self.x, self.y, 'PAID')
        # canvas.setLineWidth(4)
        # canvas.setLineJoin(1)  # Round join
        # canvas.rect(self.x - .25 * inch, self.y - .25 *
        #             inch, width=2*inch, height=inch)
        canvas.drawInlineImage("/Users/jongregis/Python/JobAutomation/PyInvoiceMaster/771 Cool Creek Rd LLC Logo.png",
                               self.x-5.75 * inch, self.y + 3 * inch, width=2*inch, height=2*inch)
        canvas.restoreState()


class PaidStampPSF(object):
    def __init__(self, x, y):
        self.x = x
        self.y = y

    def __call__(self, canvas, doc):
        # "PAID"
        canvas.saveState()
        # canvas.setFontSize(50)
        # canvas.setFillColor(colors.red)
        # canvas.setStrokeColor(colors.red)
        # canvas.rotate(45)
        # canvas.drawString(self.x, self.y, 'PAID')
        # canvas.setLineWidth(4)
        # canvas.setLineJoin(1)  # Round join
        # canvas.rect(self.x - .25 * inch, self.y - .25 *
        #             inch, width=2*inch, height=inch)
        canvas.drawInlineImage("/Users/jongregis/Python/JobAutomation/PyInvoiceMaster/PSF LOGO 2019.png",
                               self.x-5.75 * inch, self.y + 3 * inch, width=2*inch, height=2*inch)
        canvas.restoreState()
