from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle

D = "/usr/share/fonts/truetype/dejavu/"
pdfmetrics.registerFont(TTFont("DV", D + "DejaVuSans.ttf"))
pdfmetrics.registerFont(TTFont("DVB", D + "DejaVuSans-Bold.ttf"))
pdfmetrics.registerFont(TTFont("DVI", D + "DejaVuSans.ttf"))
from reportlab.lib.fonts import addMapping
addMapping("DV", 0, 0, "DV"); addMapping("DV", 1, 0, "DVB"); addMapping("DV", 0, 1, "DVI"); addMapping("DV", 1, 1, "DVB")

LACI = colors.HexColor("#1F3A5F")
ACIK = colors.HexColor("#DCE6F0")
KIRMIZI = colors.HexColor("#B42318")
GRI = colors.HexColor("#555555")
W, H = A4

def stil(ad, size=8.5, font="DV", leading=None, color=colors.black, **kw):
    return ParagraphStyle(ad, fontName=font, fontSize=size, leading=leading or size * 1.25, textColor=color, **kw)

def ust_bilgi(c, baslik, alt, renk=LACI):
    c.setFillColor(renk)
    c.rect(0, H - 22 * mm, W, 22 * mm, stroke=0, fill=1)
    c.setFillColor(colors.white)
    c.setFont("DVB", 16)
    c.drawString(14 * mm, H - 12 * mm, baslik)
    c.setFont("DV", 8.5)
    c.drawString(14 * mm, H - 18 * mm, alt)

def alt_bilgi(c, metin):
    c.setFont("DV", 7)
    c.setFillColor(GRI)
    c.drawString(14 * mm, 8 * mm, metin)
    c.drawRightString(W - 14 * mm, 8 * mm, f"Sayfa {c.getPageNumber()}")
