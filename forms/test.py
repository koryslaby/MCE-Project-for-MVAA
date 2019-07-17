from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfform


c = canvas.Canvas('MCE-TEMPLATE-6.25.2019-radio-restricted.pdf')

print(c)


from tabula import read_pdf

df = read_pdf("MCE-TEMPLATE-6.25.2019-radio-restricted.pdf")

print(df)