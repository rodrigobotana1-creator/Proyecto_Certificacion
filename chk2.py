import pymupdf

RUTA = r"C:\Users\Rodrigo\Desktop\LA MANTOVANA S.A. - ACTAS JULIO 2026\LA MANTOVANA Actas a certificar Julio Cert 34.pdf"   # <-- CAMBIAR
PAGINA = 1

doc = pymupdf.open(RUTA)
page = doc.load_page(PAGINA - 1)

y_emp = None
y_calle = None
for w in page.get_text("words"):
    if "empresa" in w[4].lower():
        y_emp = w[3]
        print(f"'Empresa' -> y1 = {w[3]:.1f}")
    if w[4].strip().upper() == "BALLIVIAN":
        y_calle = w[3]
        print(f"'BALLIVIAN' -> y1 = {w[3]:.1f}")

if y_emp is None or y_calle is None:
    print("ERROR: falta alguna")
    doc.close()
    raise SystemExit

margen = 28.35

y_min = min(y_emp, y_calle)
y_max = max(y_emp, y_calle)

y0 = max(0, y_min - margen)
y1 = min(page.rect.height, y_max + margen)

print(f"page.rect.height = {page.rect.height}")
print(f"Recorte PDF: y de {y0:.1f} a {y1:.1f}")

# Guardar el recorte como PNG para verlo
mat = pymupdf.Matrix(2.0, 2.0)
pix = page.get_pixmap(matrix=mat, alpha=False)

from PyQt5.QtGui import QImage
from PyQt5.QtWidgets import QApplication
import sys
app = QApplication.instance() or QApplication(sys.argv)

img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
escala = img.height() / page.rect.height
py0 = int(y0 * escala)
py1 = int(y1 * escala)
recorte = img.copy(0, py0, img.width(), py1 - py0)
recorte.save("chk2_recorte.png")

print("Guardado: chk2_recorte.png")
doc.close()