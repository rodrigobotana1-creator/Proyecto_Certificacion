import pymupdf

RUTA = r"C:\Users\Rodrigo\Desktop\LA MANTOVANA S.A. - ACTAS JULIO 2026\LA MANTOVANA Actas a certificar Julio Cert 34.pdf"   # <-- CAMBIAR

doc = pymupdf.open(RUTA)
print("Páginas:", doc.page_count)
print()
page = doc.load_page(0)

print("=== Palabras que contienen 'mpresa' ===")
for w in page.get_text("words"):
    if "mpresa" in w[4]:
        print("  ", w[4], "y0=", round(w[1],1), "y1=", round(w[3],1))

print()
print("=== Palabras que contienen 'BALLIV' ===")
for w in page.get_text("words"):
    if "BALLIV" in w[4].upper():
        print("  ", w[4], "y0=", round(w[1],1), "y1=", round(w[3],1))

print()
print("=== Palabras en la zona del croquis (y entre 100 y 400) ===")
for w in page.get_text("words"):
    if 100 < w[1] < 400:
        print("  ", w[4], "y=", round(w[1],1))

doc.close()