import pdfplumber
import re

RUTA = r"data\actas\ACTA_4241.pdf"

with pdfplumber.open(RUTA) as pdf:
    p = pdf.pages[0]
    alto = p.height
    ancho = p.width
    print(f"Página: {ancho:.1f} x {alto:.1f}")
    print(f"30% del alto: {alto * 0.30:.1f}")
    print()
    print("Números de 4-5 dígitos:")
    print(f"{'texto':<10} {'x0':>8} {'x1':>8} {'top':>8} {'pasa 30%?':>10}")
    print("-" * 50)
    for w in p.extract_words():
        t = w["text"].strip()
        if re.fullmatch(r"\d{4,5}", t):
            pasa = w["top"] <= alto * 0.30
            print(f"{t:<10} {w['x0']:>8.1f} {w['x1']:>8.1f} {w['top']:>8.1f} {str(pasa):>10}")

    print()
    print("Palabras NO numéricas con minúsculas, top entre 100 y 260:")
    for w in p.extract_words():
        t = w["text"].strip()
        if not t or any(c.isdigit() for c in t):
            continue
        if not any(c.islower() for c in t):
            continue
        if 100 <= w["top"] <= 260:
            print(f"  {t:<20} x0={w['x0']:.1f} x1={w['x1']:.1f} top={w['top']:.1f}")