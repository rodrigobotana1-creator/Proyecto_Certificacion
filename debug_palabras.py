import pdfplumber

RUTA = r"C:\Users\Matias Botana\Desktop\comparador_actas\data\actas\ACTA_4241.pdf"

with pdfplumber.open(RUTA) as pdf:
    pagina = pdf.pages[0]
    palabras = pagina.extract_words(use_text_flow=False, keep_blank_chars=False)
    alto = pagina.height
    ancho = pagina.width
    print(f"Página: {ancho:.0f} x {alto:.0f} px")
    print()
    print("Palabras en la zona superior (top < 55% del alto):")
    print(f"{'texto':<25} {'x0':>7} {'x1':>7} {'top':>7} {'bottom':>7}")
    print("-" * 60)
    for p in palabras:
        if p["top"] < alto * 0.55:
            print(f"{p['text']:<25} {p['x0']:>7.1f} {p['x1']:>7.1f} {p['top']:>7.1f} {p['bottom']:>7.1f}")