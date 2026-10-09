import fitz  # pymupdf

# CAMBIÁ ESTA RUTA por el PDF real de una de las actas
RUTA_PDF = r"C:\Users\Rodrigo\Desktop\LA MANTOVANA S.A. - ACTAS JULIO 2026\LA MANTOVANA Actas a certificar Julio Cert 34.pdf"


doc = fitz.open(RUTA_PDF)
page = doc.load_page(0)

print("=" * 70)
print("PALABRAS EN MAYÚSCULA (candidatas a cartel)")
print("=" * 70)
palabras = page.get_text("words")
lineas = {}
for w in palabras:
    x0, y0, x1, y1, texto, *_ = w
    y_key = round(y0 / 3.0)
    lineas.setdefault(y_key, []).append((x0, y0, x1, y1, texto))

for y_key, ws in sorted(lineas.items()):
    ws_sorted = sorted(ws, key=lambda t: t[0])
    texto_linea = " ".join(t[4] for t in ws_sorted).strip()
    if not texto_linea:
        continue
    y_prom = sum(t[1] for t in ws_sorted) / len(ws_sorted)
    x_min = min(t[0] for t in ws_sorted)
    x_max = max(t[2] for t in ws_sorted)
    print(f"y={y_prom:7.1f}  x=[{x_min:6.1f}-{x_max:6.1f}]  '{texto_linea}'")

print()
print("=" * 70)
print("DRAWINGS (rectángulos / líneas dibujadas)")
print("=" * 70)
dibujos = page.get_drawings()
print(f"Total: {len(dibujos)}")
print()
print("Primeros 40 rects:")
for i, d in enumerate(dibujos[:40]):
    r = d.get("rect")
    if r:
        print(f"[{i:3d}] ({r.x0:7.1f}, {r.y0:7.1f}) -> ({r.x1:7.1f}, {r.y1:7.1f})   "
              f"ancho={r.x1-r.x0:7.1f}  alto={r.y1-r.y0:7.1f}")

doc.close()