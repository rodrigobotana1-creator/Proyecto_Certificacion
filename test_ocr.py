"""
Test de OCR v5 - con upscale, OEM 3 y variantes.
Soporta PDF, JPG, JPEG, PNG.
"""
import os
import re
import sys
import io

try:
    import pytesseract
    from PIL import Image, ImageOps
    import pymupdf as fitz
except ImportError as e:
    print("FALTA UNA DEPENDENCIA:", e)
    print("Instalá con:  pip install pytesseract pillow PyMuPDF")
    sys.exit(1)


# ===== CONFIGURACIÓN =====
RUTA = r"data\actas_test\foto_acta.jpeg"

TESSERACT_EXE = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
IDIOMA_PREFERIDO = "spa"
DPI = 300


# --- Configurar Tesseract ---
if os.path.exists(TESSERACT_EXE):
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_EXE
else:
    print(f"AVISO: no encontré Tesseract en {TESSERACT_EXE}")
    print("Usando el del PATH. Si falla, corregí TESSERACT_EXE.")


# ===== FUNCIONES =====
def cargar_imagen(ruta):
    """Acepta PDF o imagen (JPG/JPEG/PNG). Devuelve un PIL.Image."""
    ext = os.path.splitext(ruta)[1].lower()
    if ext == ".pdf":
        doc = fitz.open(ruta)
        pagina = doc[0]
        pix = pagina.get_pixmap(dpi=DPI)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        doc.close()
        return img
    else:
        return Image.open(ruta)


def upscale(img, factor):
    w, h = img.size
    return img.resize((w * factor, h * factor), Image.LANCZOS)


def variante_original(img):
    return img


def variante_gris(img):
    return img.convert("L")


def variante_gris_x2(img):
    return upscale(img.convert("L"), 2)


def variante_gris_x3(img):
    return upscale(img.convert("L"), 3)


def variante_gris_x4(img):
    return upscale(img.convert("L"), 4)


def variante_contraste_x2(img):
    g = img.convert("L")
    g = ImageOps.autocontrast(g, cutoff=1)
    return upscale(g, 2)


VARIANTES = [
    ("original",       variante_original),
    ("grises",         variante_gris),
    ("grises_x2",      variante_gris_x2),
    ("grises_x3",      variante_gris_x3),
    ("grises_x4",      variante_gris_x4),
    ("contraste_x2",   variante_contraste_x2),
]


def ocr_imagen(img, idioma, psm=3):
    config = f"--oem 3 --psm {psm}"
    try:
        return pytesseract.image_to_string(img, lang=idioma, config=config)
    except pytesseract.TesseractError as e:
        print(f"    ERROR Tesseract: {e}")
        return ""


def main():
    # --- Verificar Tesseract ---
    print("=" * 60)
    print("1) Verificando Tesseract")
    print("=" * 60)
    try:
        version = pytesseract.get_tesseract_version()
        print(f"  Tesseract versión: {version}")
    except Exception as e:
        print(f"  ERROR al ejecutar Tesseract: {e}")
        sys.exit(1)

    try:
        idiomas = pytesseract.get_languages(config="")
        print(f"  Idiomas disponibles: {idiomas}")
        idioma = IDIOMA_PREFERIDO if IDIOMA_PREFERIDO in idiomas else "eng"
    except Exception as e:
        print(f"  AVISO: no pude listar idiomas ({e}). Sigo con 'eng'.")
        idioma = "eng"

    print(f"  Idioma usado: {idioma}")

    # --- Verificar archivo ---
    print()
    print("=" * 60)
    print("2) Cargando imagen")
    print("=" * 60)

    if not os.path.exists(RUTA):
        print(f"  ERROR: no existe el archivo {RUTA}")
        sys.exit(1)

    img_original = cargar_imagen(RUTA)
    print(f"  Archivo: {RUTA}")
    print(f"  Tamaño original: {img_original.size[0]}x{img_original.size[1]} px")
    print(f"  Modo: {img_original.mode}")

    # --- Probar cada variante ---
    print()
    print("=" * 60)
    print("3) Probando variantes y PSM")
    print("=" * 60)

    resultados = {}
    for nombre_var, fn_var in VARIANTES:
        try:
            img = fn_var(img_original)
        except Exception as e:
            print(f"\n  Variante {nombre_var}: ERROR al procesar: {e}")
            continue

        img.save(f"test_v5_{nombre_var}.png")

        print(f"\n  ===== Variante: {nombre_var} ({img.size[0]}x{img.size[1]}) =====")
        for psm in [3, 4, 6, 11, 12]:
            texto = ocr_imagen(img, idioma, psm=psm)
            n = len(texto.strip())
            print(f"    psm={psm:2d} → {n} caracteres")
            if n > 0:
                resultados[(nombre_var, psm)] = texto

    # --- Mejor resultado ---
    print()
    print("=" * 60)
    print("4) Mejor resultado")
    print("=" * 60)

    if not resultados:
        print("  ¡Ninguna variante sacó texto!")
        return

    (mejor_var, mejor_psm), mejor_texto = max(
        resultados.items(), key=lambda kv: len(kv[1])
    )
    print(f"  Variante: {mejor_var}")
    print(f"  PSM: {mejor_psm}")
    print(f"  Caracteres: {len(mejor_texto)}")
    print()
    print("  --- Texto extraído (primeros 4000 chars) ---")
    print(mejor_texto[:4000])

    # --- Guardar ---
    with open("ocr_salida.txt", "w", encoding="utf-8") as f:
        f.write(mejor_texto)
    print(f"\n  Guardado en ocr_salida.txt ({len(mejor_texto)} chars)")

    # --- Patrones clave ---
    print()
    print("=" * 60)
    print("5) Búsqueda de patrones clave")
    print("=" * 60)

    m = re.search(r"ACTA\s*DE\s*MEDICI[OÓ]N[^\d]*(\d+)", mejor_texto, re.IGNORECASE)
    print(f"  Acta: {m.group(0) if m else 'NO detectada'}")

    m = re.search(r"CAME[^\d]*(\d{5,})", mejor_texto, re.IGNORECASE)
    print(f"  CAME: {m.group(1) if m else 'NO detectado'}")

    # Números detectados
    print()
    print("  Números detectados (primeros 60):")
    numeros = re.findall(r"\d+(?:[.,]\d+)?", mejor_texto)
    print(f"  {numeros[:60]}")

    # Líneas que parecen filas
    print()
    print("  Líneas que parecen filas:")
    n = 0
    for linea in mejor_texto.splitlines():
        if re.match(r"^\s*\d{1,2}\s+\d", linea):
            print("   →", linea)
            n += 1
            if n >= 10:
                break
    if n == 0:
        print("   (ninguna)")


if __name__ == "__main__":
    main()