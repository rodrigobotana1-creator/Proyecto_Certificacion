from modules.pdf_parser import parsear_pdf_acta
import os

carpeta = r"data\actas"

for nombre in sorted(os.listdir(carpeta)):
    if not nombre.lower().endswith(".pdf"):
        continue
    ruta = os.path.join(carpeta, nombre)
    print("=" * 60)
    print(f"ARCHIVO: {nombre}")
    print("=" * 60)
    datos = parsear_pdf_acta(ruta)
    print(f"  acta_id:    {datos['acta_id']}")
    print(f"  came:       {datos['came']}")
    print(f"  calle:      {datos['calle']!r}")
    print(f"  altura_ini: {datos['altura_ini']}")
    print(f"  altura_fin: {datos['altura_fin']}")
    print(f"  filas:      {len(datos['filas'])}")
    print()