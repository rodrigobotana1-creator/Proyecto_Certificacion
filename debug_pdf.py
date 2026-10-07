import pdfplumber
import json

RUTA = "data/actas/ACTA_4241.pdf"  # ajustá a tu ruta real

with pdfplumber.open(RUTA) as pdf:
    for i, pagina in enumerate(pdf.pages):
        print(f"\n========== PÁGINA {i+1} ==========")
        tablas = pagina.extract_tables()
        print(f"Cantidad de tablas detectadas: {len(tablas)}")
        for j, tabla in enumerate(tablas):
            print(f"\n--- Tabla {j+1} ({len(tabla)} filas) ---")
            for fila in tabla:
                print(fila)