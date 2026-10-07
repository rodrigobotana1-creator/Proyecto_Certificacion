from modules.pdf_parser import parsear_pdf_acta

RUTA = r"data\actas\ACTA_4241.pdf"

datos = parsear_pdf_acta(RUTA)
print("\n===== RESULTADO =====")
print("ACTA:", datos["acta_id"])
print("CAME:", datos["came"])
print("Cantidad filas:", len(datos["filas"]))
for f in datos["filas"]:
    print(f)
print("Totales:", datos["totales"])