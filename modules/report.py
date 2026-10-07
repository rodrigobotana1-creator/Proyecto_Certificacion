import pandas as pd


def generar_reporte(resultados_por_acta, ruta_salida):
    """
    resultados_por_acta: lista de dicts con:
       {
         "acta_id": str,
         "encontrado": bool,
         "filas_ok": int,
         "filas_total": int,
         "totales_ok": int,
         "totales_total": int,
         "detalle": [ {operacion, nro, excel, pdf, dif, ok}, ... ],
         "totales_detalle": [ {operacion, suma_excel, total_pdf, dif, ok}, ... ],
       }
    """
    filas_resumen = []
    filas_detalle = []
    filas_totales = []

    for r in resultados_por_acta:
        estado = "SIN PDF" if not r["encontrado"] else (
            "OK" if r["filas_ok"] == r["filas_total"]
                    and r["totales_ok"] == r["totales_total"]
            else "CON DIFERENCIAS"
        )
        filas_resumen.append({
            "Acta": r["acta_id"],
            "Estado": estado,
            "Filas OK": f"{r['filas_ok']}/{r['filas_total']}",
            "Totales OK": f"{r['totales_ok']}/{r['totales_total']}",
        })

        for d in r.get("detalle", []):
            filas_detalle.append({
                "Acta": r["acta_id"],
                "Operación": d["operacion"],
                "N° fila": d["nro"],
                "Excel": d["excel"],
                "PDF": d["pdf"],
                "Diferencia": d["dif"],
                "OK": "OK" if d["ok"] else "DIF",
            })

        for t in r.get("totales_detalle", []):
            filas_totales.append({
                "Acta": r["acta_id"],
                "Operación": t["operacion"],
                "Suma Excel": t["suma_excel"],
                "Total PDF": t["total_pdf"],
                "Diferencia": t["dif_total"],
                "OK": "OK" if t["ok_total"] else "DIF",
            })

    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
        pd.DataFrame(filas_resumen).to_excel(writer, sheet_name="Resumen", index=False)
        pd.DataFrame(filas_detalle).to_excel(writer, sheet_name="Detalle por fila", index=False)
        pd.DataFrame(filas_totales).to_excel(writer, sheet_name="Totales", index=False)