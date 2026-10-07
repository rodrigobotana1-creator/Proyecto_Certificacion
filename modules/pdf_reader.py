import pdfplumber
import re


# Mapeo de índice de columna -> nombre de operación
# Sacado de la fila de header de la Tabla 3 del acta
COLUMNAS = {
    6:  "1.4.3",
    8:  "2.1",
    11: "2.6",
    13: "3.3.1",
    15: "3.3.2",
    17: "4.1.1",
    19: "4.1.2",
}

# Columnas "extra" que no son operaciones pero nos interesan
COL_LARGO = 2
COL_ANCHO = 3
COL_H = 4
COL_TOTAL = 22


def _a_float(s):
    """Convierte '12,00' o '12.00' o '' en float o None."""
    if s is None:
        return None
    s = str(s).strip()
    if s == "":
        return None
    s = s.replace(".", "").replace(",", ".")  # 1.234,56 -> 1234.56 ; 12,00 -> 12.00
    try:
        return float(s)
    except ValueError:
        return None


def parsear_pdf_acta(ruta_pdf: str) -> dict:
    """
    Devuelve:
    {
      "acta_id": "P6_2_Z01_R9_CERT_36_ACTA_4241",
      "came": "6638420",
      "filas": [
         {"nro": 1, "largo": 3.0, "ancho": 4.0, "valores": {"1.4.3": 12.0, "2.1": 3.0, ...}, "total": 12.0},
         ...
      ],
      "totales": {"1.4.3": 55.5, "2.1": 13.88, ...}
    }
    """
    resultado = {"acta_id": None, "came": None, "filas": [], "totales": {}}

    with pdfplumber.open(ruta_pdf) as pdf:
        for pagina in pdf.pages:
            tablas = pagina.extract_tables()
            for tabla in tablas:
                for fila in tabla:
                    # --- Detectar acta y CAME ---
                    for celda in fila:
                        if not celda:
                            continue
                        if "ACTA DE MEDICION" in celda and resultado["acta_id"] is None:
                            m = re.search(r"ACTA[_\s]*(\d+)", celda, re.IGNORECASE)
                            if m:
                                # reconstruyo el ID completo desde el texto
                                m2 = re.search(r"(P6[_\s]*2[_\s]*Z0?1[_\s]*R9[_\s]*CERT[_\s]*36[_\s]*ACTA[_\s]*\d+)", celda, re.IGNORECASE)
                                if m2:
                                    resultado["acta_id"] = re.sub(r"\s+", "_", m2.group(1))
                        if "CAME" in celda and resultado["came"] is None:
                            m = re.search(r"(\d{5,})", celda)
                            if m:
                                resultado["came"] = m.group(1)

                # --- Detectar filas de datos (las que empiezan con '1', '2', ...) ---
                for fila in tabla:
                    if len(fila) <= COL_TOTAL:
                        continue
                    nro = fila[1]
                    if nro is None:
                        continue
                    nro_str = str(nro).strip()
                    if not re.fullmatch(r"\d{1,2}", nro_str):
                        continue
                    # Fila TOTAL
                    if nro_str.upper() == "TOTAL":
                        continue
                    # Es fila de datos
                    valores = {}
                    for idx, nombre in COLUMNAS.items():
                        valores[nombre] = _a_float(fila[idx])
                    fila_data = {
                        "nro": int(nro_str),
                        "largo": _a_float(fila[COL_LARGO]),
                        "ancho": _a_float(fila[COL_ANCHO]),
                        "valores": valores,
                        "total": _a_float(fila[COL_TOTAL]),
                    }
                    # Solo agrego si tiene algún valor
                    if any(v is not None for v in valores.values()):
                        resultado["filas"].append(fila_data)

                # --- Fila TOTAL ---
                for fila in tabla:
                    if len(fila) <= COL_TOTAL:
                        continue
                    if fila[1] and str(fila[1]).strip().upper() == "TOTAL":
                        for idx, nombre in COLUMNAS.items():
                            resultado["totales"][nombre] = _a_float(fila[idx])
                        resultado["totales"]["TOTAL"] = _a_float(fila[COL_TOTAL])

    return resultado