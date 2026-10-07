import os
import re
from modules.pdf_parser import parsear_pdf_acta


def _normalizar_id(s: str) -> str:
    if not s:
        return ""
    return re.sub(r"[\s\-]+", "_", str(s).strip()).upper()


def _solo_numero_acta(valor) -> str:
    if not valor:
        return ""
    digitos = re.findall(r"\d+", str(valor))
    if not digitos:
        return ""
    return str(int(digitos[-1]))


def escanear_carpeta_pdfs(carpeta: str, progreso=None) -> dict:
    archivos = [f for f in os.listdir(carpeta) if f.lower().endswith(".pdf")]
    archivos.sort()

    indice = {
        "por_id": {},
        "por_care": {},
        "por_came": {},
        "por_nro": {},
        "sin_id": [],
        "errores": [],
    }

    total = len(archivos)
    for i, nombre in enumerate(archivos):
        ruta = os.path.join(carpeta, nombre)
        if progreso:
            progreso(i + 1, total, nombre)
        try:
            actas_pdf = parsear_pdf_acta(ruta)
        except Exception as e:
            indice["errores"].append({"ruta": ruta, "error": str(e)})
            continue

        for datos in actas_pdf:
            acta_id = datos.get("acta_id")
            care = datos.get("care")
            came = datos.get("came")

            if not acta_id:
                indice["sin_id"].append({"ruta": ruta, "datos": datos})
                continue

            id_norm = _normalizar_id(acta_id)
            entry = {
                "ruta": ruta,
                "datos": datos,
                "care": care,
                "came": came,
                "pagina": datos.get("pagina"),
            }
            indice["por_id"][id_norm] = entry

            if care:
                indice["por_care"][str(care).strip()] = id_norm
            if came:
                indice["por_came"][str(came).strip()] = id_norm

            nro = _solo_numero_acta(acta_id)
            if nro:
                indice["por_nro"][nro] = id_norm

    return indice


def matchear_acta_excel(acta_excel: str, fila_excel, indice: dict):
    from modules.excel_reader import obtener_numero_acta, obtener_came, obtener_care

    nro = obtener_numero_acta(acta_excel)
    if nro and nro in indice["por_nro"]:
        return indice["por_id"].get(indice["por_nro"][nro])

    came = obtener_came(fila_excel)
    if came and came in indice["por_came"]:
        return indice["por_id"].get(indice["por_came"][came])

    care = obtener_care(fila_excel)
    if care and care in indice["por_care"]:
        return indice["por_id"].get(indice["por_care"][care])

    return None


def obtener_actas_de_indice(indice):
    resultado = []
    for id_norm, entry in indice["por_id"].items():
        resultado.append({
            "acta_id": entry["datos"].get("acta_id") or id_norm,
            "datos": entry["datos"],
            "ruta": entry["ruta"],
            "pagina": entry.get("pagina"),
        })
    return resultado