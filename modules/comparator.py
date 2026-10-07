import re as _re
import unicodedata


_ABREVIATURAS = {
    "PJE": "PASAJE", "PJE.": "PASAJE",
    "BLVD": "BOULEVARD", "BLVD.": "BOULEVARD",
    "CNEL": "CORONEL", "CNEL.": "CORONEL",
    "TTE": "TENIENTE", "TTE.": "TENIENTE",
    "SGT": "SARGENTO", "SGT.": "SARGENTO",
    "CDTE": "COMANDANTE", "CDTE.": "COMANDANTE",
    "DR": "DOCTOR", "DR.": "DOCTOR",
    "ING": "INGENIERO", "ING.": "INGENIERO",
    "STA": "SANTA", "STA.": "SANTA",
    "STO": "SANTO", "STO.": "SANTO",
    "SN": "SAN", "SN.": "SAN",
}


def _quitar_tildes(s: str) -> str:
    if not s:
        return ""
    return "".join(
        c for c in unicodedata.normalize("NFD", str(s))
        if unicodedata.category(c) != "Mn"
    )


def comparar(df_excel_acta, datos_pdf, tolerancia=0.05):
    from modules.excel_reader import obtener_operacion, obtener_valor

    grupos_excel = {}
    orden_ops = []
    for _, fila in df_excel_acta.iterrows():
        op = obtener_operacion(fila)
        if not op:
            continue
        if op not in grupos_excel:
            grupos_excel[op] = []
            orden_ops.append(op)
        grupos_excel[op].append(obtener_valor(fila))

    grupos_pdf = {}
    for it in datos_pdf.get("items", []):
        op = it["item"]
        if op not in grupos_pdf:
            grupos_pdf[op] = []
        grupos_pdf[op].append(it["valor"])

    resultados = []
    for op in orden_ops:
        valores_excel = grupos_excel[op]
        valores_pdf = grupos_pdf.get(op, [])

        filas_cmp = []
        n = max(len(valores_excel), len(valores_pdf))
        for i in range(n):
            v_ex = valores_excel[i] if i < len(valores_excel) else None
            v_pdf = valores_pdf[i] if i < len(valores_pdf) else None
            ok = (
                v_ex is not None and v_pdf is not None
                and round(abs(v_ex - v_pdf), 5) <= tolerancia
            )
            filas_cmp.append({
                "nro": i + 1,
                "excel": v_ex,
                "pdf": v_pdf,
                "ok": ok,
                "dif": (round(v_ex - v_pdf, 3) if v_ex is not None and v_pdf is not None else None),
            })

        suma_excel = round(sum(v for v in valores_excel if v is not None), 3)
        suma_pdf = round(sum(v for v in valores_pdf if v is not None), 3)
        ok_total = (
            round(abs(suma_excel - suma_pdf), 5) <= tolerancia
            if valores_excel or valores_pdf else False
        )

        resultados.append({
            "operacion": op,
            "filas": filas_cmp,
            "suma_excel": suma_excel,
            "total_pdf": suma_pdf,
            "ok_total": ok_total,
            "dif_total": round(suma_excel - suma_pdf, 3),
        })

    return {"por_operacion": resultados}


def _expandir_abreviaturas(s):
    if not s:
        return ""
    palabras = s.split()
    return " ".join(_ABREVIATURAS.get(p, p) for p in palabras)


def _normalizar_calle(s):
    if not s:
        return ""
    s = _quitar_tildes(str(s).strip().upper())
    s = s.replace("\xa0", " ").replace("\t", " ")
    s = s.replace(",", " ")
    s = _re.sub(r"\s+", " ", s).strip()
    s = _expandir_abreviaturas(s)
    palabras_no_calle = {
        "CALZADA", "ACERA", "CALLE",
        "PASAJE", "BOULEVARD",
        "DE", "DEL", "LA", "LAS", "LOS", "EL",
    }
    palabras = [p for p in s.split(" ") if p and p not in palabras_no_calle]
    return " ".join(palabras)


def _parsear_direccion_excel(texto_excel):
    if not texto_excel:
        return None
    t = str(texto_excel).strip()
    t = t.replace("\xa0", " ").replace("\t", " ")
    t = _re.sub(r"\s+", " ", t).strip()

    m = _re.search(
        r"^(.+?)\s+(\d{1,15})\s*-\s*(\d{1,15})(?:\s|$)",
        t,
    )
    if not m:
        return None
    return {
        "calle": _normalizar_calle(m.group(1)),
        "ini": int(m.group(2)),
        "fin": int(m.group(3)),
    }


def _calles_coinciden(calle_a, calle_b):
    if not calle_a or not calle_b:
        return False
    if calle_a == calle_b:
        return True
    palabras_a = set(calle_a.split())
    palabras_b = set(calle_b.split())
    return bool(palabras_a & palabras_b)


def _rango_coincide(sap_ini, sap_fin, acta_ini, acta_fin, tol=2):
    if None in (sap_ini, sap_fin, acta_ini, acta_fin):
        return False
    return (
        round(abs(sap_ini - acta_ini), 5) <= tol
        and round(abs(sap_fin - acta_fin), 5) <= tol
    )


def _alturas_coinciden(sap_ini, sap_fin, alturas_pdf, tol=2):
    if sap_ini is None or sap_fin is None:
        return False
    if not alturas_pdf:
        return False
    for a in alturas_pdf:
        if _rango_coincide(sap_ini, sap_fin, a["ini"], a["fin"], tol=tol):
            return True
    return False


def comparar_direccion(texto_excel, datos_pdf):
    ex = _parsear_direccion_excel(texto_excel)

    calle_pdf = _normalizar_calle(datos_pdf.get("calle"))
    alturas_pdf = datos_pdf.get("alturas", [])

    ok_calle = bool(ex and _calles_coinciden(ex["calle"], calle_pdf))
    ok_alturas = bool(
        ex
        and _alturas_coinciden(ex["ini"], ex["fin"], alturas_pdf, tol=2)
    )

    return {
        "excel": ex,
        "pdf": {
            "calle": calle_pdf,
            "alturas": alturas_pdf,
        },
        "ok_calle": ok_calle,
        "ok_alturas": ok_alturas,
        "ok": ok_calle and ok_alturas,
    }