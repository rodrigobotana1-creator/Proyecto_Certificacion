import pdfplumber
import re


_PALABRAS_VACIAS = {
    "DE", "DEL", "LA", "EL", "LOS", "LAS", "Y", "A", "AL",
    "UN", "UNA", "UNOS", "UNAS", "EN", "CON", "POR", "PARA",
}


def _a_float(s):
    if s is None:
        return None
    s = str(s).strip()
    if s == "":
        return None
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def _extraer_acta_id(texto):
    if not texto:
        return None
    texto_norm = texto.upper().replace("Ó", "O").replace("Í", "I")
    idx = texto_norm.find("ACTA")
    if idx == -1:
        return None
    resto = texto[idx:]
    m = re.search(r"ACTA\D{0,40}?(\d{3,5})", resto, re.IGNORECASE)
    if m:
        return m.group(1)
    return None


def _extraer_care_came(texto):
    if not texto:
        return None, None

    texto_norm = (texto.upper()
                  .replace("Ó", "O").replace("Í", "I")
                  .replace("Á", "A").replace("É", "E"))

    m = re.search(
        r"CARE\s*/\s*CAME\s*[:\s]*(\d{4,10})\s*/\s*(\d{4,10})",
        texto_norm,
    )
    if m:
        return m.group(1), m.group(2)

    m = re.search(
        r"CAME\s*/\s*CARE\s*[:\s]*(\d{4,10})\s*/\s*(\d{4,10})",
        texto_norm,
    )
    if m:
        return m.group(2), m.group(1)

    care = None
    came = None
    m_care = re.search(r"CARE\s*[:\s]+(\d{4,10})", texto_norm)
    if m_care:
        care = m_care.group(1)
    m_came = re.search(r"CAME\s*[:\s]+(\d{4,10})", texto_norm)
    if m_came:
        came = m_came.group(1)

    return care, came


def _top_de_la_tabla(pagina):
    palabras = pagina.extract_words(use_text_flow=False, keep_blank_chars=False)
    if not palabras:
        return pagina.height * 0.5
    claves = {"UBICACIÓN", "UBICACION", "LARGO", "ANCHO", "ESPESOR",
              "BASE", "MAYOR", "MENOR", "OBSERVACIONES"}
    top_min = None
    for p in palabras:
        txt = p["text"].strip().upper()
        if txt in claves:
            if top_min is None or p["top"] < top_min:
                top_min = p["top"]
    return top_min if top_min is not None else pagina.height * 0.5


def _es_altura(txt):
    return bool(re.fullmatch(r"\d{3,5}", txt))


def _es_altura_corta(txt):
    return bool(re.fullmatch(r"\d{1,2}", txt))


def _es_solo_numeros(txt):
    return bool(re.fullmatch(r"[\d\s\.,\-]+", txt))


def _contar_letras(txt):
    return len(re.sub(r"[^A-Za-zÁÉÍÓÚÑáéíóúñ]", "", txt))


def _es_palabra_vacia(txt):
    # DESACTIVADO: queremos guardar la calle tal cual el PDF
    return False


def _es_titulo_tabla(txt):
    # DESACTIVADO: queremos guardar la calle tal cual el PDF
    return False


def _agrupar_por_top(items, tolerancia=20):
    if not items:
        return []
    items = sorted(items, key=lambda i: i["top"])
    grupos = [[items[0]]]
    for it in items[1:]:
        ult = grupos[-1]
        top_prom = sum(x["top"] for x in ult) / len(ult)
        if abs(it["top"] - top_prom) < tolerancia:
            ult.append(it)
        else:
            grupos.append([it])
    return grupos


def _texto_dentro_de_rect(rect, palabras):
    x0, x1 = rect["x0"], rect["x1"]
    top, bottom = rect["top"], rect["bottom"]
    adentro = []
    for p in palabras:
        xc = (p["x0"] + p["x1"]) / 2
        yc = (p["top"] + p["bottom"]) / 2
        if x0 <= xc <= x1 and top <= yc <= bottom:
            adentro.append(p)
    adentro.sort(key=lambda p: p["x0"])
    return adentro


def _buscar_alturas_normales(pagina, palabras, top_tabla):
    alturas = []
    for p in palabras:
        txt = p["text"].strip()
        if not _es_altura(txt):
            continue
        if p["top"] >= top_tabla:
            continue
        alturas.append({
            "x0": p["x0"],
            "x1": p["x1"],
            "top": p["top"],
            "valor": int(txt),
        })
    return alturas


def _detectar_curvas(pagina):
    curvas = []
    try:
        curvas = pagina.curves or []
    except Exception:
        curvas = []

    if not curvas:
        return [], []

    items = []
    for c in curvas:
        items.append({
            "x0": c["x0"],
            "x1": c["x1"],
            "top": c["top"],
            "bottom": c["bottom"],
        })

    if not items:
        return [], []

    items.sort(key=lambda c: c["top"])
    grupos = _agrupar_por_top(items, tolerancia=40)

    if len(grupos) >= 2:
        grupo_arriba = min(grupos, key=lambda g: min(c["top"] for c in g))
        grupo_abajo = max(grupos, key=lambda g: min(c["top"] for c in g))
        return grupo_arriba, grupo_abajo

    return [], []


def _buscar_alturas_cortas(pagina, palabras, top_tabla):
    curvas_arriba, curvas_abajo = _detectar_curvas(pagina)
    if not curvas_arriba and not curvas_abajo:
        return []

    candidatos = []
    for p in palabras:
        txt = p["text"].strip()
        if not _es_altura_corta(txt):
            continue
        if p["top"] >= top_tabla:
            continue
        candidatos.append({
            "x0": p["x0"],
            "x1": p["x1"],
            "top": p["top"],
            "valor": int(txt),
        })

    if not candidatos:
        return []

    alturas = []

    if curvas_arriba:
        top_curva_arriba = sum(c["top"] for c in curvas_arriba) / len(curvas_arriba)
        for c in candidatos:
            if abs(c["top"] - top_curva_arriba) <= 30:
                alturas.append(c)

    if curvas_abajo:
        top_curva_abajo = sum(c["top"] for c in curvas_abajo) / len(curvas_abajo)
        for c in candidatos:
            if abs(c["top"] - top_curva_abajo) <= 30:
                alturas.append(c)

    vistos = set()
    resultado = []
    for a in alturas:
        key = (round(a["x0"]), round(a["top"]))
        if key in vistos:
            continue
        vistos.add(key)
        resultado.append(a)

    return resultado


def _buscar_alturas(pagina, palabras, top_tabla):
    normales = _buscar_alturas_normales(pagina, palabras, top_tabla)
    if len(normales) >= 2:
        return normales

    cortas = _buscar_alturas_cortas(pagina, palabras, top_tabla)
    if len(cortas) >= 2:
        return cortas

    return normales


def _formar_combinaciones_alturas(alturas):
    if len(alturas) < 2:
        return []
    filas = _agrupar_por_top(alturas, tolerancia=20)
    if not filas:
        return []
    for f in filas:
        f.sort(key=lambda a: a["x0"])
    filas.sort(key=lambda f: f[0]["top"])

    if len(filas) >= 2:
        fila_arriba = filas[-2]
        fila_abajo = filas[-1]
        if len(fila_arriba) >= 2 and len(fila_abajo) >= 2:
            a_izq = fila_arriba[0]["valor"]
            a_der = fila_arriba[-1]["valor"]
            b_izq = fila_abajo[0]["valor"]
            b_der = fila_abajo[-1]["valor"]
            pares = []
            for a, b in [(a_izq, a_der), (b_izq, b_der)]:
                ini, fin = sorted([a, b])
                pares.append({"ini": ini, "fin": fin})
            return pares

    fila = filas[-1]
    valores = [a["valor"] for a in fila]
    ini, fin = min(valores), max(valores)
    return [{"ini": ini, "fin": fin}]


def _limpiar_texto_bloque(texto):
    return re.sub(r"\s+", " ", texto).strip()


def _es_texto_valido_calle(texto):
    if not texto:
        return False
    # Solo descartamos el bloque entero si es puramente numérico (alturas)
    if _es_solo_numeros(texto):
        return False
    return True


def _normalizar_calle_para_guardar(texto):
    """
    Devuelve la calle SIN las palabras vacías.
    Se usa para comparar contra el Excel.
    """
    if not texto:
        return ""
    palabras = texto.split()
    limpias = [p for p in palabras if p.upper() not in _PALABRAS_VACIAS]
    return " ".join(limpias)


def _extraer_calle(pagina, palabras, alturas, top_tabla):
    if not alturas:
        return None

    filas = _agrupar_por_top(alturas, tolerancia=20)
    if not filas:
        return None
    for f in filas:
        f.sort(key=lambda a: a["x0"])
    filas.sort(key=lambda f: f[0]["top"])

    fila_abajo = filas[-1]
    if len(fila_abajo) < 2:
        return None

    x_izq = fila_abajo[0]["x0"]
    x_der = fila_abajo[-1]["x1"]
    top_fila_abajo = fila_abajo[0]["top"]
    x_centro = (x_izq + x_der) / 2

    rects = pagina.rects or []
    candidatos = []
    for r in rects:
        if r["top"] >= top_tabla:
            continue
        if r["x0"] < x_izq - 20 or r["x1"] > x_der + 20:
            continue
        if r["top"] < top_fila_abajo - 5:
            continue

        adentro = _texto_dentro_de_rect(r, palabras)
        if not adentro:
            continue
        texto = " ".join(p["text"].strip() for p in adentro).strip()
        texto = _limpiar_texto_bloque(texto)

        if not _es_texto_valido_calle(texto):
            continue

        candidatos.append({
            "texto": texto,
            "top": r["top"],
            "x0": r["x0"],
            "x1": r["x1"],
        })

    if candidatos:
        candidatos.sort(key=lambda c: abs((c["x0"] + c["x1"]) / 2 - x_centro))
        return candidatos[0]["texto"]

    # Fallback: usar palabras individuales tal cual, sin filtrar nada
    palabras_cand = []
    for p in palabras:
        txt = p["text"].strip()
        if not txt:
            continue
        if p["top"] >= top_tabla:
            continue
        if p["top"] < top_fila_abajo - 5:
            continue
        xc = (p["x0"] + p["x1"]) / 2
        if xc < x_izq - 20 or xc > x_der + 20:
            continue

        palabras_cand.append({
            "texto": txt,
            "top": p["top"],
            "xc": xc,
        })

    if not palabras_cand:
        return None

    palabras_cand.sort(key=lambda p: (round(p["top"] / 10), p["xc"]))

    grupos = []
    for p in palabras_cand:
        if not grupos:
            grupos.append([p])
        else:
            ult = grupos[-1]
            ult_top = sum(x["top"] for x in ult) / len(ult)
            ult_x_max = max(x["xc"] for x in ult)
            if abs(p["top"] - ult_top) < 15 and (p["xc"] - ult_x_max) < 80:
                ult.append(p)
            else:
                grupos.append([p])

    grupos_validos = []
    for g in grupos:
        texto = " ".join(p["texto"] for p in g)
        texto = _limpiar_texto_bloque(texto)
        if _es_texto_valido_calle(texto):
            grupos_validos.append({
                "texto": texto,
                "xc_prom": sum(p["xc"] for p in g) / len(g),
            })

    if not grupos_validos:
        return None

    grupos_validos.sort(key=lambda c: abs(c["xc_prom"] - x_centro))
    return grupos_validos[0]["texto"]


def _es_item(txt):
    if not txt:
        return None
    t = str(txt).strip()
    m = re.match(r"^\s*(\d+(?:\.\d+)+)\s*(?:\(|$)", t)
    if m:
        return m.group(1)
    return None


def _detectar_tabla_y_items(tabla):
    idx_fila_ubicacion = None
    idx_col_ubicacion = None

    for i, fila in enumerate(tabla):
        celdas = [str(c).strip().upper() if c else "" for c in fila]
        if any("UBICACI" in c for c in celdas):
            idx_fila_ubicacion = i
            for j, c in enumerate(celdas):
                if "UBICACI" in c:
                    idx_col_ubicacion = j
                    break
            break

    if idx_fila_ubicacion is not None:
        fila_encabezado = tabla[idx_fila_ubicacion]
        items_por_col = {}
        for j, c in enumerate(fila_encabezado):
            item = _es_item(c)
            if item:
                items_por_col[j] = item

        columnas_total = []
        for i in range(idx_fila_ubicacion + 1, min(len(tabla), idx_fila_ubicacion + 4)):
            fila = tabla[i]
            for j, c in enumerate(fila):
                if not c:
                    continue
                if str(c).strip().upper() == "TOTAL":
                    item = None
                    if j in items_por_col:
                        item = items_por_col[j]
                    else:
                        for jj in range(j, -1, -1):
                            if jj in items_por_col:
                                item = items_por_col[jj]
                                break
                    if item:
                        columnas_total.append({
                            "idx_col": j,
                            "item": item,
                        })
            if columnas_total:
                break

        return {
            "modo": "ubicacion",
            "idx_fila_ubicacion": idx_fila_ubicacion,
            "idx_col_ubicacion": idx_col_ubicacion,
            "columnas_total": columnas_total,
        }

    idx_fila_largo = None
    for i, fila in enumerate(tabla):
        celdas = [str(c).strip().upper() if c else "" for c in fila]
        if any("LARGO" in c for c in celdas):
            idx_fila_largo = i
            break

    if idx_fila_largo is None:
        return None

    filas_a_mirar = []
    if idx_fila_largo - 1 >= 0:
        filas_a_mirar.append(tabla[idx_fila_largo - 1])
    filas_a_mirar.append(tabla[idx_fila_largo])

    items_por_col = {}
    for fila_enc in filas_a_mirar:
        for j, c in enumerate(fila_enc):
            item = _es_item(c)
            if item:
                items_por_col[j] = item

    if not items_por_col:
        return None

    columnas_total = []
    for i in range(idx_fila_largo + 1, len(tabla)):
        fila = tabla[i]
        for j, c in enumerate(fila):
            if not c:
                continue
            if str(c).strip().upper() == "TOTAL":
                item = None
                if j in items_por_col:
                    item = items_por_col[j]
                else:
                    for jj in range(j, -1, -1):
                        if jj in items_por_col:
                            item = items_por_col[jj]
                            break
                if item:
                    columnas_total.append({
                        "idx_col": j,
                        "item": item,
                    })
        if columnas_total:
            break

    return {
        "modo": "largo",
        "idx_fila_ubicacion": idx_fila_largo,
        "idx_col_ubicacion": 0,
        "columnas_total": columnas_total,
    }


def _procesar_tabla(tabla, info):
    resultados = []
    idx_fila = info["idx_fila_ubicacion"]
    idx_col_ubic = info["idx_col_ubicacion"]
    columnas_total = info["columnas_total"]
    modo = info.get("modo", "ubicacion")

    if not columnas_total:
        return resultados

    for fila in tabla[idx_fila + 1:]:
        es_total = False
        for c in fila:
            if c and str(c).strip().upper() == "TOTAL":
                es_total = True
                break
        if es_total:
            continue

        if modo == "ubicacion":
            if idx_col_ubic >= len(fila):
                continue
            celda = fila[idx_col_ubic]
            if not celda:
                continue
            celda_str = str(celda).strip()
            if not re.fullmatch(r"\d{1,3}", celda_str):
                continue
        else:
            tiene_datos = False
            for col in columnas_total:
                idx_col = col["idx_col"]
                if idx_col >= len(fila):
                    continue
                if _a_float(fila[idx_col]) is not None:
                    tiene_datos = True
                    break
            if not tiene_datos:
                continue

        for col in columnas_total:
            idx_col = col["idx_col"]
            item = col["item"]
            if idx_col >= len(fila):
                continue
            valor = _a_float(fila[idx_col])
            if valor is None:
                continue
            if valor == 0:
                continue
            resultados.append({
                "item": item,
                "valor": valor,
            })

    return resultados


def _extraer_items_de_tablas(pagina):
    resultados = []
    tablas = pagina.extract_tables() or []
    for tabla in tablas:
        if not tabla:
            continue
        info = _detectar_tabla_y_items(tabla)
        if info is None:
            continue
        items = _procesar_tabla(tabla, info)
        resultados.extend(items)
    return resultados


def _procesar_pagina(pagina) -> dict:
    resultado = {
        "acta_id": None,
        "care": None,
        "came": None,
        "calle": None,
        "calle_norm": None,
        "alturas": [],
        "items": [],
        "pagina": pagina.page_number,
    }

    palabras = pagina.extract_words(use_text_flow=False, keep_blank_chars=False)
    texto_pagina = pagina.extract_text() or ""

    resultado["acta_id"] = _extraer_acta_id(texto_pagina)
    care, came = _extraer_care_came(texto_pagina)
    resultado["care"] = care
    resultado["came"] = came

    if palabras:
        top_tabla = _top_de_la_tabla(pagina)
        alturas = _buscar_alturas(pagina, palabras, top_tabla)
        if len(alturas) >= 2:
            resultado["alturas"] = _formar_combinaciones_alturas(alturas)
            calle = _extraer_calle(pagina, palabras, alturas, top_tabla)
            if calle:
                resultado["calle"] = calle
                resultado["calle_norm"] = _normalizar_calle_para_guardar(calle)

    resultado["items"] = _extraer_items_de_tablas(pagina)
    return resultado


def parsear_pdf_acta(ruta_pdf: str) -> list:
    actas = []
    with pdfplumber.open(ruta_pdf) as pdf:
        for pagina in pdf.pages:
            resultado = _procesar_pagina(pagina)
            if resultado["acta_id"] or resultado["items"]:
                actas.append(resultado)
    return actas