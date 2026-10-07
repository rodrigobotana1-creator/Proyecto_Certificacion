"""
modules/excel_reader.py
"""

import re
import unicodedata
import pandas as pd


COL_TEXTO_BREVE    = "Texto breve"
COL_ORDEN_SUPERIOR = "Orden superior"
COL_ORDEN          = "Orden"
COL_DENOMINACION   = "Denominación"
COL_OPERACION      = "Txt.brv.oper."
COL_VALOR          = "Duración normal"

COL_CLASE_AVISO    = "Clase de aviso"

_PALABRAS_TIPO_VIA = {
    "CALZADA", "ACERA", "CALLE", "AV", "AVDA", "AVENIDA",
    "PASAJE", "PJE", "BOULEVARD", "BLVD",
}

IDX_COL_L = 11
IDX_COL_M = 12


def _quitar_tildes(s: str) -> str:
    if not s:
        return ""
    return "".join(
        c for c in unicodedata.normalize("NFD", str(s))
        if unicodedata.category(c) != "Mn"
    )


def obtener_nombres_hojas(ruta_excel: str) -> list:
    try:
        xls = pd.ExcelFile(ruta_excel)
        return list(xls.sheet_names)
    except Exception:
        return []


def leer_planilla(ruta_excel: str, hoja=None) -> pd.DataFrame:
    if hoja is None:
        df = pd.read_excel(ruta_excel, dtype=str)
    else:
        df = pd.read_excel(ruta_excel, sheet_name=hoja, dtype=str)

    df.columns = [str(c).strip() for c in df.columns]

    if COL_VALOR in df.columns:
        df[COL_VALOR] = df[COL_VALOR].apply(_a_float)

    return df


def _a_float(valor):
    if valor is None:
        return None
    if isinstance(valor, (int, float)):
        try:
            return float(valor)
        except (TypeError, ValueError):
            return None
    s = str(valor).strip().replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def obtener_actas_unicas(df: pd.DataFrame) -> list:
    if COL_TEXTO_BREVE not in df.columns:
        return []
    vistos = []
    for v in df[COL_TEXTO_BREVE].dropna().tolist():
        v = str(v).strip()
        if v and v not in vistos:
            vistos.append(v)
    return vistos


def obtener_datos_acta(df: pd.DataFrame, acta_id: str) -> pd.DataFrame:
    return df[df[COL_TEXTO_BREVE].astype(str).str.strip() == str(acta_id).strip()].copy()


def obtener_numero_acta(texto_breve: str) -> str:
    if not texto_breve:
        return ""
    digitos = re.findall(r"\d+", str(texto_breve))
    return str(int(digitos[-1])) if digitos else ""


def obtener_care(fila) -> str:
    return str(fila.get(COL_ORDEN_SUPERIOR, "")).strip()


def obtener_came(fila) -> str:
    return str(fila.get(COL_ORDEN, "")).strip()


def obtener_operacion(fila) -> str:
    raw = str(fila.get(COL_OPERACION, "")).strip()
    if not raw:
        return ""
    return raw.split("_")[0].strip()


def obtener_valor(fila):
    return _a_float(fila.get(COL_VALOR))


def obtener_calle_y_alturas(fila):
    texto = str(fila.get(COL_DENOMINACION, "")).strip()
    if not texto:
        return None
    texto = texto.replace("\xa0", " ").replace("\t", " ")
    texto = re.sub(r"\s+", " ", texto).strip()

    m = re.search(r"^(.+?)\s+(\d{1,15})\s*-\s*(\d{1,15})(?:\s|$)", texto)
    if not m:
        return None

    calle = m.group(1)
    ini = int(m.group(2))
    fin = int(m.group(3))

    calle = _normalizar_calle(calle)
    if not calle:
        return None

    return {"calle": calle, "ini": ini, "fin": fin}


def _normalizar_calle(texto: str) -> str:
    if not texto:
        return ""
    s = _quitar_tildes(str(texto).strip().upper())
    s = s.replace("\xa0", " ").replace("\t", " ")
    s = s.replace(",", " ")
    s = re.sub(r"[\.]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    palabras = [p for p in s.split(" ") if p and p not in _PALABRAS_TIPO_VIA]
    return " ".join(palabras)


# =====================================================================
#  Chequeo: Cl.actividad PM vs columna M
# =====================================================================

def obtener_cl_actividad_pm(fila) -> str:
    return str(fila.get("Cl.actividad PM", "")).strip()


def obtener_col_m(fila) -> str:
    try:
        val = fila.iloc[IDX_COL_M]
    except Exception:
        return ""
    if val is None:
        return ""
    return str(val).strip()


def chequear_actividad(fila) -> dict:
    pm = obtener_cl_actividad_pm(fila)
    m = obtener_col_m(fila)
    ok = bool(pm and m and pm == m)
    return {
        "valor_pm": pm,
        "valor_m": m,
        "ok": ok,
    }


# =====================================================================
#  Chequeo: Denominación vs columna L
# =====================================================================

def obtener_col_l(fila) -> str:
    try:
        val = fila.iloc[IDX_COL_L]
    except Exception:
        return ""
    if val is None:
        return ""
    return str(val).strip()


def _parsear_col_l(texto_l):
    if not texto_l:
        return None
    t = str(texto_l).strip()
    t = t.replace("\xa0", " ").replace("\t", " ")
    t = re.sub(r"\s+", " ", t).strip()

    m = re.search(r"^(.+?)\s+(\d{1,15})(?:/(\d{2}))?", t)
    if not m:
        return None

    calle_raw = m.group(1)
    try:
        altura1 = int(m.group(2))
    except ValueError:
        return None
    altura2_extra = m.group(3)

    calle = _normalizar_calle(calle_raw)
    if not calle:
        return None

    alturas = [altura1]

    if altura2_extra:
        primeros = str(altura1)[:2]
        try:
            alturas.append(int(primeros + altura2_extra))
        except ValueError:
            pass

    return {"calle": calle, "alturas": alturas}


def _calles_coinciden(calle_a, calle_b):
    if not calle_a or not calle_b:
        return False
    if calle_a == calle_b:
        return True
    palabras_a = set(calle_a.split())
    palabras_b = set(calle_b.split())
    return bool(palabras_a & palabras_b)


def chequear_direccion_col_l(fila, tolerancia_extremos=2) -> dict:
    texto_den = str(fila.get(COL_DENOMINACION, "")).strip()
    texto_l = obtener_col_l(fila)

    den = obtener_calle_y_alturas(fila)
    l = _parsear_col_l(texto_l)

    if den is None or l is None:
        return {
            "denominacion": texto_den,
            "col_l": texto_l,
            "ok_calle": False,
            "ok_altura": False,
            "ok": False,
        }

    ok_calle = _calles_coinciden(den["calle"], l["calle"])

    ini = den["ini"]
    fin = den["fin"]

    ok_altura = False
    for a in l["alturas"]:
        if ini <= a <= fin:
            ok_altura = True
            break
        if (
            round(abs(a - ini), 5) <= tolerancia_extremos
            or round(abs(a - fin), 5) <= tolerancia_extremos
        ):
            ok_altura = True
            break

    return {
        "denominacion": texto_den,
        "col_l": texto_l,
        "ok_calle": ok_calle,
        "ok_altura": ok_altura,
        "ok": ok_calle and ok_altura,
    }


# =====================================================================
#  Chequeo: Clase de aviso según Cl.actividad PM
# =====================================================================

def obtener_clase_aviso(fila) -> str:
    """Devuelve el valor del encabezado 'Clase de aviso'."""
    return str(fila.get(COL_CLASE_AVISO, "")).strip().upper()


def chequear_clase_aviso(fila) -> dict:
    """
    Reglas:
      - Cl.actividad PM y col M deben coincidir. Si no → error.
      - Si ambas son 'G02' → clase de aviso debe ser AP, EM o CG.
      - Si ambas son otro valor → clase de aviso debe ser OF, SU o RE.
    """
    pm = obtener_cl_actividad_pm(fila)
    m = obtener_col_m(fila)
    aviso = obtener_clase_aviso(fila)

    # Normalizar a mayúsculas para comparar
    pm_up = pm.upper()
    m_up = m.upper()

    # --- Caso 1: no coinciden las clases de actividad ---
    if not pm or not m:
        return {
            "valor_pm": pm,
            "valor_m": m,
            "valor_aviso": aviso,
            "ok": False,
            "mensaje": "Falta clase de actividad",
        }

    if pm_up != m_up:
        return {
            "valor_pm": pm,
            "valor_m": m,
            "valor_aviso": aviso,
            "ok": False,
            "mensaje": "CLASES DE ACTIVIDAD DISTINTAS, CORREGIR",
        }

    # --- Caso 2: coinciden. Aplicar regla según valor ---
    if pm_up == "G02":
        validos = {"AP", "EM", "CG"}
    else:
        validos = {"OF", "SU", "RE"}

    if not aviso:
        return {
            "valor_pm": pm,
            "valor_m": m,
            "valor_aviso": "",
            "ok": False,
            "mensaje": "Clase de aviso vacía",
        }

    ok = aviso in validos
    if ok:
        mensaje = f"Clase de aviso: {aviso} ✓"
    else:
        mensaje = f"Clase de aviso incorrecta: {aviso} (esperado: {'/'.join(sorted(validos))})"

    return {
        "valor_pm": pm,
        "valor_m": m,
        "valor_aviso": aviso,
        "ok": ok,
        "mensaje": mensaje,
    }


# =====================================================================
#  Nombres de encabezados de columnas L y M
# =====================================================================

def obtener_nombre_col_l(df) -> str:
    try:
        if df is None or len(df.columns) <= IDX_COL_L:
            return "Col L"
        nombre = str(df.columns[IDX_COL_L]).strip()
        return nombre if nombre else "Col L"
    except Exception:
        return "Col L"


def obtener_nombre_col_m(df) -> str:
    try:
        if df is None or len(df.columns) <= IDX_COL_M:
            return "Col M"
        nombre = str(df.columns[IDX_COL_M]).strip()
        return nombre if nombre else "Col M"
    except Exception:
        return "Col M"