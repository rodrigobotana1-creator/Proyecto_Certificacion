import os
import sys
import re
import base64
import platform
import subprocess
import urllib.parse
import webbrowser
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QTableWidget, QTableWidgetItem,
    QSplitter, QProgressBar, QListWidget, QListWidgetItem,
    QMessageBox, QFrame, QAbstractItemView, QScrollArea,
    QDialog, QListWidget as QListWidgetDialog, QDialogButtonBox,
    QGraphicsDropShadowEffect, QSizePolicy, QCheckBox, QGridLayout,
    QStackedWidget, QTableView, QHeaderView, QMenu, QLineEdit,
    QStyleOptionButton, QStyle
)
from PyQt5.QtGui import (
    QColor, QFont, QPixmap, QTransform, QImage,
    QStandardItemModel, QStandardItem
)
from PyQt5.QtCore import (
    Qt, QSize, QTimer, QSortFilterProxyModel, QModelIndex
)

import qtawesome as qta

try:
    import pymupdf as fitz
    PYMUPDF_OK = True
except ImportError:
    try:
        import fitz
        PYMUPDF_OK = True
    except ImportError:
        PYMUPDF_OK = False

from modules.excel_reader import (
    leer_planilla,
    obtener_nombres_hojas,
    obtener_actas_unicas,
    obtener_datos_acta,
    obtener_numero_acta,
    obtener_care,
    obtener_came,
    obtener_operacion,
    obtener_valor,
    chequear_actividad,
    chequear_direccion_col_l,
    chequear_clase_aviso,
    obtener_nombre_col_l,
    obtener_nombre_col_m,
    COL_TEXTO_BREVE,
    COL_ORDEN_SUPERIOR,
    COL_ORDEN,
    COL_DENOMINACION,
    COL_OPERACION,
    COL_VALOR,
)
from modules.pdf_parser import parsear_pdf_acta
from modules.comparator import comparar, comparar_direccion
from modules.batch_loader import (
    escanear_carpeta_pdfs,
    matchear_acta_excel,
    obtener_actas_de_indice,
    _solo_numero_acta,
    _normalizar_id,
)


BG          = "#F1F3F7"
CARD        = "#FFFFFF"
CARD_2      = "#F9FAFB"
BORDER      = "#E4E7EC"
BORDER_2    = "#D0D5DD"
TEXT        = "#101828"
TEXT_MUTED  = "#667085"
TEXT_DIM    = "#98A2B3"
ACCENT      = "#2563EB"
ACCENT_H    = "#1D4ED8"
ACCENT_BG   = "#EFF6FF"

OK_BG       = "#ECFDF5"
OK_FG       = "#065F46"
OK_DOT      = "#10B981"
WARN_BG     = "#FFFBEB"
WARN_FG     = "#92400E"
WARN_DOT    = "#F59E0B"
ERR_BG      = "#FEF2F2"
ERR_FG      = "#B91C1C"
ERR_DOT     = "#EF4444"
NONE_DOT    = "#CBD5E1"

SEL_BG      = "#EFF6FF"
INFO_BG     = "#F9FAFB"


EXT_IMAGENES = {
    ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".webp",
    ".tiff", ".tif", ".ico", ".svg",
}

EXT_DOCUMENTOS = {
    ".pdf",
    ".doc", ".docx",
    ".xls", ".xlsx",
    ".ppt", ".pptx",
    ".csv", ".rtf",
    ".odt", ".ods", ".odp",
}

EXT_HTML = {".html", ".htm"}

EXT_VALIDAS = EXT_IMAGENES | EXT_DOCUMENTOS


ROTACIONES = {}


def _get_rotacion(ruta):
    return ROTACIONES.get(os.path.abspath(ruta), 0)


def _set_rotacion(ruta, grados):
    ROTACIONES[os.path.abspath(ruta)] = grados % 360


_PATRON_BASE64 = re.compile(
    r'src=["\']data:image/(jpeg|jpg|png|gif|webp|bmp);base64,([A-Za-z0-9+/=\s]+)["\']',
    re.IGNORECASE | re.DOTALL,
)

_PATRON_BASE64_GENERICO = re.compile(
    r'data:image/(jpeg|jpg|png|gif|webp|bmp);base64,([A-Za-z0-9+/=\s]+)',
    re.IGNORECASE | re.DOTALL,
)


def _extraer_imagenes_de_html(ruta_html):
    try:
        with open(ruta_html, "r", encoding="utf-8", errors="ignore") as f:
            contenido = f.read()
    except Exception:
        return []

    matches = list(_PATRON_BASE64.finditer(contenido))
    if not matches:
        matches = list(_PATRON_BASE64_GENERICO.finditer(contenido))
    if not matches:
        return []

    carpeta_html = os.path.dirname(ruta_html)
    carpeta_cache = os.path.join(carpeta_html, "_cache")
    try:
        os.makedirs(carpeta_cache, exist_ok=True)
    except Exception:
        return []

    nombre_base = os.path.splitext(os.path.basename(ruta_html))[0]
    rutas = []

    for i, m in enumerate(matches, 1):
        ext = m.group(1).lower()
        if ext == "jpg":
            ext = "jpeg"
        datos_b64 = re.sub(r"\s+", "", m.group(2))
        try:
            datos_bytes = base64.b64decode(datos_b64)
        except Exception:
            continue
        if len(matches) == 1:
            nombre_img = f"{nombre_base}.{ext}"
        else:
            nombre_img = f"{nombre_base}_{i}.{ext}"
        ruta_img = os.path.join(carpeta_cache, nombre_img)
        try:
            with open(ruta_img, "wb") as f:
                f.write(datos_bytes)
            rutas.append(ruta_img)
        except Exception:
            continue

    return rutas


def _clasificar_archivo(ruta):
    nombre = os.path.basename(ruta).lower()
    if "care" in nombre:
        return "care"
    if "came" in nombre:
        return "came"
    return None


def _es_archivo_analizable(ruta):
    ext = os.path.splitext(ruta)[1].lower()
    if ext in EXT_HTML:
        return False
    if ext in (".txt", ".rtf"):
        return False
    if ext in EXT_VALIDAS:
        return True
    return False


def _dot(color: str, size: int = 10) -> QLabel:
    lbl = QLabel()
    lbl.setFixedSize(size, size)
    lbl.setStyleSheet(f"background-color: {color}; border-radius: {size // 2}px;")
    return lbl


def _make_icon(name: str, color: str = "#374151"):
    return qta.icon(name, color=color)


def _fmt(v):
    if v is None:
        return "—"
    if isinstance(v, float):
        return f"{v:.2f}".replace(".", ",")
    return str(v)


def _formatear_fecha(valor):
    if valor is None:
        return ""
    s = str(valor).strip()
    if not s:
        return ""
    if " 00:00:00" in s:
        return s.split(" 00:00:00")[0]
    if "T00:00:00" in s:
        return s.split("T00:00:00")[0]
    return s


def _es_fecha(valor):
    if valor is None:
        return False
    s = str(valor).strip()
    if not s:
        return False
    return bool(re.match(r"^\d{4}-\d{2}-\d{2}", s))


def _es_numero(valor):
    if valor is None:
        return False
    if isinstance(valor, (int, float)):
        return True
    s = str(valor).strip()
    if not s:
        return False
    try:
        float(s.replace(",", "."))
        return True
    except ValueError:
        return False


def _titulo_seccion(texto):
    lbl = QLabel(texto)
    lbl.setObjectName("SectionLabel")
    return lbl


def _abrir_con_sistema(ruta):
    try:
        if platform.system() == "Windows":
            os.startfile(ruta)
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", ruta])
        else:
            subprocess.Popen(["xdg-open", ruta])
        return True
    except Exception:
        return False


def _abrir_en_maps(calle, altura=None):
    if not calle:
        return False

    partes = [str(calle).strip()]
    if altura:
        partes.append(str(altura).strip())
    partes.append("CABA")
    partes.append("Buenos Aires")
    partes.append("Argentina")

    query = ", ".join(p for p in partes if p)
    url = "https://www.google.com/maps/search/?api=1&query=" + urllib.parse.quote_plus(query)

    try:
        webbrowser.open(url)
        return True
    except Exception:
        return False


def _calle_altura_de_resultado(r):
    if not r:
        return None, None

    calle = None
    altura = None

    datos_pdf = r.get("datos_pdf") or {}
    pdf_calle = (datos_pdf.get("calle") or "").strip()
    pdf_alturas = datos_pdf.get("alturas") or []

    if pdf_calle:
        calle = pdf_calle

    if pdf_alturas:
        a = pdf_alturas[0]
        ini_s = str(a.get("ini", "")).strip()
        fin_s = str(a.get("fin", "")).strip()
        try:
            ini = int(ini_s)
            fin = int(fin_s)
            altura = (ini + fin) // 2
        except Exception:
            try:
                altura = int(ini_s)
            except Exception:
                altura = None

    if not calle or altura is None:
        df_acta = r.get("df_acta")
        fila0 = None
        if df_acta is not None and not df_acta.empty:
            fila0 = df_acta.iloc[0]

        if fila0 is not None:
            if not calle:
                texto_dir = str(fila0.get(COL_DENOMINACION, "")).strip()
                calle = texto_dir
            if altura is None:
                m = re.search(r"(\d+)\s*$", str(calle) if calle else "")
                if m:
                    altura = int(m.group(1))

    return calle, altura


def _es_imagen(ruta):
    ext = os.path.splitext(ruta)[1].lower()
    return ext in EXT_IMAGENES


def _es_pdf(ruta):
    ext = os.path.splitext(ruta)[1].lower()
    return ext == ".pdf"


def _cargar_pixmap(ruta):
    ext = os.path.splitext(ruta)[1].lower()

    if ext in EXT_IMAGENES:
        pm = QPixmap(ruta)
        return pm if not pm.isNull() else None

    if ext == ".pdf" and PYMUPDF_OK:
        try:
            doc = fitz.open(ruta)
            if doc.page_count == 0:
                doc.close()
                return None
            pagina = doc.load_page(0)
            mat = fitz.Matrix(1.5, 1.5)
            pix = pagina.get_pixmap(matrix=mat, alpha=False)
            img = QImage(
                pix.samples,
                pix.width,
                pix.height,
                pix.stride,
                QImage.Format_RGB888,
            )
            pm = QPixmap.fromImage(img.copy())
            doc.close()
            return pm
        except Exception:
            return None

    return None


def _cargar_pixmaps_pdf(ruta):
    if not (PYMUPDF_OK and _es_pdf(ruta)):
        return []
    try:
        doc = fitz.open(ruta)
        if doc.page_count == 0:
            doc.close()
            return []
        pms = []
        mat = fitz.Matrix(1.5, 1.5)
        for i in range(doc.page_count):
            pagina = doc.load_page(i)
            pix = pagina.get_pixmap(matrix=mat, alpha=False)
            img = QImage(
                pix.samples,
                pix.width,
                pix.height,
                pix.stride,
                QImage.Format_RGB888,
            )
            pms.append(QPixmap.fromImage(img.copy()))
        doc.close()
        return pms
    except Exception:
        return []


def _cargar_todas_las_paginas(ruta):
    ext = os.path.splitext(ruta)[1].lower()
    if ext in EXT_IMAGENES:
        pm = QPixmap(ruta)
        return [pm] if not pm.isNull() else []
    if ext == ".pdf":
        return _cargar_pixmaps_pdf(ruta)
    return []


def _icono_para_archivo(ruta):
    ext = os.path.splitext(ruta)[1].lower()
    if ext == ".pdf":
        return ("fa5s.file-pdf", "#DC2626")
    if ext in (".doc", ".docx"):
        return ("fa5s.file-word", "#2563EB")
    if ext in (".xls", ".xlsx", ".csv", ".ods"):
        return ("fa5s.file-excel", "#166534")
    if ext in (".ppt", ".pptx", ".odp"):
        return ("fa5s.file-powerpoint", "#EA580C")
    if ext in (".txt", ".rtf", ".odt"):
        return ("fa5s.file-alt", "#475569")
    return ("fa5s.file", "#64748B")


def _fila_info(etiqueta, valor, color_valor=None):
    w = QWidget()
    h = QHBoxLayout(w)
    h.setContentsMargins(0, 4, 0, 4)
    h.setSpacing(12)
    lbl_e = QLabel(etiqueta + ":")
    lbl_e.setStyleSheet(f"color: {TEXT_MUTED}; font-size: 9.5pt;")
    lbl_e.setFixedWidth(200)
    h.addWidget(lbl_e)
    lbl_v = QLabel(str(valor) if valor not in (None, "") else "—")
    color = color_valor or TEXT
    lbl_v.setStyleSheet(f"color: {color}; font-size: 9.5pt; font-weight: 500;")
    lbl_v.setWordWrap(True)
    lbl_v.setTextInteractionFlags(Qt.TextSelectableByMouse)
    h.addWidget(lbl_v, 1)
    return w


def _fila_cmp(etiqueta, valor_sap, valor_pdf, ok, extra=""):
    w = QWidget()
    w.setStyleSheet(f"background-color: {OK_BG if ok else ERR_BG}; border-radius: 6px;")
    h = QHBoxLayout(w)
    h.setContentsMargins(12, 8, 12, 8)
    h.setSpacing(10)
    ico = QLabel("✓" if ok else "✗")
    ico.setStyleSheet(f"color: {OK_FG if ok else ERR_FG}; font-weight: 700; font-size: 12pt;")
    ico.setFixedWidth(20)
    h.addWidget(ico)
    lbl_e = QLabel(etiqueta)
    lbl_e.setStyleSheet(f"color: {TEXT}; font-size: 9.5pt; font-weight: 600;")
    lbl_e.setMinimumWidth(220)
    lbl_e.setWordWrap(False)
    h.addWidget(lbl_e)
    lbl_sap = QLabel(str(valor_sap) if valor_sap not in (None, "") else "—")
    lbl_sap.setStyleSheet(f"color: {TEXT}; font-size: 9.5pt;")
    lbl_sap.setMinimumWidth(200)
    lbl_sap.setWordWrap(False)
    lbl_sap.setTextInteractionFlags(Qt.TextSelectableByMouse)
    h.addWidget(lbl_sap)
    lbl_igual = QLabel("=" if ok else "≠")
    lbl_igual.setStyleSheet(f"color: {TEXT_MUTED}; font-size: 10pt; font-weight: 600;")
    lbl_igual.setFixedWidth(24)
    lbl_igual.setAlignment(Qt.AlignCenter)
    h.addWidget(lbl_igual)
    lbl_pdf = QLabel(str(valor_pdf) if valor_pdf not in (None, "") else "—")
    lbl_pdf.setStyleSheet(f"color: {TEXT}; font-size: 9.5pt;")
    lbl_pdf.setMinimumWidth(200)
    lbl_pdf.setWordWrap(False)
    lbl_pdf.setTextInteractionFlags(Qt.TextSelectableByMouse)
    h.addWidget(lbl_pdf, 1)
    if extra:
        lbl_extra = QLabel(extra)
        lbl_extra.setStyleSheet(f"color: {TEXT_MUTED}; font-size: 9pt; font-style: italic;")
        lbl_extra.setWordWrap(False)
        h.addWidget(lbl_extra)
    return w


def _fila_check(etiqueta, valor, ok, mensaje=""):
    w = QWidget()
    w.setStyleSheet(f"background-color: {OK_BG if ok else ERR_BG}; border-radius: 6px;")
    h = QHBoxLayout(w)
    h.setContentsMargins(12, 8, 12, 8)
    h.setSpacing(10)
    ico = QLabel("✓" if ok else "✗")
    ico.setStyleSheet(f"color: {OK_FG if ok else ERR_FG}; font-weight: 700; font-size: 12pt;")
    ico.setFixedWidth(20)
    h.addWidget(ico)
    lbl_e = QLabel(etiqueta)
    lbl_e.setStyleSheet(f"color: {TEXT}; font-size: 9.5pt; font-weight: 600;")
    lbl_e.setMinimumWidth(280)
    lbl_e.setWordWrap(False)
    h.addWidget(lbl_e)
    lbl_v = QLabel(str(valor) if valor not in (None, "") else "—")
    lbl_v.setStyleSheet(f"color: {TEXT}; font-size: 9.5pt; font-weight: 700;")
    lbl_v.setMinimumWidth(60)
    lbl_v.setTextInteractionFlags(Qt.TextSelectableByMouse)
    h.addWidget(lbl_v)
    h.addStretch(1)
    if mensaje:
        lbl_m = QLabel(mensaje)
        color_msg = OK_FG if ok else ERR_FG
        lbl_m.setStyleSheet(f"color: {color_msg}; font-size: 9pt; font-weight: 600;")
        lbl_m.setWordWrap(False)
        h.addWidget(lbl_m)
    return w


class FiltroMenu(QMenu):
    def __init__(self, columna, valores_unicos, seleccionados_actuales,
                 callback_aplicar, parent=None):
        super().__init__(parent)
        self.columna = columna
        self.valores_unicos = list(valores_unicos)
        self.seleccionados = set(seleccionados_actuales)
        self.callback_aplicar = callback_aplicar

        self.setStyleSheet(f"""
            QMenu {{
                background-color: {CARD};
                border: 1px solid {BORDER};
                border-radius: 8px;
                padding: 4px;
            }}
        """)

        widget = QWidget()
        v = QVBoxLayout(widget)
        v.setContentsMargins(8, 8, 8, 8)
        v.setSpacing(6)

        self.input_buscar = QLineEdit()
        self.input_buscar.setPlaceholderText("Buscar…")
        self.input_buscar.setStyleSheet(f"""
            QLineEdit {{
                background-color: {CARD};
                border: 1px solid {BORDER_2};
                border-radius: 5px;
                padding: 6px 10px;
                font-size: 9pt;
                color: {TEXT};
            }}
            QLineEdit:focus {{
                border-color: {ACCENT};
            }}
        """)
        self.input_buscar.textChanged.connect(self._filtrar_valores)
        v.addWidget(self.input_buscar)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setMinimumHeight(200)
        scroll.setMaximumHeight(300)
        scroll.setStyleSheet(f"""
            QScrollArea {{
                background-color: {CARD};
                border: 1px solid {BORDER};
                border-radius: 5px;
            }}
        """)

        self.lista_widget = QWidget()
        self.lista_widget.setStyleSheet(f"background-color: {CARD};")
        self.lista_layout = QVBoxLayout(self.lista_widget)
        self.lista_layout.setContentsMargins(6, 6, 6, 6)
        self.lista_layout.setSpacing(2)

        self._checkboxes = {}

        for val in self.valores_unicos:
            chk = QCheckBox(str(val) if val != "" else "(vacío)")
            chk.setStyleSheet(f"""
                QCheckBox {{
                    color: {TEXT};
                    font-size: 9pt;
                    padding: 4px 6px;
                    border-radius: 4px;
                }}
                QCheckBox:hover {{
                    background-color: {CARD_2};
                }}
                QCheckBox::indicator {{
                    width: 15px;
                    height: 15px;
                    border-radius: 3px;
                    border: 1px solid {BORDER_2};
                    background: {CARD};
                }}
                QCheckBox::indicator:checked {{
                    background-color: {ACCENT};
                    border-color: {ACCENT};
                }}
            """)
            if val in self.seleccionados:
                chk.setChecked(True)
            chk.setCursor(Qt.PointingHandCursor)
            self._checkboxes[val] = chk
            self.lista_layout.addWidget(chk)

        self.lista_layout.addStretch(1)
        scroll.setWidget(self.lista_widget)
        v.addWidget(scroll)

        botones = QHBoxLayout()
        botones.setSpacing(6)

        btn_todas = QPushButton("Todas")
        btn_todas.setCursor(Qt.PointingHandCursor)
        btn_todas.setStyleSheet(f"""
            QPushButton {{
                background-color: {CARD};
                color: {TEXT};
                border: 1px solid {BORDER_2};
                border-radius: 5px;
                padding: 5px 10px;
                font-size: 8.5pt;
                font-weight: 600;
            }}
            QPushButton:hover {{ background-color: {CARD_2}; }}
        """)
        btn_todas.clicked.connect(self._marcar_todas)
        botones.addWidget(btn_todas)

        btn_limpiar = QPushButton("Limpiar")
        btn_limpiar.setCursor(Qt.PointingHandCursor)
        btn_limpiar.setStyleSheet(f"""
            QPushButton {{
                background-color: {CARD};
                color: {TEXT};
                border: 1px solid {BORDER_2};
                border-radius: 5px;
                padding: 5px 10px;
                font-size: 8.5pt;
                font-weight: 600;
            }}
            QPushButton:hover {{ background-color: {CARD_2}; }}
        """)
        btn_limpiar.clicked.connect(self._limpiar)
        botones.addWidget(btn_limpiar)

        botones.addStretch(1)

        btn_ok = QPushButton("Aplicar")
        btn_ok.setCursor(Qt.PointingHandCursor)
        btn_ok.setStyleSheet(f"""
            QPushButton {{
                background-color: {ACCENT};
                color: #FFFFFF;
                border: none;
                border-radius: 5px;
                padding: 6px 14px;
                font-size: 8.5pt;
                font-weight: 700;
            }}
            QPushButton:hover {{ background-color: {ACCENT_H}; }}
        """)
        btn_ok.clicked.connect(self._aplicar)
        botones.addWidget(btn_ok)

        v.addLayout(botones)

        from PyQt5.QtWidgets import QWidgetAction
        accion = QWidgetAction(self)
        accion.setDefaultWidget(widget)
        self.addAction(accion)

    def _filtrar_valores(self, texto):
        texto = texto.strip().lower()
        for val, chk in self._checkboxes.items():
            visible = texto == "" or texto in str(val).lower()
            chk.setVisible(visible)

    def _marcar_todas(self):
        for chk in self._checkboxes.values():
            if chk.isVisible():
                chk.setChecked(True)

    def _limpiar(self):
        for chk in self._checkboxes.values():
            if chk.isVisible():
                chk.setChecked(False)

    def _aplicar(self):
        seleccionados = {val for val, chk in self._checkboxes.items() if chk.isChecked()}
        self.callback_aplicar(self.columna, seleccionados)
        self.close()


class ProxyTabla(QSortFilterProxyModel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.filtros = {}

    def set_filtro_columna(self, columna, valores):
        if valores is None:
            self.filtros.pop(columna, None)
        else:
            self.filtros[columna] = set(valores)
        self.invalidateFilter()

    def limpiar_filtros(self):
        self.filtros = {}
        self.invalidateFilter()

    def filterAcceptsRow(self, source_row, source_parent):
        model = self.sourceModel()
        for columna, valores in self.filtros.items():
            idx = model.index(source_row, columna, source_parent)
            val = model.data(idx)
            if val is None:
                val = ""
            if val not in valores:
                return False
        return True


class TablaFiltrable(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.columnas = []
        self.datos_completos = []

        v = QVBoxLayout(self)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(0)

        self.model = QStandardItemModel()
        self.proxy = ProxyTabla()
        self.proxy.setSourceModel(self.model)

        self.table = QTableView()
        self.table.setModel(self.proxy)
        self.table.setAlternatingRowColors(True)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionMode(QAbstractItemView.NoSelection)
        self.table.setShowGrid(False)
        self.table.setWordWrap(False)
        self.table.verticalHeader().setVisible(False)
        self.table.setSortingEnabled(False)

        header = self.table.horizontalHeader()
        header.setSectionsClickable(True)
        header.setHighlightSections(False)
        header.sectionClicked.connect(self._abrir_menu_filtro)
        header.setStyleSheet(f"""
            QHeaderView::section {{
                background-color: #F9FAFB;
                color: {TEXT_MUTED};
                padding: 11px 12px;
                border: none;
                border-bottom: 1px solid {BORDER};
                border-right: 1px solid #F3F4F6;
                font-weight: 700;
                font-size: 8.5pt;
                letter-spacing: 0.4px;
            }}
            QHeaderView::section:hover {{
                background-color: {SEL_BG};
            }}
        """)

        v.addWidget(self.table, 1)

    def cargar_datos(self, columnas, filas):
        self.columnas = list(columnas)
        self.datos_completos = [list(f) for f in filas]

        self.model.clear()
        self.model.setHorizontalHeaderLabels(self.columnas)

        for fila in self.datos_completos:
            items = []
            for val in fila:
                item = QStandardItem(str(val) if val is not None else "")
                item.setEditable(False)
                items.append(item)
            self.model.appendRow(items)

        self.table.resizeColumnsToContents()
        for c in range(self.table.model().columnCount()):
            w = self.table.columnWidth(c)
            if w < 120:
                self.table.setColumnWidth(c, 120)
            elif w > 320:
                self.table.setColumnWidth(c, 320)

        self.proxy.limpiar_filtros()

    def _valores_unicos_disponibles(self, columna):
        filtro_propio = self.proxy.filtros.pop(columna, None)
        self.proxy.invalidateFilter()

        valores = set()
        for row in range(self.proxy.rowCount()):
            idx = self.proxy.index(row, columna)
            val = self.proxy.data(idx)
            if val is None:
                val = ""
            valores.add(val)

        if filtro_propio is not None:
            self.proxy.filtros[columna] = filtro_propio
            self.proxy.invalidateFilter()

        def _key(v):
            try:
                return (0, float(v.replace(",", ".")))
            except Exception:
                return (1, str(v).lower())

        return sorted(valores, key=_key)

    def _seleccionados_actuales(self, columna):
        if columna in self.proxy.filtros:
            return self.proxy.filtros[columna]
        return set(self._valores_unicos_completos(columna))

    def _valores_unicos_completos(self, columna):
        valores = set()
        for row in range(self.model.rowCount()):
            idx = self.model.index(row, columna)
            val = self.model.data(idx)
            if val is None:
                val = ""
            valores.add(val)
        return valores

    def _abrir_menu_filtro(self, columna):
        valores_disponibles = self._valores_unicos_disponibles(columna)
        seleccionados = self._seleccionados_actuales(columna)
        seleccionados = {s for s in seleccionados if s in valores_disponibles}

        menu = FiltroMenu(
            columna,
            valores_disponibles,
            seleccionados,
            self._aplicar_filtro,
            self,
        )

        header = self.table.horizontalHeader()
        x = header.sectionViewportPosition(columna)
        y = header.height()
        global_pos = header.mapToGlobal(header.rect().topLeft())
        global_pos.setX(global_pos.x() + x)
        global_pos.setY(global_pos.y() + y)
        menu.exec_(global_pos)

    def _aplicar_filtro(self, columna, valores):
        todos = self._valores_unicos_completos(columna)
        if valores == todos:
            self.proxy.set_filtro_columna(columna, None)
        else:
            self.proxy.set_filtro_columna(columna, valores)


class ZoomSpin(QLineEdit):
    """Campo de texto que muestra y permite escribir el % exacto de zoom."""

    def __init__(self, celda_getter, aplicar_fn, ancho=60, parent=None):
        super().__init__(parent)
        self.celda_getter = celda_getter
        self.aplicar_fn = aplicar_fn
        self.setFixedWidth(ancho)
        self.setAlignment(Qt.AlignCenter)
        self.setStyleSheet(f"""
            QLineEdit {{
                background-color: {CARD};
                color: {TEXT};
                border: 1px solid {BORDER_2};
                border-radius: 5px;
                padding: 3px 4px;
                font-size: 9pt;
                font-weight: 700;
            }}
            QLineEdit:focus {{
                border-color: {ACCENT};
                background-color: {ACCENT_BG};
            }}
        """)
        self.setToolTip("Escribí un % y apretá Enter (ej: 80)")
        self.returnPressed.connect(self._aplicar_escrito)
        self.editingFinished.connect(self._al_perder_foco)
        self.actualizar_texto()

    def actualizar_texto(self):
        if self.hasFocus() and self.isModified():
            return
        try:
            celda = self.celda_getter()
            if celda is not None:
                pct = celda.porcentaje_actual()
                self.blockSignals(True)
                self.setText(str(pct))
                self.blockSignals(False)
                self.setModified(False)
        except Exception:
            pass

    def _aplicar_escrito(self):
        texto = self.text().strip().replace("%", "").replace(",", ".")
        if not texto:
            self.actualizar_texto()
            return
        try:
            valor = float(texto)
        except ValueError:
            self.actualizar_texto()
            return
        if valor <= 0:
            self.actualizar_texto()
            return
        celda = self.celda_getter()
        if celda is not None:
            celda.zoom_set(valor)
        self.aplicar_fn()
        self.setModified(False)

    def _al_perder_foco(self):
        self.actualizar_texto()
        self.setModified(False)

    def mousePressEvent(self, event):
        super().mousePressEvent(event)
        QTimer.singleShot(0, self.selectAll)


class CeldaFoto(QScrollArea):
    def __init__(self, ruta_o_pixmap, parent=None, ajustar_al_alto=False):
        super().__init__(parent)
        self.ajustar_al_alto = ajustar_al_alto

        if isinstance(ruta_o_pixmap, QPixmap):
            self.ruta = None
            self._pm_inicial = ruta_o_pixmap
        else:
            self.ruta = ruta_o_pixmap
            self._pm_inicial = None

        self.rotacion = _get_rotacion(self.ruta) if self.ruta else 0
        self.pixmaps = []
        self.zoom = None

        self.setWidgetResizable(True)
        self.setFrameShape(QFrame.NoFrame)
        self.setStyleSheet("QScrollArea { background: transparent; border: none; }")
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        self._cont = QWidget()
        self._cont.setStyleSheet("background: transparent;")
        self._cont.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Minimum)

        self._vbox = QVBoxLayout(self._cont)
        self._vbox.setContentsMargins(0, 0, 0, 0)
        self._vbox.setSpacing(8)
        self._vbox.setAlignment(Qt.AlignTop | Qt.AlignHCenter)
        self.setWidget(self._cont)

        self.viewport().installEventFilter(self)
        self._cargar_contenido()

    def eventFilter(self, obj, event):
        if obj is self.viewport() and event.type() == event.Wheel:
            if event.modifiers() & Qt.ControlModifier:
                delta = event.angleDelta().y()
                if delta > 0:
                    self.zoom_in()
                elif delta < 0:
                    self.zoom_out()
                return True
        return super().eventFilter(obj, event)

    def _cargar_contenido(self):
        while self._vbox.count():
            item = self._vbox.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()
        self.pixmaps = []

        if self._pm_inicial is not None:
            self._agregar_pagina(self._pm_inicial)
            return

        if _es_imagen(self.ruta):
            pm = QPixmap(self.ruta)
            if pm.isNull():
                self._agregar_mensaje("⚠ No se pudo cargar la imagen")
                return
            self._agregar_pagina(pm)
            return

        if _es_pdf(self.ruta):
            pms = _cargar_pixmaps_pdf(self.ruta)
            if not pms:
                if PYMUPDF_OK:
                    self._agregar_mensaje("⚠ No se pudo cargar el PDF")
                else:
                    self._agregar_mensaje(
                        "Para previsualizar PDFs, instalá pymupdf:\n\n"
                        "pip install pymupdf"
                    )
                return
            for pm in pms:
                self._agregar_pagina(pm)
            return

        self._agregar_mensaje("No se puede previsualizar este tipo de archivo.")

    def _agregar_pagina(self, pm):
        lbl = QLabel()
        lbl.setAlignment(Qt.AlignCenter)
        lbl.setStyleSheet("background: transparent; border: none;")
        lbl.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self._vbox.addWidget(lbl)
        self.pixmaps.append((lbl, pm))

    def _agregar_mensaje(self, texto):
        lbl = QLabel(texto)
        lbl.setAlignment(Qt.AlignCenter)
        lbl.setStyleSheet(
            f"color: {TEXT_MUTED}; font-size: 11pt; background: transparent;"
        )
        lbl.setWordWrap(True)
        self._vbox.addWidget(lbl)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._aplicar_zoom()

    def rotar(self):
        if self.ruta:
            self.rotacion = (self.rotacion + 90) % 360
            _set_rotacion(self.ruta, self.rotacion)
        else:
            self.rotacion = (self.rotacion + 90) % 360
        self._aplicar_zoom()

    def zoom_in(self):
        base = self._zoom_base_actual()
        self.zoom = min(base * 1.15, 8.0)
        self._aplicar_zoom()
        self._notificar_zoom()

    def zoom_out(self):
        base = self._zoom_base_actual()
        self.zoom = max(base / 1.15, 0.1)
        self._aplicar_zoom()
        self._notificar_zoom()

    def zoom_reset(self):
        self.zoom = None
        self.ajustar_al_alto = False
        self._aplicar_zoom()
        self._notificar_zoom()

    def zoom_ajustar_alto(self):
        self.zoom = None
        self.ajustar_al_alto = True
        self._aplicar_zoom()
        self._notificar_zoom()

    def zoom_set(self, valor_porcentaje):
        """Fija el zoom a un valor exacto (en porcentaje, ej: 80 para 80%)."""
        try:
            pct = float(valor_porcentaje)
        except (TypeError, ValueError):
            return
        if pct <= 0:
            return
        self.zoom = pct / 100.0
        self.ajustar_al_alto = False
        self._aplicar_zoom()
        self._notificar_zoom()

    def _zoom_base_actual(self):
        if self.zoom is not None:
            return self.zoom
        if not self.pixmaps:
            return 1.0
        vw = self.viewport().width() - 36
        vh = self.viewport().height() - 36
        if vw <= 20 or vh <= 20:
            return 1.0
        _, pm = self.pixmaps[0]
        pm_rot = pm
        if self.rotacion != 0:
            pm_rot = pm.transformed(
                QTransform().rotate(self.rotacion), Qt.SmoothTransformation
            )
        if pm_rot.width() == 0 or pm_rot.height() == 0:
            return 1.0
        if self.ajustar_al_alto:
            return min(vh / pm_rot.height(), 1.0)
        fw = vw / pm_rot.width()
        fh = vh / pm_rot.height()
        return min(fw, fh, 1.0)

    def _notificar_zoom(self):
        padre = self.window()
        if hasattr(padre, "_actualizar_indicador_zoom"):
            padre._actualizar_indicador_zoom(self)

    def porcentaje_actual(self):
        return int(round(self._zoom_base_actual() * 100))

    def _aplicar_zoom(self):
        if not self.pixmaps:
            return
        vw = self.viewport().width()
        vh = self.viewport().height()
        if vw <= 20 or vh <= 20:
            return

        for lbl, pm in self.pixmaps:
            pm_rot = pm
            if self.rotacion != 0:
                pm_rot = pm.transformed(
                    QTransform().rotate(self.rotacion),
                    Qt.SmoothTransformation,
                )
            if self.zoom is None:
                if self.ajustar_al_alto and pm_rot.height() > 0:
                    factor = (vh - 30) / pm_rot.height()
                    factor = min(factor, 1.0)
                    nuevo_ancho = max(1, int(pm_rot.width() * factor))
                    nuevo_alto = max(1, int(pm_rot.height() * factor))
                    pm_final = pm_rot.scaled(
                        nuevo_ancho, nuevo_alto,
                        Qt.KeepAspectRatio, Qt.SmoothTransformation
                    )
                elif pm_rot.width() <= vw and pm_rot.height() <= vh:
                    pm_final = pm_rot
                else:
                    pm_contain = pm_rot.scaled(
                        vw - 20, vh - 20,
                        Qt.KeepAspectRatio, Qt.SmoothTransformation
                    )
                    if pm_contain.width() < (vw - 20) * 0.4:
                        pm_final = pm_rot
                    else:
                        pm_final = pm_contain
            else:
                nuevo_ancho = max(1, int(pm_rot.width() * self.zoom))
                nuevo_alto = max(1, int(pm_rot.height() * self.zoom))
                pm_final = pm_rot.scaled(
                    nuevo_ancho, nuevo_alto,
                    Qt.KeepAspectRatio, Qt.SmoothTransformation
                )
            lbl.setPixmap(pm_final)
            lbl.setFixedSize(pm_final.size())
            lbl.setScaledContents(False)

        self._cont.adjustSize()
        self._cont.setMinimumSize(self._cont.sizeHint())
        self._cont.updateGeometry()


class ColumnaImagenes(QWidget):
    def __init__(self, rutas, titulo="", color_titulo="#374151", parent=None):
        super().__init__(parent)
        self.rutas = list(rutas)
        self.zoom = None

        v = QVBoxLayout(self)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(6)

        barra = QHBoxLayout()
        barra.setContentsMargins(0, 0, 0, 0)
        barra.setSpacing(6)

        self.lbl_titulo = QLabel(titulo)
        self.lbl_titulo.setStyleSheet(
            f"color: {color_titulo}; font-size: 11pt; font-weight: 700; "
            f"background: transparent; letter-spacing: 0.5px;"
        )
        barra.addWidget(self.lbl_titulo)

        barra.addStretch(1)

        self.btn_zoom_out = QPushButton()
        self.btn_zoom_out.setObjectName("Nav")
        self.btn_zoom_out.setIcon(_make_icon("fa5s.search-minus", "#374151"))
        self.btn_zoom_out.setIconSize(QSize(11, 11))
        self.btn_zoom_out.setCursor(Qt.PointingHandCursor)
        self.btn_zoom_out.setToolTip("Alejar (Ctrl + -)")
        self.btn_zoom_out.clicked.connect(self.zoom_out)
        barra.addWidget(self.btn_zoom_out)

        self.lbl_zoom = QLabel("100%")
        self.lbl_zoom.setAlignment(Qt.AlignCenter)
        self.lbl_zoom.setFixedWidth(48)
        self.lbl_zoom.setStyleSheet(
            f"color: {TEXT_MUTED}; font-size: 8.5pt; font-weight: 700; "
            f"background: transparent;"
        )
        barra.addWidget(self.lbl_zoom)

        self.btn_zoom_in = QPushButton()
        self.btn_zoom_in.setObjectName("Nav")
        self.btn_zoom_in.setIcon(_make_icon("fa5s.search-plus", "#374151"))
        self.btn_zoom_in.setIconSize(QSize(11, 11))
        self.btn_zoom_in.setCursor(Qt.PointingHandCursor)
        self.btn_zoom_in.setToolTip("Acercar (Ctrl + +)")
        self.btn_zoom_in.clicked.connect(self.zoom_in)
        barra.addWidget(self.btn_zoom_in)

        self.btn_zoom_reset = QPushButton("  Ajustar")
        self.btn_zoom_reset.setObjectName("Secondary")
        self.btn_zoom_reset.setIcon(_make_icon("fa5s.expand-arrows-alt", "#374151"))
        self.btn_zoom_reset.setIconSize(QSize(11, 11))
        self.btn_zoom_reset.setCursor(Qt.PointingHandCursor)
        self.btn_zoom_reset.setToolTip("Ajustar al ancho (Ctrl + 0)")
        self.btn_zoom_reset.clicked.connect(self.zoom_reset)
        barra.addWidget(self.btn_zoom_reset)

        v.addLayout(barra)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.NoFrame)
        self.scroll.setStyleSheet(
            f"QScrollArea {{ background-color: {CARD_2}; "
            f"border: 1px solid {BORDER}; border-radius: 8px; }}"
        )
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.scroll.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        self._cont = QWidget()
        self._cont.setStyleSheet("background: transparent;")
        self._cont.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Minimum)

        self._vbox = QVBoxLayout(self._cont)
        self._vbox.setContentsMargins(12, 12, 12, 12)
        self._vbox.setSpacing(14)
        self._vbox.setAlignment(Qt.AlignTop | Qt.AlignHCenter)
        self.scroll.setWidget(self._cont)

        self.pixmaps = []

        if self.rutas:
            for ruta in self.rutas:
                pms = _cargar_todas_las_paginas(ruta)
                if not pms:
                    lbl = QLabel(
                        f"[{os.path.basename(ruta)}]"
                    )
                    lbl.setAlignment(Qt.AlignCenter)
                    lbl.setStyleSheet(
                        f"color: {TEXT_MUTED}; font-size: 10pt; "
                        f"font-style: italic; background: transparent;"
                    )
                    lbl.setWordWrap(True)
                    self._vbox.addWidget(lbl)
                    continue
                for pm in pms:
                    lbl = QLabel()
                    lbl.setAlignment(Qt.AlignCenter)
                    lbl.setStyleSheet("background: transparent; border: none;")
                    lbl.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
                    self._vbox.addWidget(lbl)
                    self.pixmaps.append((lbl, pm))
        else:
            vacio = QLabel("Sin fotos")
            vacio.setAlignment(Qt.AlignCenter)
            vacio.setStyleSheet(
                f"color: {TEXT_MUTED}; font-size: 14pt; font-style: italic; "
                f"background: transparent;"
            )
            self._vbox.addWidget(vacio)

        self.scroll.viewport().installEventFilter(self)
        v.addWidget(self.scroll, 1)

    def eventFilter(self, obj, event):
        if obj is self.scroll.viewport() and event.type() == event.Wheel:
            if event.modifiers() & Qt.ControlModifier:
                delta = event.angleDelta().y()
                if delta > 0:
                    self.zoom_in()
                elif delta < 0:
                    self.zoom_out()
                return True
        return super().eventFilter(obj, event)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._aplicar_zoom()

    def zoom_in(self):
        base = self._zoom_base_actual()
        self.zoom = min(base * 1.15, 8.0)
        self._aplicar_zoom()

    def zoom_out(self):
        base = self._zoom_base_actual()
        self.zoom = max(base / 1.15, 0.1)
        self._aplicar_zoom()

    def zoom_reset(self):
        self.zoom = None
        self._aplicar_zoom()

    def _zoom_base_actual(self):
        if self.zoom is not None:
            return self.zoom
        if not self.pixmaps:
            return 1.0
        vw = self.scroll.viewport().width() - 40
        if vw <= 20:
            return 1.0
        _, pm = self.pixmaps[0]
        if pm.width() == 0:
            return 1.0
        return min(vw / pm.width(), 1.0)

    def _aplicar_zoom(self):
        if not self.pixmaps:
            self.lbl_zoom.setText("100%")
            return

        vw = self.scroll.viewport().width() - 40
        if vw <= 20:
            return

        for lbl, pm in self.pixmaps:
            if pm.isNull():
                continue
            if self.zoom is None:
                ancho = min(vw, pm.width())
                pm_final = pm.scaledToWidth(ancho, Qt.SmoothTransformation)
            else:
                ancho = max(1, int(pm.width() * self.zoom))
                pm_final = pm.scaledToWidth(ancho, Qt.SmoothTransformation)

            lbl.setPixmap(pm_final)
            lbl.setFixedSize(pm_final.size())
            lbl.setScaledContents(False)

        self._cont.adjustSize()
        self._cont.setMinimumSize(self._cont.sizeHint())
        self._cont.updateGeometry()

        pct = int(round(self._zoom_base_actual() * 100))
        self.lbl_zoom.setText(f"{pct}%")


class VistaFotos(QDialog):
    def __init__(self, numero_acta, rutas_care, rutas_came,
                 acta_anterior_fn, acta_siguiente_fn,
                 modo_inicial="grilla", ruta_foco=None, parent=None):
        super().__init__(parent)
        self.setMinimumSize(1000, 720)
        self.setStyleSheet(f"background-color: {BG};")

        self.numero_acta_actual = str(numero_acta)
        self.obtener_care_came_fn = None
        self.obtener_direccion_fn = None
        self.acta_anterior_fn = acta_anterior_fn
        self.acta_siguiente_fn = acta_siguiente_fn
        self.modo = modo_inicial
        self.ruta_foco = ruta_foco

        self.rutas_individual = []
        self.indice_foto_actual = 0

        self.rutas_care = list(rutas_care)
        self.rutas_came = list(rutas_came)

        self.zoom_guardado = None
        self.ajustar_alto_guardado = True

        v = QVBoxLayout(self)
        v.setContentsMargins(16, 16, 16, 16)
        v.setSpacing(12)

        top = QHBoxLayout()
        top.setSpacing(8)

        self.btn_acta_prev = QPushButton()
        self.btn_acta_prev.setObjectName("Nav")
        self.btn_acta_prev.setIcon(_make_icon("fa5s.angle-double-left", "#374151"))
        self.btn_acta_prev.setIconSize(QSize(14, 14))
        self.btn_acta_prev.setCursor(Qt.PointingHandCursor)
        self.btn_acta_prev.setToolTip("Acta anterior")
        self.btn_acta_prev.clicked.connect(self._ir_acta_anterior)
        top.addWidget(self.btn_acta_prev)

        self._titulo_lbl = QLabel("")
        self._titulo_lbl.setStyleSheet(
            f"color: {TEXT}; font-size: 12pt; font-weight: 700;"
        )
        top.addWidget(self._titulo_lbl)

        self.btn_acta_next = QPushButton()
        self.btn_acta_next.setObjectName("Nav")
        self.btn_acta_next.setIcon(_make_icon("fa5s.angle-double-right", "#374151"))
        self.btn_acta_next.setIconSize(QSize(14, 14))
        self.btn_acta_next.setCursor(Qt.PointingHandCursor)
        self.btn_acta_next.setToolTip("Acta siguiente")
        self.btn_acta_next.clicked.connect(self._ir_acta_siguiente)
        top.addWidget(self.btn_acta_next)

        top.addStretch(1)

        self.btn_maps = QPushButton("  Ver en Maps")
        self.btn_maps.setObjectName("Secondary")
        self.btn_maps.setIcon(_make_icon("fa5s.map-marked-alt", "#B91C1C"))
        self.btn_maps.setIconSize(QSize(12, 12))
        self.btn_maps.setCursor(Qt.PointingHandCursor)
        self.btn_maps.setToolTip("Abrir la calle y altura del acta en Google Maps")
        self.btn_maps.clicked.connect(self._abrir_maps)
        self.btn_maps.setVisible(False)
        top.addWidget(self.btn_maps)

        self.btn_modo = QPushButton()
        self.btn_modo.setObjectName("Secondary")
        self.btn_modo.setIconSize(QSize(12, 12))
        self.btn_modo.setCursor(Qt.PointingHandCursor)
        self.btn_modo.clicked.connect(self._toggle_modo)
        top.addWidget(self.btn_modo)

        self.btn_acta_todo = QPushButton("  Acta + fotos")
        self.btn_acta_todo.setObjectName("Secondary")
        self.btn_acta_todo.setIcon(_make_icon("fa5s.columns", "#374151"))
        self.btn_acta_todo.setIconSize(QSize(12, 12))
        self.btn_acta_todo.setCursor(Qt.PointingHandCursor)
        self.btn_acta_todo.setToolTip("Ver el acta PDF junto con CAME y CARE")
        self.btn_acta_todo.clicked.connect(self._toggle_modo_acta)
        top.addWidget(self.btn_acta_todo)

        self.btn_expandir = QPushButton("  Pantalla completa")
        self.btn_expandir.setObjectName("Secondary")
        self.btn_expandir.setIcon(_make_icon("fa5s.expand", "#374151"))
        self.btn_expandir.setIconSize(QSize(12, 12))
        self.btn_expandir.setCursor(Qt.PointingHandCursor)
        self.btn_expandir.clicked.connect(self._toggle_fullscreen)
        top.addWidget(self.btn_expandir)

        btn_cerrar = QPushButton("  Cerrar")
        btn_cerrar.setObjectName("Secondary")
        btn_cerrar.setIcon(_make_icon("fa5s.times", "#374151"))
        btn_cerrar.setIconSize(QSize(12, 12))
        btn_cerrar.setCursor(Qt.PointingHandCursor)
        btn_cerrar.clicked.connect(self.accept)
        top.addWidget(btn_cerrar)

        v.addLayout(top)

        self.stack = QStackedWidget()
        self.stack.setStyleSheet("background: transparent; border: none;")
        v.addWidget(self.stack, 1)

        # ---------- VISTA 1: GRILLA ----------
        self.vista_split = QWidget()
        vs = QHBoxLayout(self.vista_split)
        vs.setContentsMargins(0, 0, 0, 0)
        vs.setSpacing(12)

        self.col_came = None
        self.col_care = None

        vs.addWidget(self._contenedor_columna("CAME", "came"), 1)
        vs.addWidget(self._contenedor_columna("CARE", "care"), 1)

        self.stack.addWidget(self.vista_split)

        # ---------- VISTA 2: INDIVIDUAL ----------
        self.vista_individual = QWidget()
        self.vista_individual.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        vi = QHBoxLayout(self.vista_individual)
        vi.setContentsMargins(0, 0, 0, 0)
        vi.setSpacing(8)

        self.btn_foto_prev = QPushButton()
        self.btn_foto_prev.setObjectName("NavGrande")
        self.btn_foto_prev.setIcon(_make_icon("fa5s.chevron-left", "#374151"))
        self.btn_foto_prev.setIconSize(QSize(24, 24))
        self.btn_foto_prev.setCursor(Qt.PointingHandCursor)
        self.btn_foto_prev.setToolTip("Foto anterior (←)")
        self.btn_foto_prev.clicked.connect(self._ir_foto_anterior)
        self.btn_foto_prev.setFixedWidth(56)
        vi.addWidget(self.btn_foto_prev)

        self.centro = QFrame()
        self.centro.setStyleSheet(
            f"background-color: {CARD}; border: 1px solid {BORDER}; border-radius: 10px;"
        )
        self.centro.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        cl = QVBoxLayout(self.centro)
        cl.setContentsMargins(12, 12, 12, 12)
        cl.setSpacing(8)

        self.img_container = QStackedWidget()
        self.img_container.setStyleSheet("background: transparent; border: none;")
        self.img_container.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        cl.addWidget(self.img_container, 1)

        self._nombre_lbl = QLabel("")
        self._nombre_lbl.setAlignment(Qt.AlignCenter)
        self._nombre_lbl.setStyleSheet(
            f"color: {TEXT_MUTED}; font-size: 9pt; background: transparent;"
        )
        cl.addWidget(self._nombre_lbl)

        acciones = QHBoxLayout()
        acciones.setContentsMargins(0, 0, 0, 0)
        acciones.setSpacing(8)
        acciones.addStretch(1)

        self.btn_rotar_ind = QPushButton("  Rotar")
        self.btn_rotar_ind.setObjectName("Secondary")
        self.btn_rotar_ind.setIcon(_make_icon("fa5s.sync-alt", "#374151"))
        self.btn_rotar_ind.setIconSize(QSize(12, 12))
        self.btn_rotar_ind.setCursor(Qt.PointingHandCursor)
        self.btn_rotar_ind.clicked.connect(self._rotar_actual)
        acciones.addWidget(self.btn_rotar_ind)

        self.btn_abrir_ind = QPushButton("  Abrir con sistema")
        self.btn_abrir_ind.setObjectName("Secondary")
        self.btn_abrir_ind.setIcon(_make_icon("fa5s.external-link-alt", "#374151"))
        self.btn_abrir_ind.setIconSize(QSize(12, 12))
        self.btn_abrir_ind.setCursor(Qt.PointingHandCursor)
        self.btn_abrir_ind.clicked.connect(self._abrir_actual)
        acciones.addWidget(self.btn_abrir_ind)

        acciones.addStretch(1)

        self.btn_zoom_out = QPushButton()
        self.btn_zoom_out.setObjectName("Nav")
        self.btn_zoom_out.setIcon(_make_icon("fa5s.search-minus", "#374151"))
        self.btn_zoom_out.setIconSize(QSize(12, 12))
        self.btn_zoom_out.setCursor(Qt.PointingHandCursor)
        self.btn_zoom_out.setToolTip("Alejar (Ctrl + -)")
        self.btn_zoom_out.clicked.connect(self._zoom_out)
        acciones.addWidget(self.btn_zoom_out)

        self.lbl_zoom = ZoomSpin(
            celda_getter=lambda: self._celda_actual(),
            aplicar_fn=self._actualizar_zoom_individual,
            ancho=60,
        )
        acciones.addWidget(self.lbl_zoom)

        self.btn_zoom_in = QPushButton()
        self.btn_zoom_in.setObjectName("Nav")
        self.btn_zoom_in.setIcon(_make_icon("fa5s.search-plus", "#374151"))
        self.btn_zoom_in.setIconSize(QSize(12, 12))
        self.btn_zoom_in.setCursor(Qt.PointingHandCursor)
        self.btn_zoom_in.setToolTip("Acercar (Ctrl + +)")
        self.btn_zoom_in.clicked.connect(self._zoom_in)
        acciones.addWidget(self.btn_zoom_in)

        self.btn_zoom_reset = QPushButton("  Ajustar")
        self.btn_zoom_reset.setObjectName("Secondary")
        self.btn_zoom_reset.setIcon(_make_icon("fa5s.expand-arrows-alt", "#374151"))
        self.btn_zoom_reset.setIconSize(QSize(12, 12))
        self.btn_zoom_reset.setCursor(Qt.PointingHandCursor)
        self.btn_zoom_reset.setToolTip("Ajustar a ventana (Ctrl + 0)")
        self.btn_zoom_reset.clicked.connect(self._zoom_reset)
        acciones.addWidget(self.btn_zoom_reset)

        acciones.addStretch(1)
        cl.addLayout(acciones)

        vi.addWidget(self.centro, 1)

        self.btn_foto_next = QPushButton()
        self.btn_foto_next.setObjectName("NavGrande")
        self.btn_foto_next.setIcon(_make_icon("fa5s.chevron-right", "#374151"))
        self.btn_foto_next.setIconSize(QSize(24, 24))
        self.btn_foto_next.setCursor(Qt.PointingHandCursor)
        self.btn_foto_next.setToolTip("Foto siguiente (→)")
        self.btn_foto_next.clicked.connect(self._ir_foto_siguiente)
        self.btn_foto_next.setFixedWidth(56)
        vi.addWidget(self.btn_foto_next)

        self.stack.addWidget(self.vista_individual)

        # ---------- VISTA 3: ACTA + CAME + CARE ----------
        self._build_vista_acta_todo()
        self.stack.addWidget(self.vista_acta_todo)

        self._refrescar()

    def _build_vista_acta_todo(self):
        self.vista_acta_todo = QWidget()
        vat = QVBoxLayout(self.vista_acta_todo)
        vat.setContentsMargins(0, 0, 0, 0)
        vat.setSpacing(0)

        self.splitter_acta_todo = QSplitter(Qt.Horizontal)
        self.splitter_acta_todo.setHandleWidth(10)
        vat.addWidget(self.splitter_acta_todo)

        # -------- Columna 1: ACTA (destacada) --------
        self.col_acta_frame = QFrame()
        self.col_acta_frame.setObjectName("ColActaDestacada")
        self.col_acta_frame.setMinimumWidth(400)
        self.col_acta_frame.setStyleSheet(f"""
            QFrame#ColActaDestacada {{
                background-color: {ACCENT_BG};
                border: 2px solid {ACCENT};
                border-radius: 10px;
            }}
        """)
        cla = QVBoxLayout(self.col_acta_frame)
        cla.setContentsMargins(10, 10, 10, 10)
        cla.setSpacing(8)

        header_acta = QHBoxLayout()
        header_acta.setContentsMargins(4, 0, 4, 0)
        header_acta.setSpacing(8)

        ico_acta = QLabel()
        ico_acta.setPixmap(_make_icon("fa5s.file-pdf", ACCENT).pixmap(16, 16))
        ico_acta.setStyleSheet("background: transparent;")
        header_acta.addWidget(ico_acta)

        lbl_acta = QLabel("ACTA")
        lbl_acta.setStyleSheet(
            f"color: {ACCENT}; font-size: 12pt; font-weight: 800; "
            f"background: transparent; letter-spacing: 1px;"
        )
        header_acta.addWidget(lbl_acta)

        self.lbl_acta_todo_num = QLabel("")
        self.lbl_acta_todo_num.setStyleSheet(
            f"color: {ACCENT}; font-size: 10pt; font-weight: 600; "
            f"background: transparent;"
        )
        header_acta.addWidget(self.lbl_acta_todo_num)

        header_acta.addStretch(1)

        self.btn_zoom_out_acta_todo = QPushButton()
        self.btn_zoom_out_acta_todo.setObjectName("Nav")
        self.btn_zoom_out_acta_todo.setIcon(_make_icon("fa5s.search-minus", "#374151"))
        self.btn_zoom_out_acta_todo.setIconSize(QSize(11, 11))
        self.btn_zoom_out_acta_todo.setCursor(Qt.PointingHandCursor)
        self.btn_zoom_out_acta_todo.setToolTip("Alejar (Ctrl + -)")
        self.btn_zoom_out_acta_todo.clicked.connect(self._zoom_out_acta_todo)
        header_acta.addWidget(self.btn_zoom_out_acta_todo)

        self.lbl_zoom_acta_todo = ZoomSpin(
            celda_getter=lambda: self.acta_visor,
            aplicar_fn=self._actualizar_zoom_acta_todo,
            ancho=55,
        )
        header_acta.addWidget(self.lbl_zoom_acta_todo)

        self.btn_zoom_in_acta_todo = QPushButton()
        self.btn_zoom_in_acta_todo.setObjectName("Nav")
        self.btn_zoom_in_acta_todo.setIcon(_make_icon("fa5s.search-plus", "#374151"))
        self.btn_zoom_in_acta_todo.setIconSize(QSize(11, 11))
        self.btn_zoom_in_acta_todo.setCursor(Qt.PointingHandCursor)
        self.btn_zoom_in_acta_todo.setToolTip("Acercar (Ctrl + +)")
        self.btn_zoom_in_acta_todo.clicked.connect(self._zoom_in_acta_todo)
        header_acta.addWidget(self.btn_zoom_in_acta_todo)

        self.btn_ajustar_alto = QPushButton("  Alto")
        self.btn_ajustar_alto.setObjectName("Secondary")
        self.btn_ajustar_alto.setIcon(_make_icon("fa5s.arrows-alt-v", "#374151"))
        self.btn_ajustar_alto.setIconSize(QSize(11, 11))
        self.btn_ajustar_alto.setCursor(Qt.PointingHandCursor)
        self.btn_ajustar_alto.setToolTip("Ajustar la imagen a la altura del visor")
        self.btn_ajustar_alto.clicked.connect(self._ajustar_alto_acta_todo)
        header_acta.addWidget(self.btn_ajustar_alto)

        self.btn_ajustar_ancho = QPushButton("  Ajustar")
        self.btn_ajustar_ancho.setObjectName("Secondary")
        self.btn_ajustar_ancho.setIcon(_make_icon("fa5s.expand-arrows-alt", "#374151"))
        self.btn_ajustar_ancho.setIconSize(QSize(11, 11))
        self.btn_ajustar_ancho.setCursor(Qt.PointingHandCursor)
        self.btn_ajustar_ancho.setToolTip("Ajustar la imagen al ancho del visor (Ctrl + 0)")
        self.btn_ajustar_ancho.clicked.connect(self._zoom_reset_acta_todo)
        header_acta.addWidget(self.btn_ajustar_ancho)

        cla.addLayout(header_acta)

        self.acta_visor = CeldaFoto(QPixmap(), ajustar_al_alto=True)
        self.acta_visor.setStyleSheet(
            f"QScrollArea {{ background-color: {CARD}; "
            f"border: 1px solid {BORDER}; border-radius: 8px; }}"
        )
        cla.addWidget(self.acta_visor, 1)

        self.splitter_acta_todo.addWidget(self.col_acta_frame)

        frame_came = self._build_columna_foto(
            titulo="CAME", color="#B45309", icono="fa5s.camera"
        )
        flc = frame_came.layout()
        self.col_came_todo = ColumnaImagenes([], titulo="", color_titulo="#B45309")
        self.col_came_todo.lbl_titulo.setVisible(False)
        self.col_came_todo.btn_zoom_out.setVisible(False)
        self.col_came_todo.lbl_zoom.setVisible(False)
        self.col_came_todo.btn_zoom_in.setVisible(False)
        self.col_came_todo.btn_zoom_reset.setVisible(False)
        flc.addWidget(self.col_came_todo, 1)
        self.splitter_acta_todo.addWidget(frame_came)

        frame_care = self._build_columna_foto(
            titulo="CARE", color=ACCENT, icono="fa5s.camera"
        )
        flr = frame_care.layout()
        self.col_care_todo = ColumnaImagenes([], titulo="", color_titulo=ACCENT)
        self.col_care_todo.lbl_titulo.setVisible(False)
        self.col_care_todo.btn_zoom_out.setVisible(False)
        self.col_care_todo.lbl_zoom.setVisible(False)
        self.col_care_todo.btn_zoom_in.setVisible(False)
        self.col_care_todo.btn_zoom_reset.setVisible(False)
        flr.addWidget(self.col_care_todo, 1)
        self.splitter_acta_todo.addWidget(frame_care)

        self.splitter_acta_todo.setSizes([600, 350, 350])
        self.splitter_acta_todo.setStretchFactor(0, 5)
        self.splitter_acta_todo.setStretchFactor(1, 3)
        self.splitter_acta_todo.setStretchFactor(2, 3)

        self.splitter_acta_todo.setCollapsible(0, False)
        self.splitter_acta_todo.setCollapsible(1, False)
        self.splitter_acta_todo.setCollapsible(2, False)

    def _build_columna_foto(self, titulo, color, icono="fa5s.camera"):
        frame = QFrame()
        frame.setStyleSheet(f"""
            QFrame {{
                background-color: {CARD};
                border: 1px solid {BORDER};
                border-radius: 10px;
            }}
        """)
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)

        header = QHBoxLayout()
        header.setContentsMargins(4, 0, 4, 0)
        header.setSpacing(8)

        ico = QLabel()
        ico.setPixmap(_make_icon(icono, color).pixmap(15, 15))
        ico.setStyleSheet("background: transparent;")
        header.addWidget(ico)

        lbl = QLabel(titulo)
        lbl.setStyleSheet(
            f"color: {color}; font-size: 12pt; font-weight: 800; "
            f"background: transparent; letter-spacing: 1px;"
        )
        header.addWidget(lbl)

        header.addStretch(1)

        layout.addLayout(header)

        return frame

    def _zoom_in_acta_todo(self):
        self.acta_visor.zoom_in()
        self._actualizar_zoom_acta_todo()

    def _zoom_out_acta_todo(self):
        self.acta_visor.zoom_out()
        self._actualizar_zoom_acta_todo()

    def _zoom_reset_acta_todo(self):
        self.acta_visor.zoom_reset()
        self._actualizar_zoom_acta_todo()

    def _ajustar_alto_acta_todo(self):
        self.acta_visor.zoom_ajustar_alto()
        self._actualizar_zoom_acta_todo()

    def _actualizar_zoom_acta_todo(self):
        try:
            self.lbl_zoom_acta_todo.actualizar_texto()
        except Exception:
            pass

    def _actualizar_zoom_individual(self):
        try:
            self.lbl_zoom.actualizar_texto()
        except Exception:
            pass

    def _guardar_zoom_actual(self):
        if self.acta_visor is not None and self.acta_visor.pixmaps:
            self.zoom_guardado = self.acta_visor.zoom
            self.ajustar_alto_guardado = self.acta_visor.ajustar_al_alto

    def _contenedor_columna(self, titulo, tipo):
        frame = QFrame()
        frame.setStyleSheet(
            f"background-color: {CARD}; border: 1px solid {BORDER}; "
            f"border-radius: 10px;"
        )
        fl = QVBoxLayout(frame)
        fl.setContentsMargins(12, 12, 12, 12)
        fl.setSpacing(8)

        if tipo == "came":
            color = "#B45309"
        else:
            color = ACCENT

        col = ColumnaImagenes([], titulo=titulo, color_titulo=color)
        fl.addWidget(col, 1)

        if tipo == "came":
            self.col_came = col
        else:
            self.col_care = col

        return frame

    def set_obtener_care_came_fn(self, fn):
        self.obtener_care_came_fn = fn

    def set_obtener_direccion_fn(self, fn):
        self.obtener_direccion_fn = fn
        self._actualizar_boton_maps()

    def _actualizar_boton_maps(self):
        calle, altura = (None, None)
        if self.obtener_direccion_fn is not None:
            try:
                calle, altura = self.obtener_direccion_fn(self.numero_acta_actual)
            except Exception:
                calle, altura = (None, None)

        if calle:
            self.btn_maps.setVisible(True)
            txt = str(calle)
            if altura:
                txt = f"{txt} {altura}"
            self.btn_maps.setToolTip(f"Ver en Google Maps: {txt}, CABA")
        else:
            self.btn_maps.setVisible(False)

    def _abrir_maps(self):
        calle, altura = (None, None)
        if self.obtener_direccion_fn is not None:
            try:
                calle, altura = self.obtener_direccion_fn(self.numero_acta_actual)
            except Exception:
                calle, altura = (None, None)

        if not calle:
            QMessageBox.information(
                self, "Sin dirección",
                "No se pudo determinar la calle ni la altura para esta acta."
            )
            return

        ok = _abrir_en_maps(calle, altura)
        if not ok:
            QMessageBox.warning(
                self, "No se pudo abrir Maps",
                "No se pudo abrir el navegador con Google Maps."
            )

    def _ir_acta_anterior(self):
        ant = self.acta_anterior_fn(self.numero_acta_actual)
        if ant is not None:
            self._guardar_zoom_actual()
            self.numero_acta_actual = str(ant)
            self.indice_foto_actual = 0
            self.ruta_foco = None
            self._refrescar()

    def _ir_acta_siguiente(self):
        sig = self.acta_siguiente_fn(self.numero_acta_actual)
        if sig is not None:
            self._guardar_zoom_actual()
            self.numero_acta_actual = str(sig)
            self.indice_foto_actual = 0
            self.ruta_foco = None
            self._refrescar()

    def _ir_foto_anterior(self):
        if self.indice_foto_actual > 0:
            self.indice_foto_actual -= 1
            self._cargar_individual()

    def _ir_foto_siguiente(self):
        if self.indice_foto_actual < len(self.rutas_individual) - 1:
            self.indice_foto_actual += 1
            self._cargar_individual()

    def _toggle_modo(self):
        if self.modo == "individual":
            self.modo = "grilla"
        else:
            self.modo = "individual"
            self.rutas_individual = list(self.rutas_came) + list(self.rutas_care)
            self.indice_foto_actual = 0
        self._refrescar()

    def _toggle_modo_acta(self):
        if self.modo == "acta_todo":
            self.modo = "grilla"
            self.btn_acta_todo.setText("  Acta + fotos")
        else:
            self.modo = "acta_todo"
            self.btn_acta_todo.setText("  Ocultar acta")
        self._refrescar()

    def _actualizar_boton_modo(self):
        if self.modo == "grilla":
            self.btn_modo.setText("  Ver una por una")
            self.btn_modo.setIcon(_make_icon("fa5s.image", "#374151"))
        elif self.modo == "individual":
            self.btn_modo.setText("  Ver todas juntas")
            self.btn_modo.setIcon(_make_icon("fa5s.th", "#374151"))
        else:
            self.btn_modo.setText("  Ver una por una")
            self.btn_modo.setIcon(_make_icon("fa5s.image", "#374151"))

    def _refrescar(self):
        if self.obtener_care_came_fn is not None:
            try:
                care, came = self.obtener_care_came_fn(self.numero_acta_actual)
                self.rutas_care = list(care or [])
                self.rutas_came = list(came or [])
            except Exception:
                pass

        self._actualizar_boton_modo()
        self._actualizar_boton_maps()

        self._titulo_lbl.setText(
            f"ACTA {self.numero_acta_actual}  ·  "
            f"CAME: {len(self.rutas_came)}  ·  CARE: {len(self.rutas_care)}"
        )
        try:
            self.lbl_acta_todo_num.setText(f"· N° {self.numero_acta_actual}")
        except Exception:
            pass

        self.btn_acta_prev.setEnabled(
            self.acta_anterior_fn(self.numero_acta_actual) is not None
        )
        self.btn_acta_next.setEnabled(
            self.acta_siguiente_fn(self.numero_acta_actual) is not None
        )

        if self.modo == "grilla":
            self.stack.setCurrentWidget(self.vista_split)
            self._recargar_columnas()
        elif self.modo == "individual":
            self.stack.setCurrentWidget(self.vista_individual)
            if not self.rutas_individual:
                self.rutas_individual = list(self.rutas_came) + list(self.rutas_care)
            if self.indice_foto_actual >= len(self.rutas_individual):
                self.indice_foto_actual = 0
            self._cargar_individual()
        elif self.modo == "acta_todo":
            self.stack.setCurrentWidget(self.vista_acta_todo)
            self._cargar_modo_acta_todo()

    def _recargar_columnas(self):
        layout_split = self.vista_split.layout()

        while layout_split.count():
            item = layout_split.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()

        self.col_came = None
        self.col_care = None

        layout_split.addWidget(self._contenedor_columna("CAME", "came"), 1)
        layout_split.addWidget(self._contenedor_columna("CARE", "care"), 1)

        self._set_rutas_columna(self.col_came, self.rutas_came)
        self._set_rutas_columna(self.col_care, self.rutas_care)

    def _set_rutas_columna(self, col, rutas):
        col.rutas = list(rutas)
        while col._vbox.count():
            item = col._vbox.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()
        col.pixmaps = []

        if not rutas:
            vacio = QLabel("Sin fotos")
            vacio.setAlignment(Qt.AlignCenter)
            vacio.setStyleSheet(
                f"color: {TEXT_MUTED}; font-size: 14pt; font-style: italic; "
                f"background: transparent;"
            )
            col._vbox.addWidget(vacio)
            col.lbl_zoom.setText("—")
            return

        for ruta in rutas:
            pms = _cargar_todas_las_paginas(ruta)
            if not pms:
                lbl = QLabel(f"[{os.path.basename(ruta)}]")
                lbl.setAlignment(Qt.AlignCenter)
                lbl.setStyleSheet(
                    f"color: {TEXT_MUTED}; font-size: 10pt; "
                    f"font-style: italic; background: transparent;"
                )
                lbl.setWordWrap(True)
                col._vbox.addWidget(lbl)
                continue
            for pm in pms:
                lbl = QLabel()
                lbl.setAlignment(Qt.AlignCenter)
                lbl.setStyleSheet("background: transparent; border: none;")
                lbl.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
                col._vbox.addWidget(lbl)
                col.pixmaps.append((lbl, pm))

        QTimer.singleShot(0, col._aplicar_zoom)

    def _cargar_modo_acta_todo(self):
        while self.acta_visor._vbox.count():
            item = self.acta_visor._vbox.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()
        self.acta_visor.pixmaps = []

        if self.zoom_guardado is not None:
            self.acta_visor.zoom = self.zoom_guardado
            self.acta_visor.ajustar_al_alto = False
        else:
            self.acta_visor.zoom = None
            self.acta_visor.ajustar_al_alto = self.ajustar_alto_guardado

        ruta_pdf = None
        num_pagina = 1

        parent = self.parent()
        if parent is None:
            parent = self.window()

        if parent is not None and hasattr(parent, "indice_pdfs") and parent.indice_pdfs is not None:
            numero = self.numero_acta_actual
            if numero in parent.indice_pdfs["por_nro"]:
                id_norm = parent.indice_pdfs["por_nro"][numero]
                entry = parent.indice_pdfs["por_id"].get(id_norm)
                if entry:
                    ruta_pdf = entry.get("ruta")
                    datos = entry.get("datos") or {}
                    num_pagina = datos.get("pagina", 1)

        if ruta_pdf and os.path.exists(ruta_pdf) and PYMUPDF_OK:
            pm = self._renderizar_pagina_pdf_local(ruta_pdf, num_pagina)
            if pm is not None:
                self.acta_visor._agregar_pagina(pm)
                QTimer.singleShot(80, self.acta_visor._aplicar_zoom)
            else:
                self.acta_visor._agregar_mensaje(
                    "⚠ No se pudo renderizar la página del PDF."
                )
        else:
            self.acta_visor._agregar_mensaje(
                "⚠ No se encontró el PDF del acta."
            )

        self._actualizar_zoom_acta_todo()

        self._set_rutas_columna(self.col_came_todo, self.rutas_came)
        self._set_rutas_columna(self.col_care_todo, self.rutas_care)

    def _renderizar_pagina_pdf_local(self, ruta_pdf, num_pagina):
        if not PYMUPDF_OK:
            return None
        try:
            doc = fitz.open(ruta_pdf)
            idx = max(0, int(num_pagina) - 1)
            if idx >= doc.page_count:
                doc.close()
                return None
            pagina = doc.load_page(idx)
            mat = fitz.Matrix(2.0, 2.0)
            pix = pagina.get_pixmap(matrix=mat, alpha=False)
            img = QImage(
                pix.samples,
                pix.width,
                pix.height,
                pix.stride,
                QImage.Format_RGB888,
            )
            pm = QPixmap.fromImage(img.copy())
            doc.close()
            return pm
        except Exception:
            return None

    def _limpiar_img_container(self):
        while self.img_container.count() > 0:
            w = self.img_container.widget(0)
            self.img_container.removeWidget(w)
            w.deleteLater()

    def _cargar_individual(self):
        self._limpiar_img_container()

        if not self.rutas_individual:
            vacio = QLabel("No hay documentos cargados.")
            vacio.setAlignment(Qt.AlignCenter)
            vacio.setStyleSheet(
                f"color: {TEXT_MUTED}; font-size: 12pt; font-style: italic; "
                f"background: transparent;"
            )
            self.img_container.addWidget(vacio)
            self._nombre_lbl.setText("")
            self.btn_foto_prev.setEnabled(False)
            self.btn_foto_next.setEnabled(False)
            self.btn_rotar_ind.setEnabled(False)
            self.btn_abrir_ind.setEnabled(False)
            return

        if self.ruta_foco is not None and self.ruta_foco in self.rutas_individual:
            self.indice_foto_actual = self.rutas_individual.index(self.ruta_foco)
            self.ruta_foco = None

        if self.indice_foto_actual >= len(self.rutas_individual):
            self.indice_foto_actual = 0

        ruta_actual = self.rutas_individual[self.indice_foto_actual]
        self._nombre_lbl.setText(os.path.basename(ruta_actual))

        celda = CeldaFoto(ruta_actual)
        celda.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.img_container.addWidget(celda)
        self.img_container.setCurrentWidget(celda)

        QTimer.singleShot(0, self._actualizar_zoom_individual)

        self.btn_foto_prev.setEnabled(self.indice_foto_actual > 0)
        self.btn_foto_next.setEnabled(
            self.indice_foto_actual < len(self.rutas_individual) - 1
        )

        previsualizable = _es_imagen(ruta_actual) or (
            _es_pdf(ruta_actual) and PYMUPDF_OK
        )
        self.btn_rotar_ind.setEnabled(previsualizable)
        self.btn_abrir_ind.setEnabled(True)

    def _rotar_actual(self):
        if self.modo == "individual":
            w = self.img_container.currentWidget()
            if isinstance(w, CeldaFoto):
                w.rotar()

    def _abrir_actual(self):
        if not self.rutas_individual:
            return
        if self.indice_foto_actual >= len(self.rutas_individual):
            return
        ruta = self.rutas_individual[self.indice_foto_actual]
        ok = _abrir_con_sistema(ruta)
        if not ok:
            QMessageBox.warning(
                self,
                "No se pudo abrir",
                f"No se pudo abrir el archivo:\n{ruta}\n\n"
                f"Asegurate de tener una aplicación asociada a este tipo de archivo."
            )

    def _celda_actual(self):
        if self.modo != "individual":
            return None
        w = self.img_container.currentWidget()
        if isinstance(w, CeldaFoto):
            return w
        return None

    def _zoom_in(self):
        c = self._celda_actual()
        if c is not None:
            c.zoom_in()
            self._actualizar_zoom_individual()

    def _zoom_out(self):
        c = self._celda_actual()
        if c is not None:
            c.zoom_out()
            self._actualizar_zoom_individual()

    def _zoom_reset(self):
        c = self._celda_actual()
        if c is not None:
            c.zoom_reset()
            self._actualizar_zoom_individual()

    def _actualizar_indicador_zoom(self, celda):
        if celda is self._celda_actual():
            self._actualizar_zoom_individual()
        elif celda is self.acta_visor:
            self._actualizar_zoom_acta_todo()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            if self.isFullScreen():
                self._toggle_fullscreen()
            else:
                self.accept()
        elif event.key() == Qt.Key_Left:
            if self.modo == "individual":
                self._ir_foto_anterior()
        elif event.key() == Qt.Key_Right:
            if self.modo == "individual":
                self._ir_foto_siguiente()
        elif event.key() in (Qt.Key_Plus, Qt.Key_Equal):
            self._zoom_in()
        elif event.key() in (Qt.Key_Minus, Qt.Key_Underscore):
            self._zoom_out()
        elif event.key() == Qt.Key_0:
            self._zoom_reset()
        else:
            super().keyPressEvent(event)

    def _toggle_fullscreen(self):
        if self.isFullScreen():
            self.showNormal()
            self.btn_expandir.setText("  Pantalla completa")
            self.btn_expandir.setIcon(_make_icon("fa5s.expand", "#374151"))
        else:
            self.showFullScreen()
            self.btn_expandir.setText("  Salir de pantalla completa")
            self.btn_expandir.setIcon(_make_icon("fa5s.compress", "#374151"))


class ThumbnailFoto(QFrame):
    def __init__(self, ruta_archivo, tamanio=(220, 180), callback_toggle=None, parent=None):
        super().__init__(parent)
        self.ruta_archivo = ruta_archivo
        self.callback_toggle = callback_toggle
        self.setFixedSize(tamanio[0], tamanio[1])
        self.setStyleSheet(f"""
            ThumbnailFoto {{
                background-color: {CARD};
                border: 1px solid {BORDER};
                border-radius: 8px;
            }}
            ThumbnailFoto:hover {{
                border: 2px solid {ACCENT};
            }}
        """)
        self.setCursor(Qt.PointingHandCursor)

        v = QVBoxLayout(self)
        v.setContentsMargins(6, 6, 6, 6)
        v.setSpacing(4)

        top_row = QHBoxLayout()
        top_row.setContentsMargins(0, 0, 0, 0)
        top_row.setSpacing(4)

        self.chk = QCheckBox()
        self.chk.setCursor(Qt.PointingHandCursor)
        self.chk.setStyleSheet(f"""
            QCheckBox {{ background: transparent; }}
            QCheckBox::indicator {{
                width: 18px; height: 18px;
                border-radius: 4px;
                border: 2px solid {ACCENT};
                background: {CARD};
            }}
            QCheckBox::indicator:checked {{
                background-color: {ACCENT};
                border-color: {ACCENT};
                image: none;
            }}
            QCheckBox::indicator:hover {{ border-color: {ACCENT_H}; }}
        """)
        self.chk.stateChanged.connect(self._on_checkbox_change)
        top_row.addWidget(self.chk, 0, Qt.AlignTop)
        top_row.addStretch(1)
        v.addLayout(top_row)

        lbl_img = QLabel()
        lbl_img.setAlignment(Qt.AlignCenter)
        lbl_img.setStyleSheet("background: transparent; border: none;")

        pm = _cargar_pixmap(ruta_archivo)
        if pm is not None and not pm.isNull():
            pm = pm.scaled(
                tamanio[0] - 30, tamanio[1] - 60,
                Qt.KeepAspectRatio, Qt.SmoothTransformation
            )
            lbl_img.setPixmap(pm)
        else:
            nombre_ico, color = _icono_para_archivo(ruta_archivo)
            icono = _make_icon(nombre_ico, color)
            tam_ico = min(tamanio[0] - 40, tamanio[1] - 70)
            lbl_img.setPixmap(icono.pixmap(tam_ico, tam_ico))

        v.addWidget(lbl_img, 1)

        nombre = os.path.basename(ruta_archivo)
        if len(nombre) > 28:
            nombre = nombre[:25] + "..."
        lbl_nombre = QLabel(nombre)
        lbl_nombre.setAlignment(Qt.AlignCenter)
        lbl_nombre.setStyleSheet(
            f"color: {TEXT_MUTED}; font-size: 7.5pt; background: transparent; "
            f"border: none;"
        )
        lbl_nombre.setWordWrap(True)
        v.addWidget(lbl_nombre)

    def _on_checkbox_change(self, state):
        if self.callback_toggle:
            self.callback_toggle(self.ruta_archivo, state == Qt.Checked)

    def is_checked(self):
        return self.chk.isChecked()

    def set_checked_silencioso(self, checked):
        self.chk.blockSignals(True)
        self.chk.setChecked(checked)
        self.chk.blockSignals(False)

    def mousePressEvent(self, event):
        if self.chk.geometry().contains(event.pos()):
            return
        if event.button() == Qt.LeftButton:
            parent = self.window()
            if hasattr(parent, "_abrir_vista_foto_en_indice"):
                nro = None
                for n, datos in parent.indice_fotos.items():
                    todas = list(datos.get("care", [])) + list(datos.get("came", []))
                    if self.ruta_archivo in todas:
                        nro = n
                        break
                if nro is not None:
                    parent._abrir_vista_foto_en_indice(
                        nro, modo="individual", ruta_foco=self.ruta_archivo
                    )
                    return


def _calcular_columnas(cantidad):
    if cantidad <= 0:
        return 1
    if cantidad == 1:
        return 1
    if cantidad == 2:
        return 2
    if cantidad == 3:
        return 3
    if cantidad == 4:
        return 2
    if cantidad in (5, 6):
        return 3
    return 3


class GaleriaFotos(QWidget):
    def __init__(self, rutas_archivos, parent=None, direccion_fn=None):
        super().__init__(parent)
        self.rutas = rutas_archivos
        self.seleccionadas = set()
        self.direccion_fn = direccion_fn

        v = QVBoxLayout(self)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(10)

        self.barra_top = QWidget()
        self.barra_top.setStyleSheet("background: transparent;")
        bt = QHBoxLayout(self.barra_top)
        bt.setContentsMargins(0, 0, 0, 0)
        bt.setSpacing(8)

        self.btn_toggle_todas = QPushButton("  Seleccionar todas")
        self.btn_toggle_todas.setObjectName("Secondary")
        self.btn_toggle_todas.setIcon(_make_icon("fa5s.check-square", "#374151"))
        self.btn_toggle_todas.setIconSize(QSize(11, 11))
        self.btn_toggle_todas.setCursor(Qt.PointingHandCursor)
        self.btn_toggle_todas.clicked.connect(self._toggle_todas)
        bt.addWidget(self.btn_toggle_todas)

        self.btn_maps = QPushButton("  Ver en Maps")
        self.btn_maps.setObjectName("Secondary")
        self.btn_maps.setIcon(_make_icon("fa5s.map-marked-alt", "#B91C1C"))
        self.btn_maps.setIconSize(QSize(12, 12))
        self.btn_maps.setCursor(Qt.PointingHandCursor)
        self.btn_maps.setToolTip("Abrir la calle y altura del acta en Google Maps")
        self.btn_maps.clicked.connect(self._abrir_maps)
        self.btn_maps.setVisible(False)
        bt.addWidget(self.btn_maps)

        bt.addStretch(1)

        self.btn_ver_sel = QPushButton("  Ver seleccionadas (0)")
        self.btn_ver_sel.setObjectName("Secondary")
        self.btn_ver_sel.setIcon(_make_icon("fa5s.images", "#374151"))
        self.btn_ver_sel.setIconSize(QSize(12, 12))
        self.btn_ver_sel.setCursor(Qt.PointingHandCursor)
        self.btn_ver_sel.clicked.connect(self._ver_seleccionadas)
        self.btn_ver_sel.setVisible(False)
        bt.addWidget(self.btn_ver_sel)

        v.addWidget(self.barra_top)

        self._actualizar_boton_maps()

        n = len(rutas_archivos)
        cols = _calcular_columnas(n)

        if cols == 1:
            tam = (520, 400)
        elif cols == 2:
            tam = (280, 220)
        else:
            tam = (200, 170)

        self.grid_widget = QWidget()
        self.grid_widget.setStyleSheet("background: transparent;")
        grid = QGridLayout(self.grid_widget)
        grid.setContentsMargins(0, 0, 0, 0)
        grid.setSpacing(10)

        self.thumbnails = []
        for i, ruta in enumerate(rutas_archivos):
            fila = i // cols
            col = i % cols
            thumb = ThumbnailFoto(
                ruta, tamanio=tam, callback_toggle=self._on_toggle,
            )
            grid.addWidget(thumb, fila, col)
            self.thumbnails.append(thumb)

        if cols == 1:
            grid.setAlignment(Qt.AlignHCenter)

        v.addWidget(self.grid_widget)

    def _actualizar_boton_maps(self):
        calle, altura = (None, None)
        if self.direccion_fn is not None:
            try:
                calle, altura = self.direccion_fn()
            except Exception:
                calle, altura = (None, None)

        if calle:
            self.btn_maps.setVisible(True)
            txt = str(calle)
            if altura:
                txt = f"{txt} {altura}"
            self.btn_maps.setToolTip(f"Ver en Google Maps: {txt}, CABA")
        else:
            self.btn_maps.setVisible(False)

    def _abrir_maps(self):
        calle, altura = (None, None)
        if self.direccion_fn is not None:
            try:
                calle, altura = self.direccion_fn()
            except Exception:
                calle, altura = (None, None)

        if not calle:
            QMessageBox.information(
                self, "Sin dirección",
                "No se pudo determinar la calle ni la altura para esta acta."
            )
            return

        ok = _abrir_en_maps(calle, altura)
        if not ok:
            QMessageBox.warning(
                self, "No se pudo abrir Maps",
                "No se pudo abrir el navegador con Google Maps."
            )

    def _on_toggle(self, ruta, checked):
        if checked:
            self.seleccionadas.add(ruta)
        else:
            self.seleccionadas.discard(ruta)
        self._actualizar_botones()

    def _actualizar_botones(self):
        n = len(self.seleccionadas)
        self.btn_ver_sel.setText(f"  Ver seleccionadas ({n})")
        self.btn_ver_sel.setVisible(n > 0)
        total = len(self.thumbnails)
        if total > 0 and n == total:
            self.btn_toggle_todas.setText("  Deseleccionar todas")
            self.btn_toggle_todas.setIcon(_make_icon("fa5s.square", "#374151"))
        else:
            self.btn_toggle_todas.setText("  Seleccionar todas")
            self.btn_toggle_todas.setIcon(_make_icon("fa5s.check-square", "#374151"))

    def _toggle_todas(self):
        total = len(self.thumbnails)
        if total == 0:
            return
        if len(self.seleccionadas) == total:
            for t in self.thumbnails:
                t.set_checked_silencioso(False)
            self.seleccionadas.clear()
        else:
            for t in self.thumbnails:
                t.set_checked_silencioso(True)
                self.seleccionadas.add(t.ruta_archivo)
        self._actualizar_botones()

    def _ver_seleccionadas(self):
        if not self.seleccionadas:
            return
        parent = self.window()
        nro = None
        primera = None
        if hasattr(parent, "indice_fotos"):
            for n, datos in parent.indice_fotos.items():
                todas = list(datos.get("care", [])) + list(datos.get("came", []))
                for r in todas:
                    if r in self.seleccionadas:
                        nro = n
                        primera = r
                        break
                if nro is not None:
                    break
        if nro is not None and hasattr(parent, "_abrir_vista_foto_en_indice"):
            parent._abrir_vista_foto_en_indice(
                nro, modo="individual", ruta_foco=primera
            )


class DialogoErrores(QDialog):
    def __init__(self, actas_error, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Actas con errores")
        self.setMinimumWidth(480)
        self.setMinimumHeight(520)
        self.numeros_seleccionados = None

        v = QVBoxLayout(self)
        v.setContentsMargins(24, 24, 24, 20)
        v.setSpacing(14)

        titulo = QLabel(f"{len(actas_error)} actas con errores detectadas")
        titulo.setStyleSheet(f"color: {TEXT}; font-size: 12pt; font-weight: 700;")
        v.addWidget(titulo)

        sub = QLabel("Marcá las que querés ver en la lista principal")
        sub.setStyleSheet(f"color: {TEXT_MUTED}; font-size: 9.5pt;")
        v.addWidget(sub)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setStyleSheet(f"""
            QScrollArea {{
                background-color: {CARD};
                border: 1px solid {BORDER};
                border-radius: 8px;
            }}
        """)

        cont = QWidget()
        cont.setStyleSheet(f"background-color: {CARD};")
        cont_layout = QVBoxLayout(cont)
        cont_layout.setContentsMargins(8, 8, 8, 8)
        cont_layout.setSpacing(2)

        self._checkboxes = {}
        for numero, motivo in actas_error:
            color = {
                "DIF": "#B91C1C",
                "SAP": "#92400E",
                "SIN PDF": "#475569",
                "PDF": "#2563EB",
            }.get(motivo, "#B91C1C")

            chk = QCheckBox(f"{numero}   ·   {motivo}")
            chk.setStyleSheet(f"""
                QCheckBox {{
                    color: {color};
                    font-size: 10pt;
                    font-weight: 500;
                    padding: 8px 10px;
                    border-radius: 6px;
                    background: transparent;
                }}
                QCheckBox:hover {{ background-color: {CARD_2}; }}
                QCheckBox::indicator {{
                    width: 16px; height: 16px;
                    border-radius: 4px;
                    border: 1px solid {color};
                    background: {CARD};
                }}
                QCheckBox::indicator:checked {{
                    background-color: {color};
                    border-color: {color};
                }}
            """)
            chk.setCursor(Qt.PointingHandCursor)
            self._checkboxes[numero] = chk
            cont_layout.addWidget(chk)

        cont_layout.addStretch(1)
        scroll.setWidget(cont)
        v.addWidget(scroll, 1)

        acciones = QHBoxLayout()
        acciones.setSpacing(8)

        btn_todas = QPushButton("  Seleccionar todas")
        btn_todas.setCursor(Qt.PointingHandCursor)
        btn_todas.setStyleSheet(f"""
            QPushButton {{
                background-color: transparent; color: {TEXT_MUTED};
                border: none; padding: 6px 10px;
                font-size: 8.5pt; font-weight: 600; text-align: left;
            }}
            QPushButton:hover {{ color: {TEXT}; }}
        """)
        btn_todas.clicked.connect(lambda: [c.setChecked(True) for c in self._checkboxes.values()])
        acciones.addWidget(btn_todas)

        btn_ninguna = QPushButton("  Limpiar")
        btn_ninguna.setCursor(Qt.PointingHandCursor)
        btn_ninguna.setStyleSheet(f"""
            QPushButton {{
                background-color: transparent; color: {TEXT_MUTED};
                border: none; padding: 6px 10px;
                font-size: 8.5pt; font-weight: 600; text-align: left;
            }}
            QPushButton:hover {{ color: {TEXT}; }}
        """)
        btn_ninguna.clicked.connect(lambda: [c.setChecked(False) for c in self._checkboxes.values()])
        acciones.addWidget(btn_ninguna)

        acciones.addStretch(1)
        v.addLayout(acciones)

        botones = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel,
            Qt.Horizontal, self
        )

        btn_ok = botones.button(QDialogButtonBox.Ok)
        btn_ok.setText("  Ver seleccionadas")
        btn_ok.setStyleSheet(f"""
            QPushButton {{
                background-color: {ACCENT}; color: #FFFFFF;
                border: none; border-radius: 6px;
                padding: 9px 20px; font-weight: 600; font-size: 9.5pt;
            }}
            QPushButton:hover {{ background-color: {ACCENT_H}; }}
        """)

        btn_cancel = botones.button(QDialogButtonBox.Cancel)
        btn_cancel.setText("  Cancelar")
        btn_cancel.setStyleSheet(f"""
            QPushButton {{
                background-color: transparent; color: #374151;
                border: 1px solid {BORDER_2}; border-radius: 6px;
                padding: 9px 20px; font-weight: 500; font-size: 9.5pt;
            }}
            QPushButton:hover {{ background-color: #F3F4F6; }}
        """)

        botones.accepted.connect(self._aceptar)
        botones.rejected.connect(self.reject)
        v.addWidget(botones)

    def _aceptar(self):
        seleccion = [n for n, chk in self._checkboxes.items() if chk.isChecked()]
        self.numeros_seleccionados = seleccion if seleccion else None
        self.accept()


class DialogoHoja(QDialog):
    def __init__(self, hojas, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Seleccionar hoja")
        self.setMinimumWidth(420)
        self.hoja_elegida = None

        v = QVBoxLayout(self)
        v.setContentsMargins(24, 24, 24, 24)
        v.setSpacing(16)

        lbl = QLabel("El archivo tiene varias hojas. ¿Cuál querés cargar?")
        lbl.setStyleSheet(f"color: {TEXT}; font-size: 10.5pt; font-weight: 600;")
        lbl.setWordWrap(True)
        v.addWidget(lbl)

        self.lista = QListWidgetDialog()
        self.lista.setStyleSheet(f"""
            QListWidget {{
                border: 1px solid {BORDER_2}; border-radius: 8px;
                padding: 6px; background-color: {CARD}; font-size: 10pt;
            }}
            QListWidget::item {{ padding: 10px 12px; border-radius: 6px; }}
            QListWidget::item:selected {{
                background-color: {SEL_BG}; color: {TEXT};
            }}
            QListWidget::item:hover {{ background-color: #F3F4F6; }}
        """)
        for h in hojas:
            self.lista.addItem(h)
        self.lista.setCurrentRow(0)
        self.lista.itemDoubleClicked.connect(self._aceptar)
        v.addWidget(self.lista)

        botones = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel,
            Qt.Horizontal, self
        )
        btn_ok = botones.button(QDialogButtonBox.Ok)
        btn_ok.setText("  Cargar")
        btn_ok.setStyleSheet(f"""
            QPushButton {{
                background-color: {ACCENT}; color: #FFFFFF;
                border: none; border-radius: 6px;
                padding: 9px 20px; font-weight: 600; font-size: 9.5pt;
            }}
            QPushButton:hover {{ background-color: {ACCENT_H}; }}
        """)
        btn_cancel = botones.button(QDialogButtonBox.Cancel)
        btn_cancel.setText("  Cancelar")
        btn_cancel.setStyleSheet(f"""
            QPushButton {{
                background-color: transparent; color: #374151;
                border: 1px solid {BORDER_2}; border-radius: 6px;
                padding: 9px 20px; font-weight: 500; font-size: 9.5pt;
            }}
            QPushButton:hover {{ background-color: #F3F4F6; }}
        """)
        botones.accepted.connect(self._aceptar)
        botones.rejected.connect(self.reject)
        v.addWidget(botones)

    def _aceptar(self):
        item = self.lista.currentItem()
        if item:
            self.hoja_elegida = item.text()
            self.accept()


class OverlayCarga(QWidget):
    def __init__(self, parent):
        super().__init__(parent)
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setStyleSheet("background-color: rgba(15, 23, 42, 140);")
        self.hide()

        self._angulo = 0
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._rotar)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setAlignment(Qt.AlignCenter)

        self.card = QFrame()
        self.card.setStyleSheet(f"background-color: {CARD}; border-radius: 14px;")
        self.card.setFixedWidth(400)
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(50)
        shadow.setColor(QColor(0, 0, 0, 90))
        shadow.setOffset(0, 10)
        self.card.setGraphicsEffect(shadow)

        v = QVBoxLayout(self.card)
        v.setContentsMargins(36, 36, 36, 36)
        v.setSpacing(20)

        self.spinner = QLabel()
        self.spinner.setFixedSize(64, 64)
        self.spinner.setAlignment(Qt.AlignCenter)
        self.spinner.setStyleSheet("background: transparent;")
        v.addWidget(self.spinner, 0, Qt.AlignHCenter)

        self.lbl_titulo = QLabel("Procesando…")
        self.lbl_titulo.setAlignment(Qt.AlignCenter)
        self.lbl_titulo.setStyleSheet(
            f"color: {TEXT}; font-size: 14pt; font-weight: 700; background: transparent;"
        )
        v.addWidget(self.lbl_titulo)

        self.lbl_sub = QLabel("")
        self.lbl_sub.setAlignment(Qt.AlignCenter)
        self.lbl_sub.setWordWrap(True)
        self.lbl_sub.setStyleSheet(
            f"color: {TEXT_MUTED}; font-size: 10pt; background: transparent;"
        )
        v.addWidget(self.lbl_sub)

        self.barra = QProgressBar()
        self.barra.setTextVisible(False)
        self.barra.setFixedHeight(6)
        self.barra.setStyleSheet(f"""
            QProgressBar {{
                border: none; background-color: #E5E7EB; border-radius: 3px;
            }}
            QProgressBar::chunk {{
                background-color: {ACCENT}; border-radius: 3px;
            }}
        """)
        v.addWidget(self.barra)

        layout.addWidget(self.card)

    def _rotar(self):
        self._angulo = (self._angulo + 20) % 360
        pixmap = _make_icon("fa5s.spinner", ACCENT).pixmap(48, 48)
        rotated = pixmap.transformed(QTransform().rotate(self._angulo))
        self.spinner.setPixmap(rotated)

    def mostrar(self, titulo="Procesando…", subtitulo="", con_progreso=False):
        self.lbl_titulo.setText(titulo)
        self.lbl_sub.setText(subtitulo)
        self.barra.setVisible(con_progreso)
        if con_progreso:
            self.barra.setMaximum(100)
            self.barra.setValue(0)
        self.setGeometry(self.parent().rect())
        self.raise_()
        self.show()
        self._angulo = 0
        self._rotar()
        self._timer.start(60)
        QApplication.processEvents()

    def actualizar(self, titulo=None, subtitulo=None, valor=None, maximo=None):
        if titulo is not None:
            self.lbl_titulo.setText(titulo)
        if subtitulo is not None:
            self.lbl_sub.setText(subtitulo)
        if maximo is not None:
            self.barra.setMaximum(maximo)
        if valor is not None:
            self.barra.setValue(valor)
        QApplication.processEvents()

    def ocultar(self):
        self._timer.stop()
        self.hide()


class ComparadorWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Comparador de Actas · SAP vs PDF")
        self.resize(1700, 980)
        self.setMinimumSize(1250, 750)

        self.df = None
        self.indice_pdfs = None
        self.actas = []
        self.idx_actual = 0
        self.resultados_por_acta = {}
        self._widgets_lista = []
        self._filtro_numeros = None

        self.indice_fotos = {}

        self.modo_tabla = "acta"

        self.celda_acta = None
        self.zoom_acta_guardado = None
        self.ajustar_alto_acta_guardado = False

        self._build_ui()
        self._actualizar_botones()
        self._actualizar_lista_actas()

        self.overlay = OverlayCarga(self)

    def _care_came_de_acta(self, numero_acta):
        datos = self.indice_fotos.get(str(numero_acta))
        if not datos:
            return [], []
        return datos.get("care", []), datos.get("came", [])

    def _direccion_de_acta(self, numero_acta):
        r = self.resultados_por_acta.get(str(numero_acta))
        return _calle_altura_de_resultado(r)

    def _direccion_de_acta_actual(self):
        if not self.actas:
            return (None, None)
        numero = self.actas[self.idx_actual]
        return self._direccion_de_acta(numero)

    def _abrir_vista_foto(self, numero_acta, modo="grilla", ruta_foco=None):
        care, came = self._care_came_de_acta(numero_acta)
        dlg = VistaFotos(
            numero_acta, care, came,
            self._acta_anterior, self._acta_siguiente,
            modo_inicial=modo, ruta_foco=ruta_foco, parent=self,
        )
        dlg.set_obtener_care_came_fn(self._care_came_de_acta)
        dlg.set_obtener_direccion_fn(self._direccion_de_acta)
        dlg.exec_()

    def _abrir_vista_foto_en_indice(self, numero_acta, modo="individual", ruta_foco=None):
        care, came = self._care_came_de_acta(numero_acta)
        dlg = VistaFotos(
            numero_acta, care, came,
            self._acta_anterior, self._acta_siguiente,
            modo_inicial=modo, ruta_foco=ruta_foco, parent=self,
        )
        dlg.set_obtener_care_came_fn(self._care_came_de_acta)
        dlg.set_obtener_direccion_fn(self._direccion_de_acta)
        dlg.exec_()

    def _acta_anterior(self, numero_acta):
        try:
            nro = str(numero_acta)
            if nro in self.actas:
                idx = self.actas.index(nro)
                if idx > 0:
                    return self.actas[idx - 1]
        except Exception:
            pass
        return None

    def _acta_siguiente(self, numero_acta):
        try:
            nro = str(numero_acta)
            if nro in self.actas:
                idx = self.actas.index(nro)
                if idx < len(self.actas) - 1:
                    return self.actas[idx + 1]
        except Exception:
            pass
        return None

    def _guardar_zoom_acta(self):
        if self.celda_acta is not None and self.celda_acta.pixmaps:
            self.zoom_acta_guardado = self.celda_acta.zoom
            self.ajustar_alto_acta_guardado = self.celda_acta.ajustar_al_alto

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, "overlay") and self.overlay.isVisible():
            self.overlay.setGeometry(self.rect())

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)

        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        header = QFrame()
        header.setObjectName("Header")
        header.setFixedHeight(72)
        hl = QHBoxLayout(header)
        hl.setContentsMargins(24, 0, 24, 0)
        hl.setSpacing(12)

        logo = QLabel()
        logo.setPixmap(_make_icon("fa5s.clipboard-check", ACCENT).pixmap(26, 26))
        logo.setStyleSheet("background: transparent;")
        hl.addWidget(logo)

        titles = QVBoxLayout()
        titles.setSpacing(0)
        titles.setContentsMargins(0, 0, 0, 0)
        t1 = QLabel("Comparador de Actas")
        t1.setObjectName("AppTitle")
        t2 = QLabel("SAP · Planilla de medición vs. PDF del acta")
        t2.setObjectName("AppSubtitle")
        titles.addWidget(t1)
        titles.addWidget(t2)
        hl.addLayout(titles)

        hl.addStretch(1)

        self.btn_excel = QPushButton("  Abrir Excel")
        self.btn_excel.setObjectName("Secondary")
        self.btn_excel.setIcon(_make_icon("fa5s.file-excel", "#166534"))
        self.btn_excel.setIconSize(QSize(14, 14))
        self.btn_excel.setCursor(Qt.PointingHandCursor)
        self.btn_excel.clicked.connect(self.cargar_excel)
        hl.addWidget(self.btn_excel)

        self.btn_carpeta = QPushButton("  Abrir carpeta PDFs")
        self.btn_carpeta.setObjectName("Secondary")
        self.btn_carpeta.setIcon(_make_icon("fa5s.folder-open", "#374151"))
        self.btn_carpeta.setIconSize(QSize(14, 14))
        self.btn_carpeta.setCursor(Qt.PointingHandCursor)
        self.btn_carpeta.clicked.connect(self.cargar_carpeta)
        hl.addWidget(self.btn_carpeta)

        self.btn_fotos = QPushButton("  Abrir documentos")
        self.btn_fotos.setObjectName("Secondary")
        self.btn_fotos.setIcon(_make_icon("fa5s.folder", "#B45309"))
        self.btn_fotos.setIconSize(QSize(14, 14))
        self.btn_fotos.setCursor(Qt.PointingHandCursor)
        self.btn_fotos.clicked.connect(self.cargar_fotos)
        hl.addWidget(self.btn_fotos)

        self.btn_limpiar = QPushButton("  Limpiar")
        self.btn_limpiar.setObjectName("Danger")
        self.btn_limpiar.setIcon(_make_icon("fa5s.trash-alt", "#B91C1C"))
        self.btn_limpiar.setIconSize(QSize(14, 14))
        self.btn_limpiar.setCursor(Qt.PointingHandCursor)
        self.btn_limpiar.clicked.connect(self.limpiar_todo)
        hl.addWidget(self.btn_limpiar)

        root.addWidget(header)

        nav = QFrame()
        nav.setObjectName("NavBar")
        nav.setFixedHeight(52)
        nl = QHBoxLayout(nav)
        nl.setContentsMargins(24, 0, 24, 0)
        nl.setSpacing(12)

        self.lbl_nav = QLabel("Sin datos cargados")
        self.lbl_nav.setObjectName("NavTitle")
        nl.addWidget(self.lbl_nav)

        nl.addStretch(1)

        self.lbl_estado = QLabel("")
        self.lbl_estado.setObjectName("NavSub")
        nl.addWidget(self.lbl_estado)

        root.addWidget(nav)

        content_wrap = QWidget()
        cw = QVBoxLayout(content_wrap)
        cw.setContentsMargins(20, 20, 20, 16)
        cw.setSpacing(0)

        self.splitter_principal = QSplitter(Qt.Horizontal)
        self.splitter_principal.setHandleWidth(18)

        sidebar = QFrame()
        sidebar.setObjectName("Card")
        sidebar.setMinimumWidth(180)
        sv = QVBoxLayout(sidebar)
        sv.setContentsMargins(14, 16, 14, 14)
        sv.setSpacing(12)

        side_header = QHBoxLayout()
        side_header.setSpacing(8)
        side_header.setContentsMargins(4, 0, 4, 0)

        lbl_side = QLabel("ACTAS")
        lbl_side.setObjectName("SectionLabel")
        side_header.addWidget(lbl_side)

        self.btn_errores = QPushButton("")
        self.btn_errores.setCursor(Qt.PointingHandCursor)
        self.btn_errores.setIcon(_make_icon("fa5s.exclamation-triangle", "#FFFFFF"))
        self.btn_errores.setIconSize(QSize(10, 10))
        self.btn_errores.setFixedHeight(24)
        self.btn_errores.setStyleSheet(f"""
            QPushButton {{
                background-color: #DC2626; color: #FFFFFF;
                border: none; border-radius: 5px;
                padding: 2px 10px; font-weight: 700; font-size: 8.5pt;
            }}
            QPushButton:hover {{ background-color: #B91C1C; }}
        """)
        self.btn_errores.clicked.connect(self.abrir_dialogo_errores)
        self.btn_errores.setVisible(False)
        side_header.addWidget(self.btn_errores)

        side_header.addStretch(1)

        self.btn_prev = QPushButton()
        self.btn_prev.setObjectName("Nav")
        self.btn_prev.setIcon(_make_icon("fa5s.chevron-left", "#374151"))
        self.btn_prev.setIconSize(QSize(10, 10))
        self.btn_prev.setCursor(Qt.PointingHandCursor)
        self.btn_prev.clicked.connect(self.ir_anterior)
        side_header.addWidget(self.btn_prev)

        self.btn_next = QPushButton()
        self.btn_next.setObjectName("Nav")
        self.btn_next.setIcon(_make_icon("fa5s.chevron-right", "#374151"))
        self.btn_next.setIconSize(QSize(10, 10))
        self.btn_next.setCursor(Qt.PointingHandCursor)
        self.btn_next.clicked.connect(self.ir_siguiente)
        side_header.addWidget(self.btn_next)

        sv.addLayout(side_header)

        self.widget_filtro = QFrame()
        self.widget_filtro.setStyleSheet(f"""
            QFrame {{
                background-color: {ERR_BG};
                border: 1px solid #FCA5A5;
                border-radius: 6px;
            }}
        """)
        wf_layout = QHBoxLayout(self.widget_filtro)
        wf_layout.setContentsMargins(10, 6, 6, 6)
        wf_layout.setSpacing(6)

        self.lbl_filtro = QLabel("")
        self.lbl_filtro.setStyleSheet(
            f"color: #B91C1C; font-size: 8.5pt; font-weight: 700; background: transparent;"
        )
        wf_layout.addWidget(self.lbl_filtro)

        wf_layout.addStretch(1)

        self.btn_ver_todas = QPushButton("Ver todas")
        self.btn_ver_todas.setCursor(Qt.PointingHandCursor)
        self.btn_ver_todas.setStyleSheet(f"""
            QPushButton {{
                background-color: {CARD}; color: #B91C1C;
                border: 1px solid #FCA5A5; border-radius: 4px;
                padding: 3px 10px; font-weight: 700; font-size: 8pt;
            }}
            QPushButton:hover {{ background-color: #FEE2E2; }}
        """)
        self.btn_ver_todas.clicked.connect(self.ver_todas)
        wf_layout.addWidget(self.btn_ver_todas)

        self.widget_filtro.setVisible(False)
        sv.addWidget(self.widget_filtro)

        self.lista_actas = QListWidget()
        self.lista_actas.itemClicked.connect(self._click_lista)
        sv.addWidget(self.lista_actas, 1)

        self.splitter_principal.addWidget(sidebar)

        panel_excel = QFrame()
        panel_excel.setObjectName("Card")
        panel_excel.setMinimumWidth(180)
        vexc = QVBoxLayout(panel_excel)
        vexc.setContentsMargins(20, 18, 20, 20)
        vexc.setSpacing(12)

        fila_titulo = QHBoxLayout()
        fila_titulo.setSpacing(8)

        self.lbl_titulo_tabla = QLabel("PLANILLA SAP · FILAS DEL ACTA")
        self.lbl_titulo_tabla.setObjectName("SectionLabel")
        fila_titulo.addWidget(self.lbl_titulo_tabla)

        fila_titulo.addStretch(1)

        self.btn_modo_central = QPushButton("  Ver imagen del acta")
        self.btn_modo_central.setObjectName("Secondary")
        self.btn_modo_central.setIcon(_make_icon("fa5s.image", "#374151"))
        self.btn_modo_central.setIconSize(QSize(12, 12))
        self.btn_modo_central.setCursor(Qt.PointingHandCursor)
        self.btn_modo_central.clicked.connect(self._toggle_modo_central)
        fila_titulo.addWidget(self.btn_modo_central)

        self.btn_toggle_tabla = QPushButton("  Ver todas las filas")
        self.btn_toggle_tabla.setObjectName("Secondary")
        self.btn_toggle_tabla.setIcon(_make_icon("fa5s.list", "#374151"))
        self.btn_toggle_tabla.setIconSize(QSize(12, 12))
        self.btn_toggle_tabla.setCursor(Qt.PointingHandCursor)
        self.btn_toggle_tabla.clicked.connect(self._toggle_vista_tabla)
        fila_titulo.addWidget(self.btn_toggle_tabla)

        vexc.addLayout(fila_titulo)

        self.stack_central = QStackedWidget()

        self.tabla = TablaFiltrable()
        self.stack_central.addWidget(self.tabla)

        self.visor_acta = QWidget()
        vl = QVBoxLayout(self.visor_acta)
        vl.setContentsMargins(0, 0, 0, 0)
        vl.setSpacing(8)

        barra_visor = QHBoxLayout()
        barra_visor.setContentsMargins(0, 0, 0, 0)
        barra_visor.setSpacing(8)

        self.btn_acta_prev = QPushButton()
        self.btn_acta_prev.setObjectName("Nav")
        self.btn_acta_prev.setIcon(_make_icon("fa5s.chevron-left", "#374151"))
        self.btn_acta_prev.setIconSize(QSize(12, 12))
        self.btn_acta_prev.setCursor(Qt.PointingHandCursor)
        self.btn_acta_prev.setToolTip("Acta anterior")
        self.btn_acta_prev.clicked.connect(self.ir_anterior)
        barra_visor.addWidget(self.btn_acta_prev)

        self.lbl_acta_visor = QLabel("—")
        self.lbl_acta_visor.setStyleSheet(
            f"color: {TEXT}; font-size: 11pt; font-weight: 700;"
        )
        barra_visor.addWidget(self.lbl_acta_visor)

        self.btn_acta_next = QPushButton()
        self.btn_acta_next.setObjectName("Nav")
        self.btn_acta_next.setIcon(_make_icon("fa5s.chevron-right", "#374151"))
        self.btn_acta_next.setIconSize(QSize(12, 12))
        self.btn_acta_next.setCursor(Qt.PointingHandCursor)
        self.btn_acta_next.setToolTip("Acta siguiente")
        self.btn_acta_next.clicked.connect(self.ir_siguiente)
        barra_visor.addWidget(self.btn_acta_next)

        barra_visor.addStretch(1)

        self.btn_zoom_out_acta = QPushButton()
        self.btn_zoom_out_acta.setObjectName("Nav")
        self.btn_zoom_out_acta.setIcon(_make_icon("fa5s.search-minus", "#374151"))
        self.btn_zoom_out_acta.setIconSize(QSize(12, 12))
        self.btn_zoom_out_acta.setCursor(Qt.PointingHandCursor)
        self.btn_zoom_out_acta.setToolTip("Alejar (Ctrl + -)")
        self.btn_zoom_out_acta.clicked.connect(self._zoom_out_acta)
        barra_visor.addWidget(self.btn_zoom_out_acta)

        self.lbl_zoom_acta = ZoomSpin(
            celda_getter=lambda: self.celda_acta,
            aplicar_fn=self._actualizar_zoom_acta,
            ancho=60,
        )
        barra_visor.addWidget(self.lbl_zoom_acta)

        self.btn_zoom_in_acta = QPushButton()
        self.btn_zoom_in_acta.setObjectName("Nav")
        self.btn_zoom_in_acta.setIcon(_make_icon("fa5s.search-plus", "#374151"))
        self.btn_zoom_in_acta.setIconSize(QSize(12, 12))
        self.btn_zoom_in_acta.setCursor(Qt.PointingHandCursor)
        self.btn_zoom_in_acta.setToolTip("Acercar (Ctrl + +)")
        self.btn_zoom_in_acta.clicked.connect(self._zoom_in_acta)
        barra_visor.addWidget(self.btn_zoom_in_acta)

        self.btn_zoom_reset_acta = QPushButton("  Ajustar")
        self.btn_zoom_reset_acta.setObjectName("Secondary")
        self.btn_zoom_reset_acta.setIcon(_make_icon("fa5s.expand-arrows-alt", "#374151"))
        self.btn_zoom_reset_acta.setIconSize(QSize(12, 12))
        self.btn_zoom_reset_acta.setCursor(Qt.PointingHandCursor)
        self.btn_zoom_reset_acta.setToolTip("Ajustar a ventana")
        self.btn_zoom_reset_acta.clicked.connect(self._zoom_reset_acta)
        barra_visor.addWidget(self.btn_zoom_reset_acta)

        self.btn_abrir_pdf_sys = QPushButton("  Abrir PDF con sistema")
        self.btn_abrir_pdf_sys.setObjectName("Secondary")
        self.btn_abrir_pdf_sys.setIcon(_make_icon("fa5s.external-link-alt", "#374151"))
        self.btn_abrir_pdf_sys.setIconSize(QSize(12, 12))
        self.btn_abrir_pdf_sys.setCursor(Qt.PointingHandCursor)
        self.btn_abrir_pdf_sys.clicked.connect(self._abrir_pdf_sistema)
        barra_visor.addWidget(self.btn_abrir_pdf_sys)

        vl.addLayout(barra_visor)

        self.celda_acta = CeldaFoto(QPixmap())
        self.celda_acta.setStyleSheet(
            f"QScrollArea {{ background-color: {CARD_2}; "
            f"border: 1px solid {BORDER}; border-radius: 8px; }}"
        )
        vl.addWidget(self.celda_acta, 1)

        self.stack_central.addWidget(self.visor_acta)

        vexc.addWidget(self.stack_central, 1)

        self.splitter_principal.addWidget(panel_excel)

        panel_der = QFrame()
        panel_der.setObjectName("Card")
        panel_der.setMinimumWidth(200)
        vder = QVBoxLayout(panel_der)
        vder.setContentsMargins(0, 0, 0, 0)
        vder.setSpacing(0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setStyleSheet("QScrollArea { background: transparent; }")

        cont = QWidget()
        cont.setStyleSheet(f"background-color: {CARD};")
        self.vcont = QVBoxLayout(cont)
        self.vcont.setContentsMargins(20, 18, 20, 20)
        self.vcont.setSpacing(18)

        scroll.setWidget(cont)
        vder.addWidget(scroll)

        self.splitter_principal.addWidget(panel_der)

        self.splitter_principal.setSizes([300, 660, 740])
        self.splitter_principal.setStretchFactor(0, 0)
        self.splitter_principal.setStretchFactor(1, 4)
        self.splitter_principal.setStretchFactor(2, 5)

        self.splitter_principal.setCollapsible(0, False)
        self.splitter_principal.setCollapsible(1, False)
        self.splitter_principal.setCollapsible(2, False)

        cw.addWidget(self.splitter_principal, 1)
        root.addWidget(content_wrap, 1)

        self.status = self.statusBar()
        self.status.showMessage("Esperando planilla Excel o carpeta de PDFs…")

        self.setStyleSheet(QSS)

    def _toggle_modo_central(self):
        if self.modo_tabla == "imagen":
            self._guardar_zoom_acta()
            self.modo_tabla = "acta"
            self.btn_modo_central.setText("  Ver imagen del acta")
            self.btn_modo_central.setIcon(_make_icon("fa5s.image", "#374151"))
            self.lbl_titulo_tabla.setText("PLANILLA SAP · FILAS DEL ACTA")
            self.btn_toggle_tabla.setVisible(True)
            self.stack_central.setCurrentWidget(self.tabla)
            self._cargar_tabla_excel()
        else:
            self.modo_tabla = "imagen"
            self.btn_modo_central.setText("  Info del acta")
            self.btn_modo_central.setIcon(_make_icon("fa5s.info-circle", "#374151"))
            self.lbl_titulo_tabla.setText("ACTA COMPLETA · PDF")
            self.btn_toggle_tabla.setVisible(False)
            self.stack_central.setCurrentWidget(self.visor_acta)
            self._cargar_imagen_acta()

    def _cargar_imagen_acta(self):
        if not self.actas:
            self.lbl_acta_visor.setText("—")
            return

        numero = self.actas[self.idx_actual]
        self.lbl_acta_visor.setText(f"ACTA N° {numero}")

        self.btn_acta_prev.setEnabled(self.idx_actual > 0)
        self.btn_acta_next.setEnabled(self.idx_actual < len(self.actas) - 1)

        ruta_pdf = None
        num_pagina = 1

        if self.indice_pdfs is not None and numero in self.indice_pdfs["por_nro"]:
            id_norm = self.indice_pdfs["por_nro"][numero]
            entry = self.indice_pdfs["por_id"].get(id_norm)
            if entry:
                ruta_pdf = entry.get("ruta")
                datos = entry.get("datos") or {}
                num_pagina = datos.get("pagina", 1)

        layout = self.visor_acta.layout()
        if self.celda_acta is not None:
            layout.removeWidget(self.celda_acta)
            self.celda_acta.deleteLater()
            self.celda_acta = None

        self.celda_acta = CeldaFoto(QPixmap())
        self.celda_acta.setStyleSheet(
            f"QScrollArea {{ background-color: {CARD_2}; "
            f"border: 1px solid {BORDER}; border-radius: 8px; }}"
        )
        layout.addWidget(self.celda_acta, 1)

        if self.zoom_acta_guardado is not None:
            self.celda_acta.zoom = self.zoom_acta_guardado
            self.celda_acta.ajustar_al_alto = False
        else:
            self.celda_acta.zoom = None
            self.celda_acta.ajustar_al_alto = self.ajustar_alto_acta_guardado

        self.lbl_zoom_acta.actualizar_texto()

        if not ruta_pdf or not os.path.exists(ruta_pdf):
            self.celda_acta._agregar_mensaje(
                "⚠ No se encontró el PDF de esta acta.\n\n"
                "Cargá la carpeta de PDFs con 'Abrir carpeta PDFs'."
            )
            return

        pm = self._renderizar_pagina_pdf(ruta_pdf, num_pagina)
        if pm is None:
            if PYMUPDF_OK:
                self.celda_acta._agregar_mensaje(
                    f"⚠ No se pudo renderizar la página {num_pagina} del PDF."
                )
            else:
                self.celda_acta._agregar_mensaje(
                    "Para previsualizar PDFs, instalá pymupdf:\n\npip install pymupdf"
                )
            return

        self.celda_acta._agregar_pagina(pm)

        QTimer.singleShot(80, self._ajustar_zoom_acta)

    def _renderizar_pagina_pdf(self, ruta_pdf, num_pagina):
        if not PYMUPDF_OK:
            return None
        try:
            doc = fitz.open(ruta_pdf)
            idx = max(0, int(num_pagina) - 1)
            if idx >= doc.page_count:
                doc.close()
                return None
            pagina = doc.load_page(idx)
            mat = fitz.Matrix(1.5, 1.5)
            pix = pagina.get_pixmap(matrix=mat, alpha=False)
            img = QImage(
                pix.samples,
                pix.width,
                pix.height,
                pix.stride,
                QImage.Format_RGB888,
            )
            pm = QPixmap.fromImage(img.copy())
            doc.close()
            return pm
        except Exception:
            return None

    def _ajustar_zoom_acta(self):
        if self.celda_acta is None:
            return
        self.celda_acta._aplicar_zoom()
        self._actualizar_zoom_acta()

    def _actualizar_zoom_acta(self):
        if self.celda_acta is None:
            return
        try:
            self.lbl_zoom_acta.actualizar_texto()
        except Exception:
            pass

    def _actualizar_indicador_zoom(self, celda):
        if celda is self.celda_acta:
            try:
                self.lbl_zoom_acta.actualizar_texto()
            except Exception:
                pass

    def _zoom_in_acta(self):
        if self.celda_acta is not None:
            self.celda_acta.zoom_in()
            self._actualizar_zoom_acta()

    def _zoom_out_acta(self):
        if self.celda_acta is not None:
            self.celda_acta.zoom_out()
            self._actualizar_zoom_acta()

    def _zoom_reset_acta(self):
        if self.celda_acta is not None:
            self.celda_acta.zoom_reset()
            self._actualizar_zoom_acta()

    def _abrir_pdf_sistema(self):
        if not self.actas:
            return
        numero = self.actas[self.idx_actual]
        ruta_pdf = None
        if self.indice_pdfs is not None and numero in self.indice_pdfs["por_nro"]:
            id_norm = self.indice_pdfs["por_nro"][numero]
            entry = self.indice_pdfs["por_id"].get(id_norm)
            if entry:
                ruta_pdf = entry.get("ruta")
        if not ruta_pdf:
            QMessageBox.information(self, "Sin PDF", "No se encontró el PDF de esta acta.")
            return
        _abrir_con_sistema(ruta_pdf)

    def _toggle_vista_tabla(self):
        if self.modo_tabla == "imagen":
            return
        if self.modo_tabla == "acta":
            self.modo_tabla = "todo"
            self.btn_toggle_tabla.setText("  Ver solo esta acta")
            self.btn_toggle_tabla.setIcon(_make_icon("fa5s.file", "#374151"))
            self.lbl_titulo_tabla.setText("PLANILLA SAP · TODAS LAS FILAS")
        else:
            self.modo_tabla = "acta"
            self.btn_toggle_tabla.setText("  Ver todas las filas")
            self.btn_toggle_tabla.setIcon(_make_icon("fa5s.list", "#374151"))
            self.lbl_titulo_tabla.setText("PLANILLA SAP · FILAS DEL ACTA")
        self._cargar_tabla_excel()

    def cargar_fotos(self):
        carpeta = QFileDialog.getExistingDirectory(
            self, "Seleccionar carpeta con subcarpetas de documentos"
        )
        if not carpeta:
            return

        self.overlay.mostrar("Cargando documentos…", "Escaneando subcarpetas…")

        indice = {}
        try:
            subcarpetas = [
                d for d in os.listdir(carpeta)
                if os.path.isdir(os.path.join(carpeta, d))
            ]
            subcarpetas.sort()

            for i, sub in enumerate(subcarpetas):
                self.overlay.actualizar(
                    titulo="Cargando documentos…",
                    subtitulo=f"{i + 1}/{len(subcarpetas)} · {sub}",
                )

                if sub == "_cache":
                    continue

                ruta_sub = os.path.join(carpeta, sub)

                for nombre in sorted(os.listdir(ruta_sub)):
                    ruta_arch = os.path.join(ruta_sub, nombre)
                    if not os.path.isfile(ruta_arch):
                        continue
                    ext = os.path.splitext(nombre)[1].lower()
                    if ext in EXT_HTML:
                        _extraer_imagenes_de_html(ruta_arch)

                candidatos = []

                for nombre in sorted(os.listdir(ruta_sub)):
                    ruta_arch = os.path.join(ruta_sub, nombre)
                    if not os.path.isfile(ruta_arch):
                        continue
                    if _es_archivo_analizable(ruta_arch):
                        candidatos.append(ruta_arch)

                carpeta_cache = os.path.join(ruta_sub, "_cache")
                if os.path.isdir(carpeta_cache):
                    for nombre in sorted(os.listdir(carpeta_cache)):
                        ruta_arch = os.path.join(carpeta_cache, nombre)
                        if not os.path.isfile(ruta_arch):
                            continue
                        if _es_archivo_analizable(ruta_arch):
                            candidatos.append(ruta_arch)

                care = []
                came = []
                for r in candidatos:
                    tipo = _clasificar_archivo(r)
                    if tipo == "care":
                        care.append(r)
                    elif tipo == "came":
                        came.append(r)

                if not care and not came:
                    continue

                id_norm = _normalizar_id(sub)
                nro = _solo_numero_acta(sub)

                if not nro:
                    continue

                indice[nro] = {
                    "id": id_norm,
                    "care": care,
                    "came": came,
                }
        finally:
            self.overlay.ocultar()

        self.indice_fotos = indice

        if not indice:
            QMessageBox.information(
                self, "Sin archivos",
                "No se encontraron archivos CARE/CAME en las subcarpetas."
            )
            self.status.showMessage("No se encontraron documentos válidos.")
            return

        total_archivos = sum(
            len(v["care"]) + len(v["came"]) for v in indice.values()
        )
        self.status.showMessage(
            f"Documentos cargados · {len(indice)} actas · {total_archivos} archivos en total"
        )

        self._refrescar_lista_actas()
        self._recalcular_todos()
        self._actualizar_lista_actas()
        self._actualizar_boton_errores()
        self._mostrar_actual()

    def _acta_tiene_error(self, numero):
        r = self.resultados_por_acta.get(numero)
        if r is None:
            return (False, "")

        hay_pdfs_global = self.indice_pdfs is not None
        tipo = r.get("tipo", "desconocido")

        chk_act_ok = True
        if r.get("chk_actividad"):
            chk_act_ok = r["chk_actividad"].get("ok", True)
        chk_l_ok = True
        if r.get("chk_col_l"):
            chk_l_ok = r["chk_col_l"].get("ok", True)
        chk_aviso_ok = True
        if r.get("chk_clase_aviso"):
            chk_aviso_ok = r["chk_clase_aviso"].get("ok", True)

        excel_ok = chk_act_ok and chk_l_ok and chk_aviso_ok

        if not hay_pdfs_global:
            if tipo == "excel":
                if not excel_ok:
                    return (True, "SAP")
                return (False, "")
            return (False, "")

        if tipo == "excel":
            return (True, "SIN PDF")

        if tipo == "pdf":
            return (False, "")

        if tipo == "ambos":
            ok_items = r["filas_ok"] == r["filas_total"]
            ok_calle = r.get("ok_calle", False)
            ok_alturas = r.get("ok_alturas", False)

            if ok_items and ok_calle and ok_alturas and excel_ok:
                return (False, "")
            return (True, "DIF")

        return (False, "")

    def _calcular_actas_error(self):
        lista = []
        for numero in self.actas:
            tiene, motivo = self._acta_tiene_error(numero)
            if tiene:
                lista.append((numero, motivo))
        return lista

    def _actualizar_boton_errores(self):
        lista = self._calcular_actas_error()
        n = len(lista)
        if n > 0:
            self.btn_errores.setText(f"  Errores  {n}")
            self.btn_errores.setVisible(True)
        else:
            self.btn_errores.setVisible(False)

        if self._filtro_numeros:
            self.lbl_filtro.setText(f"Mostrando {len(self._filtro_numeros)} con error")
            self.widget_filtro.setVisible(True)
        else:
            self.widget_filtro.setVisible(False)

    def abrir_dialogo_errores(self):
        lista = self._calcular_actas_error()
        if not lista:
            QMessageBox.information(self, "Sin errores", "No hay actas con errores.")
            return

        dlg = DialogoErrores(lista, self)
        if dlg.exec_() == QDialog.Accepted:
            self._filtro_numeros = dlg.numeros_seleccionados
            self._actualizar_lista_actas()
            self._actualizar_boton_errores()
            if self._filtro_numeros:
                self.lbl_estado.setText(f"Filtrado: {len(self._filtro_numeros)} actas")
            else:
                self.lbl_estado.setText("")

    def ver_todas(self):
        self._filtro_numeros = None
        self._actualizar_lista_actas()
        self._actualizar_boton_errores()
        self.lbl_estado.setText("")

    def cargar_excel(self):
        ruta, _ = QFileDialog.getOpenFileName(
            self, "Seleccionar Excel", "", "Excel (*.xlsx *.xls)"
        )
        if not ruta:
            return

        hojas = obtener_nombres_hojas(ruta)
        hoja_elegida = None

        if len(hojas) > 1:
            dlg = DialogoHoja(hojas, self)
            if dlg.exec_() != QDialog.Accepted:
                return
            hoja_elegida = dlg.hoja_elegida

        self.overlay.mostrar("Cargando Excel…", "Leyendo la planilla…")
        try:
            self.df = leer_planilla(ruta, hoja=hoja_elegida)
            self.status.showMessage(
                f"Excel cargado · {len(obtener_actas_unicas(self.df))} actas detectadas"
            )
            self._refrescar_lista_actas()
            self._recalcular_todos()
            self._actualizar_lista_actas()
            self._actualizar_boton_errores()
            self._mostrar_actual()
        finally:
            self.overlay.ocultar()

    def cargar_carpeta(self):
        carpeta = QFileDialog.getExistingDirectory(self, "Seleccionar carpeta con PDFs")
        if not carpeta:
            return

        self.overlay.mostrar("Procesando PDFs…", "Iniciando…", con_progreso=True)

        def cb(actual, total, nombre):
            self.overlay.actualizar(
                titulo="Procesando PDFs…",
                subtitulo=f"{actual}/{total} · {nombre}",
                valor=actual,
                maximo=total,
            )

        try:
            self.indice_pdfs = escanear_carpeta_pdfs(carpeta, progreso=cb)
        finally:
            self.overlay.ocultar()

        n_pdfs = len(self.indice_pdfs["por_id"])
        n_sin = len(self.indice_pdfs["sin_id"])
        n_err = len(self.indice_pdfs["errores"])

        self.status.showMessage(
            f"Carpeta procesada · {n_pdfs} actas válidas · {n_sin} sin acta · {n_err} con error"
        )
        self._refrescar_lista_actas()
        self._recalcular_todos()
        self._actualizar_lista_actas()
        self._actualizar_boton_errores()
        self._mostrar_actual()

    def limpiar_todo(self):
        if self.df is None and self.indice_pdfs is None and not self.indice_fotos:
            return

        resp = QMessageBox.question(
            self,
            "Limpiar todo",
            "¿Seguro que querés limpiar todo lo cargado?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if resp != QMessageBox.Yes:
            return

        self.df = None
        self.indice_pdfs = None
        self.indice_fotos = {}
        self.actas = []
        self.idx_actual = 0
        self.resultados_por_acta = {}
        self.lista_actas.clear()
        self._widgets_lista = []
        self.tabla.cargar_datos([], [])
        self._limpiar_panel_derecho()
        self._filtro_numeros = None
        self.zoom_acta_guardado = None
        self.ajustar_alto_acta_guardado = False

        self.modo_tabla = "acta"
        self.btn_modo_central.setText("  Ver imagen del acta")
        self.btn_modo_central.setIcon(_make_icon("fa5s.image", "#374151"))
        self.btn_toggle_tabla.setVisible(True)
        self.btn_toggle_tabla.setText("  Ver todas las filas")
        self.btn_toggle_tabla.setIcon(_make_icon("fa5s.list", "#374151"))
        self.lbl_titulo_tabla.setText("PLANILLA SAP · FILAS DEL ACTA")
        self.stack_central.setCurrentWidget(self.tabla)

        layout = self.visor_acta.layout()
        if self.celda_acta is not None:
            layout.removeWidget(self.celda_acta)
            self.celda_acta.deleteLater()
            self.celda_acta = None
        self.celda_acta = CeldaFoto(QPixmap())
        self.celda_acta.setStyleSheet(
            f"QScrollArea {{ background-color: {CARD_2}; "
            f"border: 1px solid {BORDER}; border-radius: 8px; }}"
        )
        layout.addWidget(self.celda_acta, 1)
        self.lbl_acta_visor.setText("—")
        self.lbl_zoom_acta.actualizar_texto()

        self.lbl_nav.setText("Sin datos cargados")
        self.lbl_estado.setText("")
        self.btn_errores.setVisible(False)
        self.widget_filtro.setVisible(False)
        self._actualizar_botones()
        self.status.showMessage("Todo limpio · Esperando planilla Excel o carpeta de PDFs…")

    def _refrescar_lista_actas(self):
        numeros = set()

        if self.df is not None:
            for texto_breve in obtener_actas_unicas(self.df):
                nro = obtener_numero_acta(texto_breve)
                if nro:
                    numeros.add(nro)

        if self.indice_pdfs is not None:
            for id_norm, entry in self.indice_pdfs["por_id"].items():
                acta_id = entry["datos"].get("acta_id")
                if acta_id:
                    nro = _solo_numero_acta(acta_id)
                    if nro:
                        numeros.add(nro)

        for nro in self.indice_fotos.keys():
            numeros.add(nro)

        try:
            self.actas = sorted(numeros, key=lambda x: int(x))
        except ValueError:
            self.actas = sorted(numeros)

        self.idx_actual = min(self.idx_actual, max(0, len(self.actas) - 1))

    def _buscar_datos_excel_por_numero(self, numero):
        if self.df is None:
            return None, None
        for texto_breve in obtener_actas_unicas(self.df):
            if obtener_numero_acta(texto_breve) == numero:
                df_acta = obtener_datos_acta(self.df, texto_breve)
                if not df_acta.empty:
                    return df_acta, texto_breve
        return None, None

    def _buscar_datos_pdf_por_numero(self, numero):
        if self.indice_pdfs is None:
            return None
        if numero in self.indice_pdfs["por_nro"]:
            id_norm = self.indice_pdfs["por_nro"][numero]
            return self.indice_pdfs["por_id"].get(id_norm)
        return None

    def _recalcular_todos(self):
        self.resultados_por_acta = {}

        for numero in self.actas:
            df_acta, texto_breve = self._buscar_datos_excel_por_numero(numero)
            entry = self._buscar_datos_pdf_por_numero(numero)

            en_excel = df_acta is not None and not df_acta.empty
            en_pdf = entry is not None

            if en_excel and en_pdf:
                tipo = "ambos"
            elif en_excel:
                tipo = "excel"
            elif en_pdf:
                tipo = "pdf"
            else:
                tipo = "desconocido"

            resultado = {
                "acta_numero": numero,
                "acta_id_excel": texto_breve,
                "tipo": tipo,
                "encontrado": en_pdf,
                "df_acta": df_acta,
                "datos_pdf": entry["datos"] if entry else None,
                "ruta_pdf": entry["ruta"] if entry else None,
                "por_operacion": [],
                "filas_ok": 0,
                "filas_total": 0,
                "ok_calle": False,
                "ok_alturas": False,
                "chk_actividad": None,
                "chk_col_l": None,
                "chk_clase_aviso": None,
            }

            if en_excel and df_acta is not None and not df_acta.empty:
                fila0 = df_acta.iloc[0]
                resultado["chk_actividad"] = chequear_actividad(fila0)
                resultado["chk_col_l"] = chequear_direccion_col_l(fila0)
                resultado["chk_clase_aviso"] = chequear_clase_aviso(fila0)

            if tipo == "ambos":
                res = comparar(df_acta, entry["datos"])
                total_filas = sum(len(o["filas"]) for o in res["por_operacion"])
                ok_filas = sum(1 for o in res["por_operacion"] for f in o["filas"] if f["ok"])
                resultado["por_operacion"] = res["por_operacion"]
                resultado["filas_ok"] = ok_filas
                resultado["filas_total"] = total_filas

                fila0 = df_acta.iloc[0]
                texto_dir = str(fila0.get(COL_DENOMINACION, "")).strip()
                dir_res = comparar_direccion(texto_dir, entry["datos"])
                resultado["ok_calle"] = dir_res["ok_calle"]
                resultado["ok_alturas"] = dir_res["ok_alturas"]

            self.resultados_por_acta[numero] = resultado

    def ir_anterior(self):
        if self.modo_tabla == "imagen":
            self._guardar_zoom_acta()
        if self.idx_actual > 0:
            self.idx_actual -= 1
            self._mostrar_actual()

    def ir_siguiente(self):
        if self.modo_tabla == "imagen":
            self._guardar_zoom_acta()
        if self.idx_actual < len(self.actas) - 1:
            self.idx_actual += 1
            self._mostrar_actual()

    def _click_lista(self, item):
        idx = item.data(Qt.UserRole)
        if idx is not None and idx != self.idx_actual:
            if self.modo_tabla == "imagen":
                self._guardar_zoom_acta()
            self.idx_actual = idx
            self._mostrar_actual()

    def _actualizar_botones(self):
        tiene_actas = bool(self.actas)
        self.btn_prev.setEnabled(tiene_actas and self.idx_actual > 0)
        self.btn_next.setEnabled(tiene_actas and self.idx_actual < len(self.actas) - 1)
        self.btn_limpiar.setEnabled(
            self.df is not None or self.indice_pdfs is not None or bool(self.indice_fotos)
        )
        self.btn_acta_prev.setEnabled(tiene_actas and self.idx_actual > 0)
        self.btn_acta_next.setEnabled(tiene_actas and self.idx_actual < len(self.actas) - 1)

    def _acta_pasa_filtro(self, numero):
        if self._filtro_numeros is None:
            return True
        return numero in self._filtro_numeros

    def _actualizar_lista_actas(self):
        self.lista_actas.clear()
        self._widgets_lista = []
        if not self.actas:
            return

        for i, numero in enumerate(self.actas):
            if not self._acta_pasa_filtro(numero):
                continue

            r = self.resultados_por_acta.get(numero)
            tipo = r["tipo"] if r else "desconocido"

            tiene_archivos = numero in self.indice_fotos
            if tiene_archivos:
                cant_archivos = (
                    len(self.indice_fotos[numero].get("care", [])) +
                    len(self.indice_fotos[numero].get("came", []))
                )
            else:
                cant_archivos = 0

            if tipo == "ambos" and r and r["encontrado"]:
                ok_items = r["filas_ok"] == r["filas_total"]
                ok_calle = r.get("ok_calle", False)
                ok_alturas = r.get("ok_alturas", False)
                ok_act = r.get("chk_actividad") is None or r["chk_actividad"].get("ok", True)
                ok_col_l = r.get("chk_col_l") is None or r["chk_col_l"].get("ok", True)
                ok_aviso = r.get("chk_clase_aviso") is None or r["chk_clase_aviso"].get("ok", True)

                if ok_items and ok_calle and ok_alturas and ok_act and ok_col_l and ok_aviso:
                    dot_color = OK_DOT
                else:
                    dot_color = ERR_DOT

                detalle = f"{r['filas_ok']}/{r['filas_total']}"
                text_color = TEXT
            elif tipo == "excel":
                ok_act = r.get("chk_actividad") is None or r["chk_actividad"].get("ok", True)
                ok_col_l = r.get("chk_col_l") is None or r["chk_col_l"].get("ok", True)
                ok_aviso = r.get("chk_clase_aviso") is None or r["chk_clase_aviso"].get("ok", True)

                if ok_act and ok_col_l and ok_aviso:
                    dot_color = TEXT_DIM
                else:
                    dot_color = ERR_DOT

                text_color = TEXT
                detalle = "SAP"
            elif tipo == "pdf":
                dot_color = ACCENT
                text_color = TEXT
                detalle = "PDF"
            else:
                dot_color = NONE_DOT
                text_color = TEXT_DIM
                detalle = "—"

            item = QListWidgetItem()
            item.setData(Qt.UserRole, i)
            item.setSizeHint(QSize(0, 46))

            w = QWidget()
            h = QHBoxLayout(w)
            h.setContentsMargins(12, 6, 12, 6)
            h.setSpacing(8)

            dot = _dot(dot_color, 10)
            h.addWidget(dot)

            num = QLabel(numero)
            num.setStyleSheet(
                f"color: {text_color}; font-weight: 600; font-size: 11pt; "
                f"background: transparent;"
            )
            h.addWidget(num)

            h.addStretch(1)

            if tiene_archivos:
                cam_ico = QLabel()
                cam_ico.setPixmap(_make_icon("fa5s.camera", "#B45309").pixmap(13, 13))
                cam_ico.setToolTip(f"{cant_archivos} archivo(s)")
                cam_ico.setStyleSheet("background: transparent;")
                h.addWidget(cam_ico)

            cnt = QLabel(detalle)
            cnt.setStyleSheet(
                f"color: {TEXT_DIM}; font-size: 8.5pt; background: transparent;"
            )
            h.addWidget(cnt)

            self.lista_actas.addItem(item)
            self.lista_actas.setItemWidget(item, w)
            self._widgets_lista.append((i, w))

        self._seleccionar_en_lista(self.idx_actual)

    def _seleccionar_en_lista(self, idx_actual):
        for i, (real_idx, w) in enumerate(self._widgets_lista):
            if real_idx == idx_actual:
                self.lista_actas.blockSignals(True)
                self.lista_actas.setCurrentRow(i)
                self.lista_actas.blockSignals(False)
                for j, (_, ww) in enumerate(self._widgets_lista):
                    if j == i:
                        ww.setStyleSheet(
                            f"background-color: {SEL_BG}; border-radius: 6px;"
                        )
                    else:
                        ww.setStyleSheet(
                            "background-color: transparent; border-radius: 6px;"
                        )
                self.lista_actas.scrollToItem(
                    self.lista_actas.item(i),
                    QAbstractItemView.EnsureVisible
                )
                return

    def _limpiar_panel_derecho(self):
        while self.vcont.count():
            item = self.vcont.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()

    def _actualizar_label_estado(self, r, tipo, tipo_txt):
        if tipo == "ambos" and r and r["encontrado"]:
            ok_items = r["filas_ok"] == r["filas_total"]
            ok_calle = r.get("ok_calle", False)
            ok_alturas = r.get("ok_alturas", False)
            ok_act = r.get("chk_actividad") is None or r["chk_actividad"].get("ok", True)
            ok_col_l = r.get("chk_col_l") is None or r["chk_col_l"].get("ok", True)
            ok_aviso = r.get("chk_clase_aviso") is None or r["chk_clase_aviso"].get("ok", True)

            if ok_items and ok_calle and ok_alturas and ok_act and ok_col_l and ok_aviso:
                self.lbl_estado.setText(
                    f"OK  ·  Filas: {r['filas_ok']}/{r['filas_total']}  ·  Origen: {tipo_txt}"
                )
                self.lbl_estado.setStyleSheet(
                    f"color: {OK_FG}; font-weight: 700; font-size: 9.5pt;"
                )
            else:
                self.lbl_estado.setText(
                    f"CON DIFERENCIAS  ·  Filas: {r['filas_ok']}/{r['filas_total']}  ·  Origen: {tipo_txt}"
                )
                self.lbl_estado.setStyleSheet(
                    f"color: {ERR_FG}; font-weight: 700; font-size: 9.5pt;"
                )
        elif tipo == "pdf":
            self.lbl_estado.setStyleSheet(
                f"color: {ACCENT}; font-weight: 700; font-size: 9.5pt;"
            )
        else:
            self.lbl_estado.setStyleSheet(
                f"color: {TEXT_MUTED}; font-weight: 700; font-size: 9.5pt;"
            )

    def _poblar_panel_derecho(self, r, tipo, tipo_txt, numero):
        if tipo == "pdf":
            if r and r["datos_pdf"]:
                self._bloque_pdf(r["datos_pdf"])
            else:
                self._aviso("PDF NO ENCONTRADO", "No se pudo leer el PDF de esta acta.")
        elif tipo == "excel":
            if r and r["df_acta"] is not None and not r["df_acta"].empty:
                fila0 = r["df_acta"].iloc[0]
                self._bloque_sap(r["df_acta"], fila0, r)
            else:
                self._aviso("SIN DATOS", "No hay filas del SAP para esta acta.")
        elif tipo == "ambos":
            fila0 = r["df_acta"].iloc[0]
            self._bloque_sap(r["df_acta"], fila0, r)
            self._bloque_pdf(r["datos_pdf"])
            self._bloque_comparacion(r["df_acta"], fila0, r["datos_pdf"], r)
        else:
            self._aviso("SIN DATOS", "No se encontraron datos para esta acta.")

        self._bloque_fotos(numero)

    def _mostrar_actual(self):
        if not self.actas:
            self.lbl_nav.setText("Sin datos cargados")
            self.lbl_estado.setText("")
            self._limpiar_panel_derecho()
            self.tabla.cargar_datos([], [])
            return

        if not self._acta_pasa_filtro(self.actas[self.idx_actual]):
            for i, a in enumerate(self.actas):
                if self._acta_pasa_filtro(a):
                    self.idx_actual = i
                    break

        numero = self.actas[self.idx_actual]
        r = self.resultados_por_acta.get(numero)
        tipo = r["tipo"] if r else "desconocido"

        tipo_txt = {
            "excel": "SAP",
            "pdf": "PDF",
            "ambos": "SAP + PDF",
        }.get(tipo, tipo)

        self.lbl_nav.setText(
            f"Acta {self.idx_actual + 1} de {len(self.actas)}  ·  Acta N° {numero}"
        )
        self.lbl_estado.setText(f"Origen: {tipo_txt}")

        if self.modo_tabla == "imagen":
            self._cargar_imagen_acta()
            self._limpiar_panel_derecho()
            self._poblar_panel_derecho(r, tipo, tipo_txt, numero)
            self._actualizar_label_estado(r, tipo, tipo_txt)
            self.vcont.addStretch(1)
            self._actualizar_botones()
            self._seleccionar_en_lista(self.idx_actual)
            return

        self._cargar_tabla_excel()
        self._limpiar_panel_derecho()
        self._poblar_panel_derecho(r, tipo, tipo_txt, numero)
        self._actualizar_label_estado(r, tipo, tipo_txt)

        self.vcont.addStretch(1)
        self._actualizar_botones()
        self._seleccionar_en_lista(self.idx_actual)

    def _cargar_tabla_excel(self):
        if self.df is None:
            self.tabla.cargar_datos([], [])
            return

        if self.modo_tabla == "acta":
            if not self.actas:
                self.tabla.cargar_datos([], [])
                return
            numero = self.actas[self.idx_actual]
            r = self.resultados_por_acta.get(numero)
            if not r or r["df_acta"] is None or r["df_acta"].empty:
                self.tabla.cargar_datos([], [])
                return
            df = r["df_acta"]
        else:
            df = self.df

        columnas = [str(c) for c in df.columns]
        filas = []
        for _, fila in df.iterrows():
            fila_vals = []
            for c in df.columns:
                val = fila[c]
                if val is None:
                    fila_vals.append("")
                else:
                    if _es_fecha(val):
                        fila_vals.append(_formatear_fecha(val))
                    else:
                        fila_vals.append(str(val))
            filas.append(fila_vals)

        self.tabla.cargar_datos(columnas, filas)

    def _bloque_fotos(self, numero_acta):
        card = QFrame()
        card.setStyleSheet(f"background-color: {INFO_BG}; border-radius: 10px;")
        v = QVBoxLayout(card)
        v.setContentsMargins(20, 18, 20, 18)
        v.setSpacing(12)

        datos = self.indice_fotos.get(numero_acta)
        care = datos.get("care", []) if datos else []
        came = datos.get("came", []) if datos else []

        cant_total = len(care) + len(came)
        if cant_total:
            v.addWidget(_titulo_seccion(
                f"DOCUMENTOS DEL ACTA  ·  CAME: {len(came)}  ·  CARE: {len(care)}"
            ))
        else:
            v.addWidget(_titulo_seccion("DOCUMENTOS DEL ACTA"))

        barra = QHBoxLayout()
        barra.setContentsMargins(0, 0, 0, 0)
        barra.setSpacing(8)

        btn_ver_todas = QPushButton("  Ver todas juntas (CARE / CAME)")
        btn_ver_todas.setObjectName("Secondary")
        btn_ver_todas.setIcon(_make_icon("fa5s.th-large", "#374151"))
        btn_ver_todas.setIconSize(QSize(12, 12))
        btn_ver_todas.setCursor(Qt.PointingHandCursor)
        btn_ver_todas.clicked.connect(
            lambda _=None, n=numero_acta: self._abrir_vista_foto(n, modo="grilla")
        )
        barra.addWidget(btn_ver_todas)
        barra.addStretch(1)
        v.addLayout(barra)

        todas = list(came) + list(care)
        if todas:
            direccion_fn = (lambda n=numero_acta: self._direccion_de_acta(n))
            galeria = GaleriaFotos(todas, direccion_fn=direccion_fn)
            v.addWidget(galeria)
        else:
            lbl = QLabel("No se encontraron documentos para esta acta.")
            lbl.setStyleSheet(
                f"color: {TEXT_MUTED}; font-size: 9.5pt; font-style: italic;"
            )
            lbl.setWordWrap(True)
            v.addWidget(lbl)

        if datos:
            id_carpeta = datos.get("id", "")
            lbl_id = QLabel(f"Carpeta: {id_carpeta}")
            lbl_id.setStyleSheet(
                f"color: {TEXT_DIM}; font-size: 8pt; font-style: italic;"
            )
            lbl_id.setWordWrap(True)
            v.addWidget(lbl_id)

        self.vcont.addWidget(card)

    def _aviso(self, titulo, texto):
        card = QFrame()
        card.setStyleSheet(f"background-color: {INFO_BG}; border-radius: 10px;")
        v = QVBoxLayout(card)
        v.setContentsMargins(20, 18, 20, 18)
        v.setSpacing(8)
        v.addWidget(_titulo_seccion(titulo))
        lbl = QLabel(texto)
        lbl.setStyleSheet(f"color: {TEXT_MUTED}; font-size: 10pt;")
        lbl.setWordWrap(True)
        v.addWidget(lbl)
        self.vcont.addWidget(card)

    def _bloque_sap(self, df_acta, fila0, r=None):
        card = QFrame()
        card.setStyleSheet(f"background-color: {INFO_BG}; border-radius: 10px;")
        v = QVBoxLayout(card)
        v.setContentsMargins(20, 18, 20, 18)
        v.setSpacing(10)

        v.addWidget(_titulo_seccion("DATOS DEL SAP (EXCEL)"))
        v.addWidget(_fila_info("Texto breve", fila0.get(COL_TEXTO_BREVE)))
        v.addWidget(_fila_info("Orden superior (CARE)", obtener_care(fila0)))
        v.addWidget(_fila_info("Orden (CAME)", obtener_came(fila0)))
        v.addWidget(_fila_info("Denominación", fila0.get(COL_DENOMINACION)))

        nombre_col_m = obtener_nombre_col_m(self.df)
        nombre_col_l = obtener_nombre_col_l(self.df)

        if r and r.get("chk_actividad"):
            chk_act = r["chk_actividad"]
            v.addWidget(_fila_cmp(
                f"Cl.actividad PM vs {nombre_col_m}",
                chk_act["valor_pm"] or "—",
                chk_act["valor_m"] or "—",
                chk_act["ok"],
            ))

        if r and r.get("chk_col_l"):
            chk_l = r["chk_col_l"]
            v.addWidget(_fila_cmp(
                f"Denominación vs {nombre_col_l}",
                chk_l["denominacion"] or "—",
                chk_l["col_l"] or "—",
                chk_l["ok"],
            ))

        if r and r.get("chk_clase_aviso"):
            chk_av = r["chk_clase_aviso"]
            v.addWidget(_fila_check(
                "Clase de aviso",
                chk_av.get("valor_aviso") or "—",
                chk_av["ok"],
                chk_av.get("mensaje", ""),
            ))

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet(f"color: {BORDER};")
        v.addWidget(sep)

        lbl = QLabel("Operaciones / items:")
        lbl.setStyleSheet(
            f"color: {TEXT_MUTED}; font-size: 9pt; font-weight: 700; letter-spacing: 0.5px;"
        )
        v.addWidget(lbl)

        for _, fila in df_acta.iterrows():
            op = obtener_operacion(fila)
            val = obtener_valor(fila)
            v.addWidget(_fila_info(f"  {op}", _fmt(val)))

        self.vcont.addWidget(card)

    def _bloque_pdf(self, datos_pdf):
        card = QFrame()
        card.setStyleSheet(f"background-color: {INFO_BG}; border-radius: 10px;")
        v = QVBoxLayout(card)
        v.setContentsMargins(20, 18, 20, 18)
        v.setSpacing(10)

        v.addWidget(_titulo_seccion("DATOS LEÍDOS DEL PDF"))
        v.addWidget(_fila_info("N° de acta", datos_pdf.get("acta_id")))
        v.addWidget(_fila_info("Página", datos_pdf.get("pagina")))
        v.addWidget(_fila_info("CARE", datos_pdf.get("care")))
        v.addWidget(_fila_info("CAME", datos_pdf.get("came")))
        v.addWidget(_fila_info("Calle", datos_pdf.get("calle")))

        alturas = datos_pdf.get("alturas", [])
        if alturas:
            for i, a in enumerate(alturas, 1):
                v.addWidget(_fila_info(f"Alturas ({i})", f"{a['ini']} - {a['fin']}"))
        else:
            v.addWidget(_fila_info("Alturas", "—"))

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet(f"color: {BORDER};")
        v.addWidget(sep)

        lbl = QLabel("Operaciones / items:")
        lbl.setStyleSheet(
            f"color: {TEXT_MUTED}; font-size: 9pt; font-weight: 700; letter-spacing: 0.5px;"
        )
        v.addWidget(lbl)

        for it in datos_pdf.get("items", []):
            v.addWidget(_fila_info(f"  {it['item']}", _fmt(it["valor"])))

        self.vcont.addWidget(card)

    def _bloque_comparacion(self, df_acta, fila0, datos_pdf, r):
        card = QFrame()
        card.setStyleSheet(
            f"background-color: {CARD}; border: 1px solid {BORDER}; border-radius: 10px;"
        )
        v = QVBoxLayout(card)
        v.setContentsMargins(20, 18, 20, 18)
        v.setSpacing(10)

        v.addWidget(_titulo_seccion("COMPARACIÓN"))

        nro_sap = obtener_numero_acta(fila0.get(COL_TEXTO_BREVE))
        nro_pdf = ""
        if datos_pdf.get("acta_id"):
            nro_pdf = _solo_numero_acta(datos_pdf["acta_id"])
        ok_nro = nro_sap and nro_pdf and nro_sap == nro_pdf
        v.addWidget(_fila_cmp("N° de acta", nro_sap, nro_pdf, ok_nro))

        care_sap = obtener_care(fila0)
        care_pdf = datos_pdf.get("care") or ""
        ok_care = bool(care_sap and care_pdf and care_sap == care_pdf)
        v.addWidget(_fila_cmp("CARE", care_sap, care_pdf or "—", ok_care))

        came_sap = obtener_came(fila0)
        came_pdf = datos_pdf.get("came") or ""
        ok_came = bool(came_sap and came_pdf and came_sap == came_pdf)
        v.addWidget(_fila_cmp("CAME", came_sap, came_pdf or "—", ok_came))

        texto_dir = str(fila0.get(COL_DENOMINACION, "")).strip()
        dir_res = comparar_direccion(texto_dir, datos_pdf)
        ex = dir_res["excel"] or {}
        pdf_dir = dir_res["pdf"] or {}

        v.addWidget(_fila_cmp(
            "Calle",
            ex.get("calle") or "—",
            pdf_dir.get("calle") or "—",
            dir_res["ok_calle"],
        ))

        sap_alt = f"{ex.get('ini', '—')}-{ex.get('fin', '—')}"
        pdf_alts = pdf_dir.get("alturas", [])
        if pdf_alts:
            pdf_alt_txt = " | ".join(f"{a['ini']}-{a['fin']}" for a in pdf_alts)
        else:
            pdf_alt_txt = "—"
        v.addWidget(_fila_cmp("Alturas", sap_alt, pdf_alt_txt, dir_res["ok_alturas"]))

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet(f"color: {BORDER};")
        v.addWidget(sep)

        lbl = QLabel("Operaciones / items:")
        lbl.setStyleSheet(
            f"color: {TEXT_MUTED}; font-size: 9pt; font-weight: 700; letter-spacing: 0.5px;"
        )
        v.addWidget(lbl)

        for o in r["por_operacion"]:
            for f in o["filas"]:
                v.addWidget(_fila_cmp(
                    f"{o['operacion']}",
                    _fmt(f["excel"]),
                    _fmt(f["pdf"]),
                    f["ok"],
                ))

        self.vcont.addWidget(card)


QSS = f"""
* {{
    font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, 'Helvetica Neue', sans-serif;
}}

QMainWindow, QWidget {{
    background-color: {BG};
    color: {TEXT};
    font-size: 10pt;
}}

QFrame#Header {{
    background-color: {CARD};
    border-bottom: 1px solid {BORDER};
}}
QLabel#AppTitle {{
    font-size: 13pt;
    font-weight: 700;
    color: {TEXT};
    padding-left: 4px;
}}
QLabel#AppSubtitle {{
    font-size: 9pt;
    color: {TEXT_MUTED};
    padding-left: 4px;
}}

QFrame#NavBar {{
    background-color: {CARD};
    border-bottom: 1px solid {BORDER};
}}
QLabel#NavTitle {{
    font-size: 11pt;
    font-weight: 700;
    color: {TEXT};
}}
QLabel#NavSub {{
    font-size: 9.5pt;
    color: {TEXT_MUTED};
    font-weight: 600;
}}

QFrame#Card {{
    background-color: {CARD};
    border: 1px solid {BORDER};
    border-radius: 10px;
}}
QLabel#SectionLabel {{
    font-size: 9pt;
    font-weight: 700;
    color: {TEXT_MUTED};
    letter-spacing: 0.8px;
    padding: 2px 0;
}}

QPushButton {{
    background-color: transparent;
    color: #374151;
    border: 1px solid transparent;
    border-radius: 8px;
    padding: 9px 18px;
    font-weight: 500;
    font-size: 9.5pt;
}}
QPushButton:hover {{
    background-color: #F3F4F6;
}}
QPushButton:pressed {{
    background-color: #E5E7EB;
}}
QPushButton:disabled {{
    color: {TEXT_DIM};
    background-color: transparent;
    border-color: transparent;
}}

QPushButton#Secondary {{
    background-color: {CARD};
    color: #101828;
    border: 1px solid {BORDER_2};
    padding: 9px 18px;
    font-weight: 600;
}}
QPushButton#Secondary:hover {{
    background-color: #F9FAFB;
    border-color: #9CA3AF;
}}
QPushButton#Secondary:disabled {{
    background-color: #FAFAFA;
    color: {TEXT_DIM};
    border-color: {BORDER};
}}

QPushButton#Danger {{
    background-color: {CARD};
    color: #B91C1C;
    border: 1px solid {BORDER_2};
    padding: 9px 18px;
    font-weight: 600;
}}
QPushButton#Danger:hover {{
    background-color: {ERR_BG};
    border-color: #FCA5A5;
}}
QPushButton#Danger:disabled {{
    color: {TEXT_DIM};
    border-color: {BORDER};
    background-color: transparent;
}}

QPushButton#Nav {{
    background-color: {CARD};
    border: 1px solid {BORDER_2};
    border-radius: 6px;
    padding: 3px;
    min-width: 28px;
    max-width: 28px;
    min-height: 28px;
    max-height: 28px;
}}
QPushButton#Nav:hover {{
    background-color: #F9FAFB;
    border-color: #9CA3AF;
}}
QPushButton#Nav:disabled {{
    background-color: #FAFAFA;
    border-color: #F3F4F6;
}}

QPushButton#NavGrande {{
    background-color: {CARD};
    border: 1px solid {BORDER_2};
    border-radius: 8px;
    padding: 8px;
}}
QPushButton#NavGrande:hover {{
    background-color: {SEL_BG};
    border-color: {ACCENT};
}}
QPushButton#NavGrande:disabled {{
    background-color: #FAFAFA;
    border-color: #F3F4F6;
}}

QTableView {{
    background-color: {CARD};
    gridline-color: transparent;
    border: 1px solid {BORDER};
    border-radius: 8px;
    selection-background-color: {SEL_BG};
    selection-color: {TEXT};
    outline: none;
    font-size: 9.5pt;
    alternate-background-color: #F9FAFB;
}}
QTableView::item {{
    padding: 9px 12px;
    border-bottom: 1px solid #F3F4F6;
}}
QTableView::item:selected {{
    background-color: {SEL_BG};
    color: {TEXT};
}}

QListWidget {{
    background-color: transparent;
    border: none;
    outline: none;
    padding: 0;
}}
QListWidget::item {{
    border-radius: 6px;
    margin: 1px 0;
    color: {TEXT};
    background-color: transparent;
}}
QListWidget::item:hover {{
    background-color: #F9FAFB;
}}
QListWidget::item:selected {{
    background-color: {SEL_BG};
}}

QProgressBar {{
    border: none;
    background-color: #E5E7EB;
    border-radius: 2px;
}}
QProgressBar::chunk {{
    background-color: {ACCENT};
    border-radius: 2px;
}}

QSplitter::handle {{
    background-color: transparent;
}}
QSplitter::handle:horizontal {{
    width: 12px;
}}
QSplitter::handle:hover {{
    background-color: {ACCENT};
    border-radius: 6px;
}}

QScrollBar:vertical {{
    background: transparent;
    width: 10px;
    margin: 0;
    border: none;
}}
QScrollBar::handle:vertical {{
    background: #CBD5E1;
    border-radius: 5px;
    min-height: 30px;
}}
QScrollBar::handle:vertical:hover {{
    background: #94A3B8;
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical,
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{
    background: transparent;
    height: 0;
}}
QScrollBar:horizontal {{
    background: transparent;
    height: 10px;
    margin: 0;
    border: none;
}}
QScrollBar::handle:horizontal {{
    background: #CBD5E1;
    border-radius: 5px;
    min-width: 30px;
}}
QScrollBar::handle:horizontal:hover {{
    background: #94A3B8;
}}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal,
QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {{
    background: transparent;
    width: 0;
}}

QStatusBar {{
    background-color: {CARD};
    border-top: 1px solid {BORDER};
    color: {TEXT_MUTED};
    font-size: 9pt;
    padding: 6px 16px;
}}
QStatusBar::item {{ border: none; }}

QToolTip {{
    background-color: #1F2937;
    color: #FFFFFF;
    border: none;
    padding: 6px 10px;
    border-radius: 6px;
    font-size: 9pt;
}}
"""


def main():
    QApplication.setStyle("Fusion")
    app = QApplication(sys.argv)
    app.setFont(QFont("Segoe UI", 10))

    win = ComparadorWindow()
    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()