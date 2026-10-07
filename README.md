<div align="center">

# 📊 Comparador de Actas · SAP vs PDF

**Herramienta de escritorio (Python + PyQt5) para verificar automáticamente actas de medición, comparando los datos de una planilla SAP (Excel) contra el PDF/HTML oficial del acta, y visualizar todos los documentos asociados (fotos, PDFs) de cada acta de forma ordenada.**

![Python](https://img.shields.io/badge/Python-3.8%2B-blue?style=for-the-badge&logo=python&logoColor=white)
![PyQt5](https://img.shields.io/badge/PyQt5-GUI-green?style=for-the-badge&logo=qt&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-yellow?style=for-the-badge)

</div>

---

## ✨ Características

### 🔍 Comparación SAP vs PDF
- Carga una **planilla de medición** en Excel (SAP).
- Carga una **carpeta de PDFs** con las actas emitidas.
- Compara campo a campo por acta: N° de acta, CARE, CAME, calle, alturas, clase de actividad PM, clase de aviso, operaciones / ítems.
- Marca cada campo con ✅ / ❌ y muestra un resumen por acta (OK / con diferencias).
- Filtra y lista las **actas con error** para revisarlas rápido.
- Exporta un **reporte Excel** con todos los resultados.

### 🖼️ Gestión visual de documentos por acta
- Escanea una carpeta con subcarpetas por acta.
- **Extrae automáticamente las imágenes embebidas en base64** de los archivos `.html` y las guarda en una subcarpeta `_cache/`.
- Clasifica los archivos según su nombre:
  - `*CAME*` → columna **CAME**.
  - `*CARE*` → columna **CARE**.
- Ignora `.html`, `.txt` y la carpeta `_cache` como archivos a mostrar.

### 🖥️ Visualización
- **Vista partida CAME | CARE**: dos columnas con scroll vertical propio.
- **Vista individual**: una imagen/PDF a la vez con navegación por flechas y teclado.
- **Zoom** independiente por columna (Ctrl + rueda, `+`/`-`, `Ctrl+0`).
- Rotación de imágenes.
- Galería de miniaturas con checkboxes.
- Abrir cualquier archivo con el programa del sistema.

### 📋 Tabla SAP filtrable
- Vista tipo Excel con **filtros en cascada por columna**.
- Alterna entre ver solo las filas del acta actual o **todas las filas**.

---

## ⚙️ Requisitos

- **Python 3.8 o superior**
- Windows / Linux / macOS

### 📦 Dependencias

| Paquete | Uso |
| :--- | :--- |
| `PyQt5` | Interfaz gráfica |
| `qtawesome` | Iconos FontAwesome |
| `pymupdf` | Renderizado de PDFs |
| `pandas` | Lectura de Excel |
| `openpyxl` | Soporte de `.xlsx` |

---

## 🚀 Instalación

### 1. Clonar el repositorio
```bash
git clone https://github.com/rodrigobotana1-creator/Proyecto_Certificacion.git
cd Proyecto_Certificacion