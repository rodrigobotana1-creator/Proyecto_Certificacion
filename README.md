\# Comparador de Actas · SAP vs PDF



Herramienta de escritorio (Python + PyQt5) para \*\*verificar automáticamente actas de medición\*\* comparando los datos de una planilla SAP (Excel) contra el PDF/HTML oficial del acta, y para \*\*visualizar todos los documentos asociados\*\* (fotos, PDFs) de cada acta de forma ordenada.



\---



\## Características



\### Comparación SAP vs PDF

\- Carga una \*\*planilla de medición\*\* en Excel (SAP).

\- Carga una \*\*carpeta de PDFs\*\* con las actas emitidas.

\- Compara campo a campo por acta: N° de acta, CARE, CAME, calle, alturas, clase de actividad PM, clase de aviso, operaciones / ítems.

\- Marca cada campo con ✓ / ✗ y muestra un resumen por acta (OK / con diferencias).

\- Filtra y lista las \*\*actas con error\*\* para revisarlas rápido.

\- Exporta un \*\*reporte Excel\*\* con todos los resultados.



\### Gestión visual de documentos por acta

\- Escanea una carpeta con subcarpetas por acta.

\- \*\*Extrae automáticamente las imágenes embebidas en base64\*\* de los archivos `.html` y las guarda en una subcarpeta `\_cache/`.

\- Clasifica los archivos según su nombre:

&#x20; - `\*CAME\*` → columna \*\*CAME\*\*.

&#x20; - `\*CARE\*` → columna \*\*CARE\*\*.

\- Ignora `.html`, `.txt` y la carpeta `\_cache` como archivos a mostrar.



\### Visualización

\- \*\*Vista partida CAME | CARE\*\*: dos columnas con scroll vertical propio.

\- \*\*Vista individual\*\*: una imagen/PDF a la vez con navegación por flechas y teclado.

\- \*\*Zoom\*\* independiente por columna (Ctrl + rueda, `+`/`-`, `Ctrl+0`).

\- Rotación de imágenes.

\- Galería de miniaturas con checkboxes.

\- Abrir cualquier archivo con el programa del sistema.



\### Tabla SAP filtrable

\- Vista tipo Excel con \*\*filtros en cascada por columna\*\*.

\- Alterna entre ver solo las filas del acta actual o \*\*todas las filas\*\*.



\---



\## Requisitos



\- \*\*Python 3.8 o superior\*\*

\- Windows / Linux / macOS



\### Dependencias



| Paquete | Uso |

|---|---|

| `PyQt5` | Interfaz gráfica |

| `qtawesome` | Iconos FontAwesome |

| `pymupdf` | Renderizado de PDFs |

| `pandas` | Lectura de Excel |

| `openpyxl` | Soporte de `.xlsx` |



\---



\## Instalación



\### 1. Clonar el repositorio



```bash

git clone https://github.com/TU-USUARIO/TU-REPO.git

cd TU-REPO

```



\### 2. Crear un entorno virtual (recomendado)



\*\*Windows:\*\*

```bash

python -m venv venv

venv\\Scripts\\activate

```



\*\*Linux / macOS:\*\*

```bash

python3 -m venv venv

source venv/bin/activate

```



\### 3. Instalar las dependencias



Opción A — desde `requirements.txt`:

```bash

pip install -r requirements.txt

```



Opción B — instalar una por una:

```bash

pip install PyQt5

pip install qtawesome

pip install pymupdf

pip install pandas

pip install openpyxl

```



\### 4. Ejecutar la aplicación



```bash

python main.py

```



\---



\## Estructura del proyecto



```

.

├── main.py                 # Interfaz principal

├── requirements.txt        # Dependencias

├── README.md

├── .gitignore

└── modules/

&#x20;   ├── excel\_reader.py     # Lectura y parseo de la planilla SAP

&#x20;   ├── pdf\_parser.py       # Extracción de datos del PDF del acta

&#x20;   ├── comparator.py       # Lógica de comparación SAP vs PDF

&#x20;   ├── batch\_loader.py     # Escaneo de carpetas y matching

&#x20;   └── report.py           # Generación del reporte Excel

```



\---



\## Flujo de uso



1\. \*\*Abrir Excel\*\* → cargar la planilla SAP (si tiene varias hojas, elegir una).

2\. \*\*Abrir carpeta PDFs\*\* → cargar las actas en PDF para comparar.

3\. \*\*Abrir documentos\*\* → cargar la carpeta con las subcarpetas por acta (fotos + PDFs).

4\. Navegar por la lista de actas a la izquierda.

5\. En el panel derecho ver los datos SAP, los datos del PDF, la comparación y los documentos.

6\. \*\*Ver todas juntas (CARE / CAME)\*\* → pantalla partida con todos los documentos apilados.

7\. \*\*Exportar reporte\*\* → Excel con el resumen de todas las actas.



\---



\## Estructura de carpetas esperada



Para la carga de documentos, la carpeta debe tener esta forma:



```

Carpeta raíz/

├── P6\_2\_Z13\_R21\_CERT\_35\_ACTA\_02421/

│   ├── data\_CAME.PDF

│   ├── image\_CARE.html

│   ├── image\_CARE (1).html

│   ├── datos.txt

│   └── \_cache/

│       ├── image\_CARE.jpeg

│       ├── image\_CARE (1)\_1.jpeg

│       ├── image\_CARE (1)\_2.jpeg

│       └── image\_CARE (1)\_3.jpeg

└── P6\_2\_Z13\_R21\_CERT\_35\_ACTA\_02422/

&#x20;   └── ...

```



\- Cada \*\*subcarpeta\*\* representa un \*\*acta\*\*.

\- El número de acta se extrae del nombre de la subcarpeta.

\- Los archivos se clasifican por nombre: si contiene `care` → CARE, si contiene `came` → CAME.



\---



\## Notas



\- La carpeta `\_cache/` se crea automáticamente junto a cada HTML y se puede borrar: se regenera la próxima vez que se abre la carpeta.

\- Los PDFs se renderizan con PyMuPDF; si no está instalado, la app sigue funcionando sin previsualización de PDFs.

\- Todo el procesamiento es \*\*local\*\*: no se sube nada a internet.



\---



\## Licencia



Este proyecto se distribuye bajo la licencia MIT. Ver el archivo `LICENSE` para más detalles.


