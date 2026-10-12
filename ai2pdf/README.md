# AI → PDF — Pro Standard

App Streamlit independiente. Navega el servidor de diseño y descarga los archivos de Illustrator (`.ai`) ya convertidos a PDF.

## Cómo correrlo localmente

```bash
cd ai2pdf
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

Se abre en `http://localhost:8501`. Debe correr en una máquina con acceso a `\\10.0.1.30`.

## Cómo funciona

1. Ruta base: `\\10.0.1.30\Design\DESIGN\Pro Standard` (editable en el sidebar).
2. Se navega por niveles: **Año → Quarter → Cápsula → Liga → Equipo**
   Ej: `2027\Q3 2027\Q3 ELITE TEAM\MLB\HOUSTON ASTROS`
3. Lista los `.ai` de la carpeta (y sus subcarpetas, opcional).
4. **Monedas y formato**:
   - **Monedas** (multiselect): CAD, MXN, EUR, GBP, AED… Sin monedas = originales USD.
   - **Una moneda por archivo** → una carpeta por moneda (`CAD/`, `MXN/`).
   - **Todas en un archivo** → carpeta `MULTI USD-CAD-MXN/`; cada precio muestra todas las monedas, en líneas o en una sola línea.
   - **Formato**: PDF y/o JPG (100–300 DPI).
5. **Generar** → tabla de revisión de precios (original → nuevo, con avisos), preview y descarga (`.zip`).

## Cambio de moneda (`currency.py`)

- Solo aplica a precios que son **texto editable**. Se quita únicamente el texto del precio (sin tapar nada) y se escribe el nuevo en la misma posición, tamaño y color.
- Si hay texto después del precio en la misma línea (ej. `$45.00 each`), se recorre para que no se encime.
- Usa la fuente original del archivo si tiene todas las letras necesarias; si no, una parecida (y lo avisa).
- Si el precio nuevo no cabe, se reduce el tamaño (mínimo 60%) y lo avisa.
- En modo "en líneas", si no hay espacio libre abajo, ese precio se pone en una sola línea (y lo avisa).
- Los PDFs con monedas cambiadas no llevan los datos privados de Illustrator; si no, al abrirlos en Illustrator mostrarían los precios USD originales.
- **Tipos de cambio provisionales** (editables en la app). Se reemplazan cuando se conecte la fuente oficial de precios.

## Notas técnicas

- Un `.ai` guardado con **"Create PDF Compatible File"** (default de Illustrator) ya es un PDF por dentro. La app lo valida con `pypdf` y lo entrega como `.pdf`, sin necesidad de Illustrator.
- Después de convertir, revisa cada PDF y dice si los precios son **texto editable** (y cuáles encontró) o si están en **contornos** (no editables).
- Se marcan como error (para exportar a mano desde Illustrator):
  - `.ai` guardados **sin** "PDF Compatible File" (solo traen una página placeholder).
  - `.ai` antiguos en formato PostScript.
- Por default se ignoran las carpetas **OLD** (`OLD`, `_OLD`, `OLD FILES`, `OLDS`…) para tomar solo lo más actual. Se puede desactivar en el sidebar.
- Las carpetas con `_` al inicio (ej. `_DALLAS COWBOYS`) se muestran sin el guion bajo.
- Paso previo a la opción **International** del Deck Builder (precios en otras monedas).
