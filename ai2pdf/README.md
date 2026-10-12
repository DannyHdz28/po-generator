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
4. **Convertir a PDF** → descarga un PDF (1 archivo) o un `.zip` con todos.

## Notas técnicas

- Un `.ai` guardado con **"Create PDF Compatible File"** (default de Illustrator) ya es un PDF por dentro. La app lo valida con `pypdf` y lo entrega como `.pdf`, sin necesidad de Illustrator.
- Después de convertir, revisa cada PDF y dice si los precios son **texto editable** (y cuáles encontró) o si están en **contornos** (no editables).
- Se marcan como error (para exportar a mano desde Illustrator):
  - `.ai` guardados **sin** "PDF Compatible File" (solo traen una página placeholder).
  - `.ai` antiguos en formato PostScript.
- Las carpetas con `_` al inicio (ej. `_DALLAS COWBOYS`) se muestran sin el guion bajo.
- Paso previo a la opción **International** del Deck Builder (precios en otras monedas).
