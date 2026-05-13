# Deck Builder — Pro Standard

App Streamlit independiente. Reemplaza automáticamente las imágenes de un PPT base con los merchboards de un equipo, basándose en las notas de cada slide.

## Cómo correrlo localmente

```bash
cd deck
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

Se abre en `http://localhost:8501`.

## Cómo funciona

1. **PPT base:** las slides de merchboard deben tener en sus **Notas** el nombre de la cápsula + género. Ej: `FLAGSHIP M`, `SPRING BREAK W`, `TEAM CITY K`.
2. **Merchboards:** los archivos se nombran `LIGA_EQUIPO_CAPSULA_GENERO[_NNN].ext`.
   - 1 imagen: `NFL_LAS VEGAS RAIDERS_FLAGSHIP_M.jpg`
   - Varias:   `NFL_LAS VEGAS RAIDERS_FLAGSHIP_M_001.jpg`, `..._002.jpg`
3. La app matchea por cápsula + género y reemplaza los bytes de la imagen embebida (conserva tamaño, posición y demás formato del slide).

## Notas técnicas

- Manipula el PPT con `python-pptx` (no toca XML a mano, no usa JSZip).
- El reemplazo se hace sobre el `ImagePart` embebido, así que respeta el `<p:pic>` original.
- Si una slide marcada no tiene imagen existente, se loguea como advertencia (la inserción de imagen nueva todavía no está implementada en server-side — se puede agregar después si la necesitas).
- Cápsulas soportadas por defecto: cualquier texto en notas. El matching es flexible: tolera `&` ↔ espacio, mayúsculas/minúsculas y número de slide al final.
