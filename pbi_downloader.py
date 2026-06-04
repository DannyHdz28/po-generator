import psycopg2
import pandas as pd
import os
import tempfile

DB_HOST = "db.maximaapparel.com"
DB_PORT = 5432
DB_NAME = "postgres"
DB_USER = "agent_ro"
DB_PASSWORD = "R3@d1234!1"

SIZES = ["2T", "3T", "4", "4T", "5", "6", "7", "OS", "XS", "S", "M", "L", "XL", "2XL", "3XL"]

TABLES_TO_TRY = ["plm.style_info", "mat_view.plm_style_info"]


def get_columns(conn, table):
    schema, tname = table.split(".")
    cur = conn.cursor()
    cur.execute("""
        SELECT column_name FROM information_schema.columns
        WHERE table_schema = %s AND table_name = %s
        ORDER BY ordinal_position
    """, (schema, tname))
    cols = [r[0] for r in cur.fetchall()]
    cur.close()
    return cols


def find_col(cols, candidates):
    for c in candidates:
        if c in cols:
            return c
    return None


def run_download(progress_fn=None):
    if progress_fn:
        progress_fn("Conectando a la base de datos...")
    try:
        conn = psycopg2.connect(
            host=DB_HOST, port=DB_PORT, dbname=DB_NAME,
            user=DB_USER, password=DB_PASSWORD,
            connect_timeout=15
        )
    except Exception as e:
        if progress_fn:
            progress_fn(f"ERROR de conexión: {e}")
        return []

    try:
        for table in TABLES_TO_TRY:
            if progress_fn:
                progress_fn(f"Revisando tabla {table}...")

            cols = get_columns(conn, table)
            if not cols:
                if progress_fn:
                    progress_fn(f"Tabla {table} no existe o sin permisos.")
                continue

            if progress_fn:
                progress_fn(f"Columnas en {table}: {', '.join(cols)}")

            # Buscar columnas necesarias
            style_col = find_col(cols, ["ivnum", "style", "style_number", "style_no", "sku", "ivstyle"])
            size_col   = find_col(cols, ["size", "talla", "size_code"])
            upc_col    = find_col(cols, ["upc", "barcode", "upc_code"])
            desc_col   = find_col(cols, ["description", "descripcion", "desc", "item_description"])
            ws_col     = find_col(cols, ["wholesale", "cost", "unit_cost", "ws", "wholesale_price"])
            msrp_col   = find_col(cols, ["msrp", "retail", "retail_price", "srp"])

            if not style_col or not upc_col:
                if progress_fn:
                    progress_fn(f"No se encontraron columnas de Style/UPC en {table}. Intentando siguiente tabla...")
                continue

            select_parts = [f"{style_col} AS style"]
            if size_col:   select_parts.append(f"{size_col} AS size")
            if upc_col:    select_parts.append(f"{upc_col} AS upc")
            if desc_col:   select_parts.append(f"{desc_col} AS description")
            if ws_col:     select_parts.append(f"{ws_col} AS wholesale")
            if msrp_col:   select_parts.append(f"{msrp_col} AS msrp")

            size_filter = ""
            if size_col:
                placeholders = ",".join(["%s"] * len(SIZES))
                size_filter = f"WHERE {size_col} IN ({placeholders})"

            query = f"SELECT {', '.join(select_parts)} FROM {table} {size_filter}"

            if progress_fn:
                progress_fn(f"Jalando datos de {table}...")

            df = pd.read_sql(query, conn, params=SIZES if size_col else None)
            conn.close()

            if progress_fn:
                progress_fn(f"Datos obtenidos: {len(df):,} registros de {table}.")

            download_dir = tempfile.mkdtemp(prefix="upcs_")
            file_path = os.path.join(download_dir, "upcs_data.xlsx")
            df.to_excel(file_path, index=False)
            return [file_path]

        conn.close()
        if progress_fn:
            progress_fn("ERROR: No se encontraron tablas con columnas de Style y UPC.")
        return []

    except Exception as e:
        if progress_fn:
            progress_fn(f"ERROR al consultar: {e}")
        try:
            conn.close()
        except Exception:
            pass
        return []
