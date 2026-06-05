import psycopg2
import pandas as pd
import os
import tempfile

DB_HOST = "db.maximaapparel.com"
DB_PORT = 5432
DB_NAME = "maxima_reporting"
DB_USER = "agent_ro"
DB_PASSWORD = "R3@d1234!1"

TABLE = "public.sss_upc_report"


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
        if progress_fn:
            progress_fn(f"Jalando datos de {TABLE}...")

        query = f"""
            SELECT
                ivstyle AS style,
                ivnum,
                size,
                upc,
                ivdesc AS description,
                brand_name,
                reporting_brand_name,
                wholesale_usd, msrp_usd,
                wholesale_cad, msrp_cad,
                wholesale_gbp, msrp_gbp,
                wholesale_eur, msrp_eur,
                wholesale_aed, msrp_aed,
                wholesale_mxn, msrp_mxn,
                wholesale_brl, msrp_brl,
                wholesale_clp, msrp_clp,
                wholesale_aud, msrp_aud,
                wholesale_nzd, msrp_nzd,
                wholesale_rmb, msrp_rmb,
                wholesale_ars, msrp_ars,
                wholesale_ecu, msrp_ecu,
                wholesale_bob, msrp_bob,
                wholesale_pen, msrp_pen
            FROM {TABLE}
        """

        df = pd.read_sql(query, conn)
        conn.close()

        if progress_fn:
            progress_fn(f"Datos obtenidos: {len(df):,} registros.")

        download_dir = tempfile.mkdtemp(prefix="upcs_")
        file_path = os.path.join(download_dir, "upcs_data.xlsx")
        df.to_excel(file_path, index=False)
        return [file_path]

    except Exception as e:
        if progress_fn:
            progress_fn(f"ERROR al consultar: {e}")
        try:
            conn.close()
        except Exception:
            pass
        return []
