import time
import psycopg2
import pandas as pd

DB_HOST = "db.maximaapparel.com"
DB_PORT = 5432
DB_NAME = "maxima_reporting"
DB_USER = "agent_ro"
DB_PASSWORD = "R3@d1234!1"

TABLE = "public.sss_upc_report"

# Cuántas veces reintenta conectar antes de rendirse, y cuánto espera entre
# intentos. Esto cubre el caso en que la laptop acaba de despertar de suspensión
# y la red (WiFi/VPN) todavía no está lista.
CONNECT_RETRIES = 4
RETRY_WAIT_SECONDS = 3


def _connect_with_retry(progress_fn=None):
    """Intenta conectar varias veces. Devuelve la conexión o lanza el último error."""
    last_error = None
    for attempt in range(1, CONNECT_RETRIES + 1):
        try:
            return psycopg2.connect(
                host=DB_HOST, port=DB_PORT, dbname=DB_NAME,
                user=DB_USER, password=DB_PASSWORD,
                connect_timeout=15
            )
        except Exception as e:
            last_error = e
            if attempt < CONNECT_RETRIES:
                if progress_fn:
                    progress_fn(
                        f"Reintentando conexión ({attempt}/{CONNECT_RETRIES})... "
                        f"la red puede estar despertando."
                    )
                time.sleep(RETRY_WAIT_SECONDS)
    raise last_error


def run_download(progress_fn=None):
    if progress_fn:
        progress_fn("Conectando a la base de datos...")
    try:
        conn = _connect_with_retry(progress_fn)
    except Exception as e:
        if progress_fn:
            progress_fn(f"ERROR de conexión: {e}")
        return []

    try:
        if progress_fn:
            progress_fn(f"Jalando datos de {TABLE}...")

        query = f"""
            SELECT
                ivnum AS style,
                ivstyle AS base_style,
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

        return df

    except Exception as e:
        if progress_fn:
            progress_fn(f"ERROR al consultar: {e}")
        try:
            conn.close()
        except Exception:
            pass
        return None
