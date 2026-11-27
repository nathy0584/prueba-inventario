import base64
import os
import openpyxl
import requests

ONEDRIVE_SHARE_URL = os.environ["ONEDRIVE_SHARE_URL"].strip()
SUPABASE_URL = os.environ["SUPABASE_URL"].rstrip("/")
SUPABASE_KEY = os.environ["SUPABASE_SERVICE_ROLE_KEY"]
PERIODO = os.environ["PERIODO"]
RUTA_EXCEL = "/tmp/consolidado_onedrive.xlsx"

HEADERS = {
    "apikey": SUPABASE_KEY,
    "Authorization": f"Bearer {SUPABASE_KEY}",
    "Content-Type": "application/json",
    "Prefer": "resolution=merge-duplicates",
}

def url_descarga_onedrive(share_url: str) -> str:
    b64 = base64.urlsafe_b64encode(share_url.encode("utf-8")).decode("utf-8").rstrip("=")
    return f"https://api.onedrive.com/v1.0/shares/u!{b64}/root/content"

def descargar_excel():
    urls_a_probar = [url_descarga_onedrive(ONEDRIVE_SHARE_URL), ONEDRIVE_SHARE_URL]
    ultimo_error = None

    for url in urls_a_probar:
        try:
            resp = requests.get(url, allow_redirects=True, timeout=60)
            resp.raise_for_status()
            contenido = resp.content
            if contenido[:2] == b"PK":
                with open(RUTA_EXCEL, "wb") as f:
                    f.write(contenido)
                return
        except Exception as e:
            ultimo_error = e

    raise RuntimeError(
        f"No se pudo descargar un .xlsx válido desde OneDrive. Último error: {ultimo_error}"
    )

def parse_sugerido(valor):
    if valor is None:
        return 0.0, False
    texto = str(valor).strip()
    if not texto or texto.lower() == "stock suficiente":
        return 0.0, True
    try:
        return float(texto.replace(",", ".")), False
    except ValueError:
        return 0.0, False

def main():
    descargar_excel()
    wb = openpyxl.load_workbook(RUTA_EXCEL, data_only=True)
    utm_sheets = [s for s in wb.sheetnames if s.upper().startswith("UTM")]
    filas = []

    for nombre_hoja in utm_sheets:
        ws = wb[nombre_hoja]
        headers = {
            str(ws.cell(1, col).value).strip().upper(): col
            for col in range(1, ws.max_column + 1)
            if ws.cell(1, col).value is not None
        }
        col_cod = headers.get("CODIGO")
        col_sug = headers.get("SUGERIDO") or headers.get("A SOLICITAR")
        if not col_cod or not col_sug:
            continue

        for fila_idx in range(2, ws.max_row + 1):
            codigo = ws.cell(fila_idx, col_cod).value
            if codigo is None or str(codigo).strip() == "":
                continue
            cant, stock_suf = parse_sugerido(ws.cell(fila_idx, col_sug).value)
            filas.append({
                "periodo": PERIODO,
                "origen_lista": nombre_hoja.strip(),
                "codigo": str(codigo).strip(),
                "cantidad_sugerida": cant,
                "stock_suficiente": stock_suf,
            })

    if not filas:
        raise RuntimeError("No se encontraron filas UTM válidas para sincronizar.")

    url = f"{SUPABASE_URL}/rest/v1/sugeridos_mensuales?on_conflict=periodo,origen_lista,codigo"
    resp = requests.post(url, headers=HEADERS, json=filas, timeout=30)
    resp.raise_for_status()
    print(f"Sincronizadas {len(filas)} filas desde OneDrive para el periodo {PERIODO}.")

if __name__ == "__main__":
    main()
