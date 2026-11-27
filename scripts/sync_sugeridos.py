import base64
import os
import openpyxl
import requests

ONEDRIVE_SHARE_URL = os.environ["ONEDRIVE_SHARE_URL"].strip()
SUPABASE_URL = os.environ["SUPABASE_URL"].rstrip("/")
SUPABASE_KEY = os.environ["SUPABASE_SERVICE_ROLE_KEY"].strip()
PERIODO = os.environ["PERIODO"].strip()
RUTA_EXCEL = "/tmp/consolidado_onedrive.xlsx"

HEADERS_SUPABASE = {
    "apikey": SUPABASE_KEY,
    "Authorization": f"Bearer {SUPABASE_KEY}",
    "Content-Type": "application/json",
    "Prefer": "resolution=merge-duplicates",
}


def _es_xlsx_valido(ruta_archivo: str) -> bool:
  try:
    with open(ruta_archivo, "rb") as f:
      firma = f.read(4)
    return firma == b"PK\x03\x04"
  except OSError:
    return False


def _intentar_descarga_directa(url: str, destino: str) -> tuple[bool, str]:
  sep = "&" if "?" in url else "?"
  url_descarga = url if "download=1" in url else f"{url}{sep}download=1"
  headers = {"User-Agent": "Mozilla/5.0"}
  try:
    resp = requests.get(
        url_descarga, headers=headers, allow_redirects=True, timeout=60
    )
    resp.raise_for_status()
    with open(destino, "wb") as f:
      f.write(resp.content)
    if _es_xlsx_valido(destino):
      return True, ""
    return False, (
        "La descarga directa no devolvió un archivo .xlsx válido (posible"
        " HTML de vista previa o login)."
    )
  except Exception as exc:
    return False, str(exc)


def descargar_excel():
  ok_directa, err_directa = _intentar_descarga_directa(
      ONEDRIVE_SHARE_URL, RUTA_EXCEL
  )
  if ok_directa:
    return

  encoded = (
      base64.urlsafe_b64encode(ONEDRIVE_SHARE_URL.encode("utf-8"))
      .decode("utf-8")
      .rstrip("=")
  )
  share_api_url = (
      f"https://api.onedrive.com/v1.0/shares/u!{encoded}/root/content"
  )
  try:
    resp = requests.get(share_api_url, allow_redirects=True, timeout=60)
    resp.raise_for_status()
    with open(RUTA_EXCEL, "wb") as f:
      f.write(resp.content)
    if _es_xlsx_valido(RUTA_EXCEL):
      return
    err_api = "El endpoint /shares no devolvió un .xlsx válido."
  except Exception as exc:
    err_api = str(exc)

  raise RuntimeError(
      "No se pudo descargar un .xlsx válido desde OneDrive. "
      f"Directa: {err_directa} | API shares: {err_api}"
  )


def _normalizar_booleano(valor) -> bool:
  if isinstance(valor, bool):
    return valor
  if valor is None:
    return False
  texto = str(valor).strip().upper()
  return texto in ("SI", "SÍ", "TRUE", "1", "OK", "SUFICIENTE", "YES")


def leer_sugeridos_desde_excel(ruta_archivo: str) -> list[dict]:
  wb = openpyxl.load_workbook(ruta_archivo, data_only=True)
  registros = []

  for nombre_hoja in wb.sheetnames:
    ws = wb[nombre_hoja]
    filas = list(ws.iter_rows(values_only=True))
    if not filas:
      continue

    encabezados = [
        str(c).strip().lower() if c is not None else "" for c in filas[0]
    ]

    def idx_col(candidatos: list[str]) -> int | None:
      for cand in candidatos:
        if cand in encabezados:
          return encabezados.index(cand)
      return None

    col_codigo = idx_col(["codigo", "código", "id_insumo", "id", "cod"])
    col_cant = idx_col(
        ["cantidad_sugerida", "sugerido", "cantidad", "cant_sugerida"]
    )
    col_stock = idx_col(
        ["stock_suficiente", "stock suficiente", "alcanza", "estado_stock"]
    )
    col_origen = idx_col(["origen_lista", "origen", "lista", "utm", "sucursal"])

    if col_codigo is None or col_cant is None:
      continue

    for fila in filas[1:]:
      codigo_raw = fila[col_codigo] if col_codigo < len(fila) else None
      cant_raw = fila[col_cant] if col_cant < len(fila) else None
      if codigo_raw is None or str(codigo_raw).strip() == "":
        continue

      try:
        cantidad_sugerida = float(cant_raw)
      except (TypeError, ValueError):
        cantidad_sugerida = 0.0

      if cantidad_sugerida <= 0:
        continue

      stock_raw = (
          fila[col_stock] if (col_stock is not None and col_stock < len(fila)) else None
      )
      origen_raw = (
          fila[col_origen]
          if (col_origen is not None and col_origen < len(fila))
          else None
      )
      origen_lista = (
          str(origen_raw).strip()
          if origen_raw is not None and str(origen_raw).strip() != ""
          else nombre_hoja.strip()
      )

      registros.append({
          "periodo": PERIODO,
          "origen_lista": origen_lista,
          "codigo": str(codigo_raw).strip(),
          "cantidad_sugerida": cantidad_sugerida,
          "stock_suficiente": _normalizar_booleano(stock_raw),
      })

  return registros


def sincronizar_a_supabase(registros: list[dict]):
  if not registros:
    print(f"No se encontraron filas válidas para el periodo {PERIODO}.")
    return

  endpoint = f"{SUPABASE_URL}/rest/v1/sugeridos_mensuales"
  tamano_lote = 200

  for i in range(0, len(registros), tamano_lote):
    lote = registros[i : i + tamano_lote]
    resp = requests.post(
        endpoint, headers=HEADERS_SUPABASE, json=lote, timeout=60
    )
    if not resp.ok:
      raise RuntimeError(
          f"Error al hacer upsert en Supabase ({resp.status_code}): {resp.text}"
      )

  print(
      f"Sincronización completada: {len(registros)} registros enviados para el"
      f" periodo {PERIODO}."
  )


def main():
  descargar_excel()
  registros = leer_sugeridos_desde_excel(RUTA_EXCEL)
  sincronizar_a_supabase(registros)


if __name__ == "__main__":
  main()
