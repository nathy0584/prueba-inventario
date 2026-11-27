import os
import openpyxl
import requests

SUPABASE_URL = os.environ["SUPABASE_URL"].rstrip("/")
SUPABASE_KEY = os.environ["SUPABASE_SERVICE_ROLE_KEY"].strip()
PERIODO = os.environ["PERIODO"].strip()
RUTA_EXCEL = "Planilla de Consolidado UTM Solicitud insumos mensual.xlsx"

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


def descargar_excel():
  if not os.path.exists(RUTA_EXCEL):
    raise FileNotFoundError(
        f"No se encontró el archivo local en el repositorio: {RUTA_EXCEL}"
    )
  if not _es_xlsx_valido(RUTA_EXCEL):
    raise RuntimeError(
        f"El archivo {RUTA_EXCEL} no parece ser un .xlsx válido."
    )


def _normalizar_booleano(valor) -> bool:
  if isinstance(valor, bool):
    return valor
  if valor is None:
    return False
  texto = str(valor).strip().upper()
  return texto in ("SI", "SÍ", "TRUE", "1", "OK", "SUFICIENTE", "YES")


def _parsear_a_solicitar(valor) -> tuple[int, bool]:
  if valor is None:
    return 0, False
  if isinstance(valor, bool):
    return 0, valor
  if isinstance(valor, (int, float)):
    num = int(round(float(valor)))
    return (num if num > 0 else 0), False

  texto = str(valor).strip()
  if not texto:
    return 0, False

  texto_upper = texto.upper()
  if "STOCK SUFICIENTE" in texto_upper or texto_upper in ("OK", "SUFICIENTE"):
    return 0, True

  try:
    num = int(round(float(texto.replace(",", "."))))
    return (num if num > 0 else 0), False
  except ValueError:
    return 0, _normalizar_booleano(texto)


def leer_sugeridos_desde_excel(ruta_archivo: str) -> list[dict]:
  wb = openpyxl.load_workbook(ruta_archivo, data_only=True)
  registros = []

  for nombre_hoja in wb.sheetnames:
    ws = wb[nombre_hoja]
    filas = list(ws.iter_rows(values_only=True))
    if len(filas) < 4:
      continue

    fila_enc_idx = None
    encabezados = []
    col_codigo = None
    col_solicitar = None

    for idx in range(min(10, len(filas))):
      candidatos_enc = [
          str(c).strip().lower() if c is not None else "" for c in filas[idx]
      ]

      def buscar_idx(nombres: list[str]) -> int | None:
        for n in nombres:
          if n in candidatos_enc:
            return candidatos_enc.index(n)
        return None

      c_prod = buscar_idx(
          ["producto", "insumo", "codigo", "código", "id_insumo", "id", "cod"]
      )
      c_sol = buscar_idx([
          "a solicitar",
          "a_solicitar",
          "solicitar",
          "cantidad_sugerida",
          "sugerido",
          "cantidad",
      ])
      if c_prod is not None and c_sol is not None:
        fila_enc_idx = idx
        encabezados = candidatos_enc
        col_codigo = c_prod
        col_solicitar = c_sol
        break

    if fila_enc_idx is None or col_codigo is None or col_solicitar is None:
      continue

    def buscar_opcional(nombres: list[str]) -> int | None:
      for n in nombres:
        if n in encabezados:
          return encabezados.index(n)
      return None

    col_stock = buscar_opcional(
        ["stock_suficiente", "stock suficiente", "alcanza", "estado_stock"]
    )
    col_origen = buscar_opcional(
        ["origen_lista", "origen", "lista", "utm", "sucursal"]
    )

    for fila in filas[fila_enc_idx + 1 :]:
      prod_raw = fila[col_codigo] if col_codigo < len(fila) else None
      if prod_raw is None or str(prod_raw).strip() == "":
        continue

      producto_str = str(prod_raw).strip()
      if producto_str.upper() in ("PRODUCTO", "TOTAL", "SUBTOTAL"):
        continue

      sol_raw = fila[col_solicitar] if col_solicitar < len(fila) else None
      cantidad_sugerida, stock_suf_por_texto = _parsear_a_solicitar(sol_raw)

      stock_raw = (
          fila[col_stock] if (col_stock is not None and col_stock < len(fila)) else None
      )
      stock_suficiente = (
          True if stock_suf_por_texto else _normalizar_booleano(stock_raw)
      )

      if cantidad_sugerida <= 0 and not stock_suficiente:
        continue

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
          "codigo": producto_str,
          "cantidad_sugerida": cantidad_sugerida,
          "stock_suficiente": stock_suficiente,
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
