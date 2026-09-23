"""Lee el detalle del PDT 621 (IGV-Renta Mensual) YA GUARDADO en Supabase
Storage (declaraciones_pdt.url_detalle) y saca las casillas reales que
declaro el contribuyente -- para que la Evolucion Compras-Ventas se llene
con lo REALMENTE declarado, no con la propuesta SIRE (que puede traer
comprobantes que el contador nunca tomo).

Soporta 2 formatos de 'detalle', detectados por la extension de la URL:
- .pdf (el caso normal): texto extraido con pdfplumber, formato regular
  'etiqueta casilla valor [casilla2 valor2]', confirmado en vivo el
  2026-09-23 con un PDF real de DIABETES (202608).
- .csv (fallback, 2026-09-23): a veces SUNAT no genera un detalle
  descargable para una declaracion ya presentada ("al consultar no nos da
  el detalle" -- confirmado con DIABETES 202503, que tiene constancia pero
  url_detalle=NULL en declaraciones_pdt). Para esos casos el despacho
  obtiene un export alterno de SUNAT con las mismas casillas en CSV
  (columnas 'Nro Casilla'/'Valor Casilla', una fila por casilla) -- se
  sube ese CSV a Storage y se usa como url_detalle en su lugar.

Las casillas del formulario 0621 son ESTANDAR de SUNAT (mismo numero en
todo el pais), asi que estos numeros de casilla no cambian entre empresas
ni periodos.

Regla de "cual declaracion es la valida" (pedido explicito del usuario):
una empresa puede tener Original + Sustitutoria + Rectificatoria para el
mismo periodo -- SUNAT asigna num_orden en orden creciente, asi que la de
mayor num_orden es siempre la mas reciente/definitiva, sea cual sea su
'Tipo de Declaracion' (ese texto tambien se extrae, solo para mostrarlo)."""
import csv
import io
import re

import httpx
import pdfplumber

from .supabase_client import sb

CASILLAS = {
    "ventas_gravadas": "100",       # Ventas Netas (BASE gravada)
    "ventas_no_gravadas": "105",    # Ventas no Gravadas (sin exportaciones)
    "ventas_no_gravadas_sin_ratio": "109",
    "igv_ventas": "131",            # Total IGV Ventas
    "compras_gravadas": "107",      # Compras netas destinada a vtas gravadas (nacional)
    "compras_no_gravadas": "120",   # Compras internas no gravadas
    "igv_compras": "178",           # TOTAL credito fiscal
    "ingresos_netos": "301",        # Base de Renta = ventas_gravadas + ventas_no_gravadas
    "renta_resultante": "302",      # Impuesto Resultante o Saldo a Favor (Renta)
    "coeficiente_aplicado": "315",  # Coeficiente o Porcentaje efectivamente aplicado ese mes
}


def _num(texto, casilla):
    # Busca "<casilla> <numero>" -- el numero siempre viene INMEDIATAMENTE
    # despues de su propia casilla (aunque la fila tenga una 2da casilla
    # despues, esa queda mas lejos y no matchea el \s+ inmediato).
    m = re.search(rf"\b{casilla}\s+(-?[\d,]+\.\d\d)", texto)
    if not m:
        return 0.0
    return float(m.group(1).replace(",", ""))


def _tipo_declaracion(texto):
    m = re.search(r"Tipo de Declaraci.n\s+(\w+)", texto)
    return m.group(1) if m else None


def _armar_resultado(valores, tipo_declaracion):
    """valores: dict {casilla_str: float}, ya sea sacado del PDF o del CSV."""
    def v(clave):
        return valores.get(CASILLAS[clave], 0.0)

    ventas_gravadas = v("ventas_gravadas")
    ventas_no_gravadas = v("ventas_no_gravadas")
    compras_gravadas = v("compras_gravadas")
    compras_no_gravadas = v("compras_no_gravadas")
    return {
        "tipo_declaracion": tipo_declaracion,
        "ventas_gravadas": ventas_gravadas,
        "ventas_no_gravadas": ventas_no_gravadas,
        "ventas_total": v("ingresos_netos") or (ventas_gravadas + ventas_no_gravadas),
        "igv_ventas": v("igv_ventas"),
        "compras_gravadas": compras_gravadas,
        "compras_no_gravadas": compras_no_gravadas,
        # OJO: no existe una sola casilla oficial de "compras total" en el
        # PDT 621 -- se suma la base gravada + no gravada (nacional). Si la
        # empresa tiene compras IMPORTADAS (casillas 114/119/122) quedarian
        # afuera; para las ~65 empresas del despacho esto no se ha visto
        # todavia, pero queda documentado por si aparece.
        "compras_total": compras_gravadas + compras_no_gravadas,
        "igv_compras": v("igv_compras"),
        "renta_resultante": v("renta_resultante"),
        # Tasa (coeficiente o 1%) que SUNAT registra como aplicada ese mes
        # -- informativo; NO se usa todavia para decidir la tasa del mes
        # que se esta calculando (ver igv_renta.regla_300_uit), porque un
        # valor de 1.0 aqui puede significar "se uso el piso de 1%", no
        # necesariamente el coeficiente real de la empresa.
        "coeficiente_aplicado": v("coeficiente_aplicado"),
    }


def parsear_pdt_pdf(pdf_bytes):
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        texto = "\n".join(p.extract_text() or "" for p in pdf.pages)
    valores = {casilla: _num(texto, casilla) for casilla in CASILLAS.values()}
    return _armar_resultado(valores, _tipo_declaracion(texto))


def parsear_pdt_csv(csv_bytes):
    """Formato real confirmado (export alterno de SUNAT, 2026-09-23):
    cabecera 'Nro Ruc,Nro Orden,Formulario,Periodo,Fecha Presentacion,
    Indicador Rectif,Nro Casilla,Valor Casilla', una fila por casilla,
    numeros con punto decimal y sin separador de miles."""
    texto = csv_bytes.decode("utf-8-sig", errors="replace")
    lector = csv.DictReader(io.StringIO(texto))
    valores = {}
    indicador_rectif = None
    for fila in lector:
        casilla = (fila.get("Nro Casilla") or "").strip().lstrip("0") or "0"
        valor = (fila.get("Valor Casilla") or "").strip()
        indicador_rectif = fila.get("Indicador Rectif", indicador_rectif)
        try:
            valores[casilla] = float(valor)
        except ValueError:
            continue
    tipo = {"0": "ORIGINAL", "1": "RECTIFICATORIA"}.get((indicador_rectif or "").strip(), None)
    return _armar_resultado(valores, tipo)


def declaracion_final(empresa_id, periodo, cod_formulario="0621"):
    """La fila de declaraciones_pdt con el num_orden mas alto para ese
    empresa+periodo+formulario -- esa es SIEMPRE la version final (gane
    Original, Sustitutoria o Rectificatoria), sin necesidad de interpretar
    el texto 'Tipo de Declaracion'."""
    r = sb.table("declaraciones_pdt").select("*").eq(
        "empresa_id", empresa_id).eq("periodo", periodo).eq(
        "cod_formulario", cod_formulario).execute()
    filas = r.data or []
    if not filas:
        return None
    return max(filas, key=lambda f: int(f["num_orden"] or 0))


def obtener_evolucion_mes(empresa_id, periodo):
    """Devuelve lo declarado en el PDT final de ese mes, o None si la
    empresa no tiene PDT 621 para ese periodo (mes aun no declarado, o
    declarado pero sin detalle descargable y sin CSV de respaldo todavia).
    'compras_gravadas'/'compras_no_gravadas' van separados (no pre-sumados)
    porque el template real de Evolucion Anual/Mensual las muestra en
    columnas distintas ('Gravados'/'No gravados'), confirmado contra el
    Excel real de MEDISALUD (columnas H/I de la hoja AF (ULTIMO))."""
    decl = declaracion_final(empresa_id, periodo)
    url = decl and decl.get("url_detalle")
    if not url:
        return None
    resp = httpx.get(url, timeout=30)
    resp.raise_for_status()
    parser = parsear_pdt_csv if url.lower().split("?")[0].endswith(".csv") else parsear_pdt_pdf
    datos = parser(resp.content)
    return {
        "ventas": datos["ventas_total"],
        "compras": datos["compras_total"],
        "compras_gravadas": datos["compras_gravadas"],
        "compras_no_gravadas": datos["compras_no_gravadas"],
        "tipo_declaracion": datos["tipo_declaracion"],
        "num_orden": decl["num_orden"],
    }
