"""Lee el PDF de detalle del PDT 621 (IGV-Renta Mensual) YA GUARDADO en
Supabase Storage (declaraciones_pdt.url_detalle) y saca las casillas reales
que declaro el contribuyente -- para que la Evolucion Compras-Ventas se
llene con lo REALMENTE declarado, no con la propuesta SIRE (que puede traer
comprobantes que el contador nunca tomo).

Confirmado en vivo el 2026-09-23 con un PDF real de DIABETES (202608): el
texto que saca pdfplumber es 'etiqueta casilla valor [casilla2 valor2]' muy
regular, ej. 'Ventas Netas 100 283,772.00 101 51,079.00'. Las casillas del
formulario 0621 son ESTANDAR de SUNAT (mismo numero en todo el pais), asi
que estos numeros de casilla no cambian entre empresas ni periodos.

Regla de "cual declaracion es la valida" (pedido explicito del usuario):
una empresa puede tener Original + Sustitutoria + Rectificatoria para el
mismo periodo -- SUNAT asigna num_orden en orden creciente, asi que la de
mayor num_orden es siempre la mas reciente/definitiva, sea cual sea su
'Tipo de Declaracion' (ese texto tambien se extrae, solo para mostrarlo)."""
import re
import httpx
import pdfplumber
import io
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


def parsear_pdt_pdf(pdf_bytes):
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        texto = "\n".join(p.extract_text() or "" for p in pdf.pages)

    ventas_gravadas = _num(texto, CASILLAS["ventas_gravadas"])
    ventas_no_gravadas = _num(texto, CASILLAS["ventas_no_gravadas"])
    compras_gravadas = _num(texto, CASILLAS["compras_gravadas"])
    compras_no_gravadas = _num(texto, CASILLAS["compras_no_gravadas"])

    return {
        "tipo_declaracion": _tipo_declaracion(texto),
        "ventas_gravadas": ventas_gravadas,
        "ventas_no_gravadas": ventas_no_gravadas,
        "ventas_total": _num(texto, CASILLAS["ingresos_netos"]) or (ventas_gravadas + ventas_no_gravadas),
        "igv_ventas": _num(texto, CASILLAS["igv_ventas"]),
        "compras_gravadas": compras_gravadas,
        "compras_no_gravadas": compras_no_gravadas,
        # OJO: no existe una sola casilla oficial de "compras total" en el
        # PDT 621 -- se suma la base gravada + no gravada (nacional). Si la
        # empresa tiene compras IMPORTADAS (casillas 114/119/122) quedarian
        # afuera; para las ~65 empresas del despacho esto no se ha visto
        # todavia, pero queda documentado por si aparece.
        "compras_total": compras_gravadas + compras_no_gravadas,
        "igv_compras": _num(texto, CASILLAS["igv_compras"]),
        "renta_resultante": _num(texto, CASILLAS["renta_resultante"]),
    }


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
    """Devuelve {ventas, compras, tipo_declaracion} declarados en el PDT
    final de ese mes, o None si la empresa no tiene PDT 621 para ese
    periodo (ej. mes aun no declarado)."""
    decl = declaracion_final(empresa_id, periodo)
    if not decl or not decl.get("url_detalle"):
        return None
    resp = httpx.get(decl["url_detalle"], timeout=30)
    resp.raise_for_status()
    datos = parsear_pdt_pdf(resp.content)
    return {
        "ventas": datos["ventas_total"],
        "compras": datos["compras_total"],
        "tipo_declaracion": datos["tipo_declaracion"],
        "num_orden": decl["num_orden"],
    }
