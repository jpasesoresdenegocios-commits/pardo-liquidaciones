"""Calculo de IGV-Renta mensual -- misma logica exacta que la vista
'Liquidacion de Impuestos' de Sistema Pardo (liqCargarPeriodo/liqRecalcular
en Sistema_Extractor_XML_Pardo_v2.html), incluido el fix del 2026-09-23
(cada fila se tiene que taguear con su registro real -- VENTAS usa
mto_exonerado+mto_inafecto, COMPRAS usa valor_adq_ng -- si no, 'No Gravadas'
de Ventas siempre sale en 0)."""
import re
from .supabase_client import sb


def _maybe(query):
    try:
        r = query.execute()
    except Exception:
        return None
    return r.data if r is not None else None


def _fecha_emision_aiso(s):
    m = re.match(r"^(\d{2})/(\d{2})/(\d{4})$", (s or "").strip())
    return f"{m.group(3)}-{m.group(2)}-{m.group(1)}" if m else ""


def _filtrar_mismo_periodo(filas, periodo):
    aaaa, mm = periodo[:4], periodo[4:6]
    out = []
    for f in filas:
        iso = _fecha_emision_aiso(f.get("fecha_emision"))
        if iso and iso[:4] == aaaa and iso[5:7] == mm:
            out.append(f)
    return out


def _resumen_igv(filas, registro):
    base_gravada = igv = no_gravado = total = 0.0
    for f in filas:
        base_gravada += float(f.get("bi_gravado") or 0)
        igv += float(f.get("igv") or 0)
        total += float(f.get("total") or 0)
        if registro == "VENTAS":
            no_gravado += float(f.get("mto_exonerado") or 0) + float(f.get("mto_inafecto") or 0)
        else:
            no_gravado += float(f.get("valor_adq_ng") or 0)
    return {
        "base_gravada": base_gravada, "igv": igv, "no_gravado": no_gravado,
        "total": total, "base_total": base_gravada + no_gravado, "n": len(filas),
    }


def obtener_filas(ruc, periodo, modo):
    data = _maybe(sb.table("almacen_datos").select("contenido").eq(
        "id_archivo", f"propuesta_detalle_{ruc}_{periodo}_{modo}").maybe_single())
    filas_crudas = data["contenido"] if (data and isinstance(data.get("contenido"), list)) else []
    filas = [{**f, "registro": modo} for f in filas_crudas]
    return _filtrar_mismo_periodo(filas, periodo)


def _periodo_anterior(periodo):
    a, m = int(periodo[:4]), int(periodo[4:6]) - 1
    if m < 1:
        m, a = 12, a - 1
    return f"{a}{m:02d}"


def calcular(ruc, empresa_id, periodo):
    """Devuelve el dict completo de la liquidacion IGV-Renta del periodo,
    sin guardar nada -- solo lectura + calculo, igual que 'Calcular' en la
    UI antes de apretar 'Guardar'."""
    filas_v = obtener_filas(ruc, periodo, "VENTAS")
    filas_c = obtener_filas(ruc, periodo, "COMPRAS")
    r_ventas = _resumen_igv(filas_v, "VENTAS")
    r_compras = _resumen_igv(filas_c, "COMPRAS")

    guardada = _maybe(sb.table("liquidaciones_impuestos").select("*").eq(
        "empresa_id", empresa_id).eq("periodo", periodo).maybe_single())

    arrastre = None
    if not guardada:
        prev = _periodo_anterior(periodo)
        arrastre = _maybe(sb.table("liquidaciones_impuestos").select("*").eq(
            "empresa_id", empresa_id).eq("periodo", prev).maybe_single())

    igv_saldo_favor = (guardada["igv_saldo_favor_anterior"] if guardada else
                        (max(0.0, -(arrastre.get("igv_resultante") or 0)) if arrastre else 0.0))
    renta_saldo_favor = (guardada["renta_saldo_favor_anterior"] if guardada else
                          (max(0.0, -(arrastre.get("renta_resultante") or 0)) if arrastre else 0.0))

    mes = int(periodo[4:6])
    anio_ejercicio = int(periodo[:4]) - (2 if mes <= 2 else 1)
    renta_tasa = None
    if guardada and guardada.get("renta_tasa_pct") is not None:
        renta_tasa = guardada["renta_tasa_pct"]
    elif arrastre and arrastre.get("renta_tasa_pct") is not None:
        renta_tasa = arrastre["renta_tasa_pct"]
    else:
        coef = _maybe(sb.table("renta_coeficientes").select("coeficiente_pct,anio_ejercicio").eq(
            "empresa_id", empresa_id).eq("anio_ejercicio", anio_ejercicio).order(
            "fecha_presentacion", desc=True).limit(1))
        if coef:
            renta_tasa = coef[0]["coeficiente_pct"]

    rp = _maybe(sb.table("retenciones_percepciones_igv").select("tipo,monto_propio").eq(
        "empresa_id", empresa_id).eq("periodo", periodo)) or []
    percep_periodo = sum(float(x["monto_propio"] or 0) for x in rp if x["tipo"] == "percepcion")
    reten_periodo = sum(float(x["monto_propio"] or 0) for x in rp if x["tipo"] == "retencion")

    igv_itan = guardada["igv_itan"] if guardada else 0.0
    igv_otros = guardada["igv_otros_creditos"] if guardada else 0.0
    igv_pagos_previos = guardada["igv_pagos_previos"] if guardada else 0.0
    igv_percep_ant = guardada["igv_percepciones_anteriores"] if guardada else 0.0
    igv_reten_ant = guardada["igv_retenciones_anteriores"] if guardada else 0.0
    renta_otros = guardada["renta_otros_creditos"] if guardada else 0.0
    renta_pagos_previos = guardada["renta_pagos_previos"] if guardada else 0.0

    running_igv = (r_ventas["igv"] - r_compras["igv"] - igv_saldo_favor - percep_periodo
                   - igv_percep_ant - reten_periodo - igv_reten_ant - igv_itan
                   - igv_otros - igv_pagos_previos)

    running_renta = None
    if renta_tasa is not None:
        renta_bruta = r_ventas["base_total"] * (float(renta_tasa) / 100)
        running_renta = renta_bruta - renta_saldo_favor - renta_otros - renta_pagos_previos

    return {
        "ruc": ruc, "periodo": periodo,
        "ventas": r_ventas, "compras": r_compras,
        "igv_saldo_favor_anterior": igv_saldo_favor,
        "igv_percepciones_periodo": percep_periodo,
        "igv_percepciones_anteriores": igv_percep_ant,
        "igv_retenciones_periodo": reten_periodo,
        "igv_retenciones_anteriores": igv_reten_ant,
        "igv_itan": igv_itan, "igv_otros_creditos": igv_otros,
        "igv_pagos_previos": igv_pagos_previos, "igv_resultante": running_igv,
        "renta_tasa_pct": renta_tasa,
        "renta_saldo_favor_anterior": renta_saldo_favor,
        "renta_otros_creditos": renta_otros, "renta_pagos_previos": renta_pagos_previos,
        "renta_resultante": running_renta,
    }
