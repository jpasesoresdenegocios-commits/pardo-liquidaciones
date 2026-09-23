"""Calculo de IGV-Renta mensual -- misma logica exacta que la vista
'Liquidacion de Impuestos' de Sistema Pardo (liqCargarPeriodo/liqRecalcular
en Sistema_Extractor_XML_Pardo_v2.html), incluido el fix del 2026-09-23
(cada fila se tiene que taguear con su registro real -- VENTAS usa
mto_exonerado+mto_inafecto, COMPRAS usa valor_adq_ng -- si no, 'No Gravadas'
de Ventas siempre sale en 0), mas la regla RMT de 300 UIT (2026-09-23,
verificada contra el Excel real de Diabetes: D59='=300*5500',
D61='=+IF(D34>D59,"SUPERO 1.5% O COEF","-")')."""
import re
from .supabase_client import sb

# UIT por anio de ejercicio -- el Excel real de Diabetes usa 5500 (literal)
# para el ejercicio 2026. Si aparece un anio nuevo sin UIT publicada todavia,
# se usa la ultima conocida (evita romper el calculo, el contador la corrige
# a mano en el Excel igual que hoy).
UIT_POR_ANIO = {2026: 5500}


def uit_del_anio(anio):
    if anio in UIT_POR_ANIO:
        return UIT_POR_ANIO[anio]
    return UIT_POR_ANIO[max(UIT_POR_ANIO)]


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


def _bucket_tasa(bi, igv):
    """Clasifica una fila de compras por tasa de IGV real (igv/bi_gravado) --
    no existe un campo 'tasa' en el SIRE, asi que se deriva por fila. La Ley
    N 31556 dejo tasas reducidas (10.5%) para ciertos rubros, por eso el
    template real separa 'Gravadas 18%' de 'Gravadas 10.5%'."""
    if not bi:
        return "18"
    rate = igv / bi
    return "105" if abs(rate - 0.105) < abs(rate - 0.18) else "18"


def _resumen_igv(filas, registro):
    base_gravada = igv = no_gravado = total = 0.0
    por_tasa = {"18": 0.0, "105": 0.0}
    igv_por_tasa = {"18": 0.0, "105": 0.0}
    for f in filas:
        bi = float(f.get("bi_gravado") or 0)
        fi = float(f.get("igv") or 0)
        base_gravada += bi
        igv += fi
        total += float(f.get("total") or 0)
        if registro == "VENTAS":
            no_gravado += float(f.get("mto_exonerado") or 0) + float(f.get("mto_inafecto") or 0)
        else:
            no_gravado += float(f.get("valor_adq_ng") or 0)
            tasa = _bucket_tasa(bi, fi)
            por_tasa[tasa] += bi
            igv_por_tasa[tasa] += fi
    resumen = {
        "base_gravada": base_gravada, "igv": igv, "no_gravado": no_gravado,
        "total": total, "base_total": base_gravada + no_gravado, "n": len(filas),
    }
    if registro != "VENTAS":
        # Solo compras: base/igv separados por tasa, para el detalle
        # "Gravadas 18% / Gravadas 10.5%" del template real. 'importadas'
        # siempre en 0 -- el SIRE no trae un campo que distinga compras
        # importadas de internas, y ninguna empresa del despacho las usa
        # todavia (el template las deja como fila oculta hasta que aparezcan).
        resumen["por_tasa"] = {
            "18": {"internas": por_tasa["18"], "importadas": 0.0, "igv": igv_por_tasa["18"]},
            "105": {"internas": por_tasa["105"], "importadas": 0.0, "igv": igv_por_tasa["105"]},
        }
    return resumen


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


def _normalizar_fraccion_pct(v):
    """Las 3 fuentes de coeficiente historico (liquidaciones_impuestos
    guardada, su arrastre, o renta_coeficientes) guardan la tasa como
    PORCENTAJE (1.5 = 1.5%), nunca como fraccion -- se normaliza una sola
    vez aqui. Bug real detectado 2026-09-23: sin este fix, regla_300_uit()
    comparaba 1.5 contra 0.015 y devolvia 1.5 sin convertir, asi que el
    Excel mostraba "150.00%" en la celda de Coeficiente y la Renta salia
    100x mas grande (507,479 en vez de 5,075)."""
    if v is None:
        return None
    v = float(v)
    return v / 100 if v > 1 else v


def regla_300_uit(ventas_acumuladas_anio, anio, coef_historico):
    """Regimen MYPE Tributario (RMT): mientras los ingresos netos
    ACUMULADOS del ejercicio no superen 300 UIT, el pago a cuenta es 1% de
    los ingresos del mes. Apenas se supera en cualquier mes, de ahi en
    adelante en todo el ejercicio se paga el MAYOR entre el coeficiente
    propio y 1.5% (Art. 85 LIR modificado). Verificado contra el Excel real:
    D59=300*UIT, D61 marca 'SUPERO 1.5% O COEF' cuando D34 (ventas acum.) >
    D59 -- y ese mes E59(Coeficiente)=1.5% en vez de 0.

    'coef_historico' DEBE venir ya normalizado a fraccion (0.015, no 1.5)
    -- ver _normalizar_fraccion_pct()."""
    limite = 300 * uit_del_anio(anio)
    supero_300_uit = ventas_acumuladas_anio > limite
    if not supero_300_uit:
        return {"tasa": 0.01, "usa_coeficiente": False, "supero_300_uit": False, "limite_300_uit": limite}
    tasa = max(float(coef_historico or 0), 0.015)
    return {"tasa": tasa, "usa_coeficiente": True, "supero_300_uit": True, "limite_300_uit": limite}


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
    coef_historico = None
    if guardada and guardada.get("renta_tasa_pct") is not None:
        coef_historico = guardada["renta_tasa_pct"]
    elif arrastre and arrastre.get("renta_tasa_pct") is not None:
        coef_historico = arrastre["renta_tasa_pct"]
    else:
        coef = _maybe(sb.table("renta_coeficientes").select("coeficiente_pct,anio_ejercicio").eq(
            "empresa_id", empresa_id).eq("anio_ejercicio", anio_ejercicio).order(
            "fecha_presentacion", desc=True).limit(1))
        if coef:
            coef_historico = coef[0]["coeficiente_pct"]
    coef_historico = _normalizar_fraccion_pct(coef_historico)

    # Regla 300 UIT: ingresos netos acumulados del ejercicio (Enero..mes
    # actual) sumando lo declarado en cada PDT 621 ya presentado. Usa la
    # MISMA fuente (PDT, no propuesta SIRE) que la hoja Evolucion, para que
    # el aviso "SUPERO 1.5% O COEF" del Excel siempre cuadre con esta regla.
    from .pdt_parser import obtener_evolucion_mes
    anio_periodo = int(periodo[:4])
    ventas_acum = 0.0
    for m in range(1, mes + 1):
        p = f"{anio_periodo}{m:02d}"
        try:
            d = obtener_evolucion_mes(empresa_id, p)
        except Exception:
            d = None
        if d:
            ventas_acum += float(d.get("ventas") or 0)
    if ventas_acum <= 0:
        ventas_acum = r_ventas["base_total"]  # sin PDT declarado todavia: usa la propuesta del propio mes

    regla_uit = regla_300_uit(ventas_acum, anio_periodo, coef_historico)
    renta_tasa = regla_uit["tasa"]

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

    renta_bruta = round(r_ventas["base_total"] * renta_tasa, 0)
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
        "renta_tasa_pct": renta_tasa * 100,
        "renta_coeficiente_historico_pct": coef_historico,
        "renta_regla_300_uit": regla_uit,
        "renta_ventas_acumuladas_anio": ventas_acum,
        "renta_saldo_favor_anterior": renta_saldo_favor,
        "renta_otros_creditos": renta_otros, "renta_pagos_previos": renta_pagos_previos,
        "renta_resultante": running_renta,
    }
