"""Arma el .xlsx final calcado del formato real del despacho (capturas del
usuario, 2026-09-23): hoja IGV-RENTA (bloque RAZON SOCIAL/RUC/REGIMEN +
franjas naranja/celeste + resumen de deuda tributaria) y hoja AF con
EXACTAMENTE 2 graficos -- uno de BARRAS para "Evolucion Anual" (4 series:
Ventas/Compras x 2025/2026) y uno de LINEAS para "Evolucion Mensual" (2
series: Ventas/Compras del anio base) -- no 5 graficos separados como se
penso al principio a partir de un Excel de OTRA empresa con un formato mas
viejo."""
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.data_source import AxDataSource, StrRef
from openpyxl.chart.marker import Marker
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage

MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
         "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre"]

NARANJA = "FFE8720C"
CELESTE = "FF5BC0DE"
CELESTE_CLARO = "FFD9F2F8"
GRIS = "FFD9D9D9"


def _fmt(n):
    return round(float(n or 0), 2)


def _hex_a_argb(hex6):
    return "FF" + hex6.lstrip("#").upper()


def _franja(ws, fila, col_ini, col_fin, texto, color, blanco=True):
    for c in range(col_ini, col_fin + 1):
        celda = ws.cell(row=fila, column=c)
        celda.fill = PatternFill("solid", fgColor=color)
        celda.font = Font(bold=True, color="FFFFFFFF" if blanco else "FF000000")
    ws.cell(row=fila, column=col_ini, value=texto)
    ws.merge_cells(start_row=fila, start_column=col_ini, end_row=fila, end_column=col_fin)
    ws.cell(row=fila, column=col_ini).alignment = Alignment(horizontal="center")


def _encabezado_empresa(ws, branding, calc_ruc):
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 46
    for col in "DE":
        ws.column_dimensions[col].width = 14

    filas = [
        ("RAZÓN SOCIAL:", branding["nombre"]),
        ("RUC:", calc_ruc),
        ("RÉGIMEN TRIBUTARIO:", branding.get("regimen_tributario") or "—"),
    ]
    for i, (etq, val) in enumerate(filas):
        ws.cell(row=1 + i, column=2, value=etq).font = Font(bold=True, size=10)
        c = ws.cell(row=1 + i, column=3, value=val)
        c.font = Font(bold=True, size=10, color=_hex_a_argb("1E88C7"))
    return 5


def _agregar_logo(ws, logo_bytes, celda):
    if not logo_bytes:
        return
    try:
        img = PILImage.open(io.BytesIO(logo_bytes))
        img.thumbnail((190, 90))
        buf = io.BytesIO()
        img.convert("RGBA").save(buf, format="PNG")
        buf.seek(0)
        ws.add_image(XLImage(buf), celda)
    except Exception:
        pass


def _hoja_igv_renta(wb, calc, branding, logo_bytes):
    ws = wb.active
    ws.title = "IGV-RENTA"
    v, c = calc["ventas"], calc["compras"]

    fila = _encabezado_empresa(ws, branding, calc["ruc"])
    _agregar_logo(ws, logo_bytes, "F1")

    meses_es = {"01": "Enero", "02": "Febrero", "03": "Marzo", "04": "Abril", "05": "Mayo",
                "06": "Junio", "07": "Julio", "08": "Agosto", "09": "Setiembre",
                "10": "Octubre", "11": "Noviembre", "12": "Diciembre"}
    ws.cell(row=fila, column=2, value="PRELIQUIDACIÓN DE IMPUESTOS  MENSUAL").font = Font(bold=True, size=12)
    fila += 1
    ws.cell(row=fila, column=2, value=meses_es[calc["periodo"][4:]]).font = Font(bold=True)
    ws.cell(row=fila, column=3, value=calc["periodo"][:4]).font = Font(bold=True)
    fila += 2

    _franja(ws, fila, 2, 5, "IGV CUENTA PROPIA", NARANJA)
    fila += 1
    for c_idx, txt in [(2, "DESCRIPCIÓN"), (4, "BASE"), (5, "TRIBUTO")]:
        ws.cell(row=fila, column=c_idx, value=txt).font = Font(bold=True, color="FFFFFFFF")
        ws.cell(row=fila, column=c_idx).fill = PatternFill("solid", fgColor=CELESTE)
    ws.merge_cells(start_row=fila, start_column=2, end_row=fila, end_column=3)
    fila += 1

    def _fila_dato(etq_signo, etq_tipo, etq_desc, base, tributo):
        nonlocal fila
        ws.cell(row=fila, column=2, value=etq_signo)
        ws.cell(row=fila, column=2).font = Font(bold=True)
        ws.cell(row=fila, column=3, value=(etq_tipo + (f" — {etq_desc}" if etq_desc else "")) if etq_tipo else etq_desc)
        if base is not None:
            ws.cell(row=fila, column=4, value=_fmt(base)).number_format = "#,##0.00"
        if tributo is not None:
            ws.cell(row=fila, column=5, value=_fmt(tributo) if tributo != "-" else "-").number_format = "#,##0.00"
        fila += 1

    ws.cell(row=fila, column=2, value="VENTAS").font = Font(bold=True)
    for cc in range(2, 6):
        ws.cell(row=fila, column=cc).fill = PatternFill("solid", fgColor=CELESTE_CLARO)
    fila += 1
    _fila_dato("(+)", "Gravadas", "Ventas netas", v["base_gravada"], v["igv"])
    _fila_dato("", "", "Descuentos concedidos y/o devoluciones", None, 0)
    _fila_dato("(+)", "No Gravadas", "Ventas netas", v["no_gravado"], None)
    ws.cell(row=fila, column=2, value="TOTAL").font = Font(bold=True)
    for cc, val in [(4, v["base_total"]), (5, v["igv"])]:
        ws.cell(row=fila, column=cc, value=_fmt(val)).font = Font(bold=True)
        ws.cell(row=fila, column=cc).number_format = "#,##0.00"
    for cc in range(2, 6):
        ws.cell(row=fila, column=cc).fill = PatternFill("solid", fgColor=GRIS)
    fila += 2

    ws.cell(row=fila, column=2, value="COMPRAS").font = Font(bold=True)
    for cc in range(2, 6):
        ws.cell(row=fila, column=cc).fill = PatternFill("solid", fgColor=CELESTE_CLARO)
    fila += 1
    _fila_dato("(-)", "Gravadas", "Compras netas - Internas", c["base_gravada"], c["igv"])
    _fila_dato("(-)", "No Gravadas", "Compras netas - Internas", c["no_gravado"], None)
    ws.cell(row=fila, column=2, value="TOTAL").font = Font(bold=True)
    for cc, val in [(4, c["base_total"]), (5, c["igv"])]:
        ws.cell(row=fila, column=cc, value=_fmt(val)).font = Font(bold=True)
        ws.cell(row=fila, column=cc).number_format = "#,##0.00"
    for cc in range(2, 6):
        ws.cell(row=fila, column=cc).fill = PatternFill("solid", fgColor=GRIS)
    fila += 2

    _franja(ws, fila, 2, 5, "DETERMINACIÓN DE LA DEUDA TRIBUTARIA - IGV", NARANJA)
    fila += 1
    ws.cell(row=fila, column=2, value="DESCRIPCIÓN").font = Font(bold=True, color="FFFFFFFF")
    ws.cell(row=fila, column=5, value="TRIBUTO").font = Font(bold=True, color="FFFFFFFF")
    for cc in (2, 5):
        ws.cell(row=fila, column=cc).fill = PatternFill("solid", fgColor=CELESTE)
    ws.merge_cells(start_row=fila, start_column=2, end_row=fila, end_column=4)
    fila += 1

    def _fila_igv(signo, etq, val, negrita=False, gris=False):
        nonlocal fila
        ws.cell(row=fila, column=2, value=signo)
        cc = ws.cell(row=fila, column=3, value=etq)
        ws.merge_cells(start_row=fila, start_column=3, end_row=fila, end_column=4)
        cv = ws.cell(row=fila, column=5, value=_fmt(val))
        cv.number_format = "#,##0.00"
        if negrita:
            cc.font = Font(bold=True); cv.font = Font(bold=True)
        if gris:
            for c2 in range(2, 6):
                ws.cell(row=fila, column=c2).fill = PatternFill("solid", fgColor=GRIS)
        fila += 1

    _fila_igv("(-)", "Impuesto a pagar", v["igv"] - c["igv"] if (v["igv"] - c["igv"]) > 0 else 0)
    _fila_igv("(-)", "Saldo a favor del periodo anterior", calc["igv_saldo_favor_anterior"])
    _fila_igv("", "Tributo a pagar", max(0, calc["igv_resultante"]), negrita=True, gris=True)
    _fila_igv("(-)", "Percepciones declaradas en el periodo", calc["igv_percepciones_periodo"])
    _fila_igv("(-)", "Percepciones declaradas en periodos anteriores", calc["igv_percepciones_anteriores"])
    _fila_igv("(-)", "Retenciones declaradas en el periodo", calc["igv_retenciones_periodo"])
    _fila_igv("(-)", "Retenciones declaradas en periodos anteriores", calc["igv_retenciones_anteriores"])
    _fila_igv("(-)", "Impuesto Temporal a los Activos Netos (ley n° 28424)", calc["igv_itan"])
    _fila_igv("(-)", "Otros créditos permitido por ley", calc["igv_otros_creditos"])
    _fila_igv("", "Total tributo a pagar", max(0, calc["igv_resultante"]), negrita=True, gris=True)
    _fila_igv("(-)", "Pagos previos", calc["igv_pagos_previos"])

    _franja(ws, fila, 2, 4, "TOTAL DEUDA TRIBUTARIA DEL IGV", NARANJA)
    ws.cell(row=fila, column=5).fill = PatternFill("solid", fgColor=NARANJA)
    ws.cell(row=fila, column=5, value=_fmt(max(0, calc["igv_resultante"])))
    ws.cell(row=fila, column=5).font = Font(bold=True, color="FFFFFFFF")
    ws.cell(row=fila, column=5).number_format = "#,##0.00"
    fila += 2

    _franja(ws, fila, 2, 5, "IMPUESTO A LA RENTA - 3ERA CATEGORIA", CELESTE)
    fila += 1
    ws.cell(row=fila, column=3, value="Tasa / coeficiente (%)").font = Font(italic=True)
    ws.cell(row=fila, column=5, value=calc["renta_tasa_pct"] if calc["renta_tasa_pct"] is not None else "—")
    fila += 1
    ws.cell(row=fila, column=2, value="DESCRIPCIÓN").font = Font(bold=True)
    ws.cell(row=fila, column=4, value="BASE").font = Font(bold=True)
    ws.cell(row=fila, column=5, value="TRIBUTO").font = Font(bold=True)
    fila += 1
    ws.cell(row=fila, column=2, value="Ingresos Netos")
    ws.cell(row=fila, column=4, value=_fmt(v["base_total"])).number_format = "#,##0.00"
    renta_bruta = (calc["renta_resultante"] + calc["renta_saldo_favor_anterior"]
                   + calc["renta_otros_creditos"] + calc["renta_pagos_previos"]) \
        if calc["renta_resultante"] is not None else 0
    ws.cell(row=fila, column=5, value=_fmt(renta_bruta)).number_format = "#,##0.00"
    fila += 2

    _franja(ws, fila, 2, 5, "DETERMINACIÓN DE LA DEUDA TRIBUTARIA - RENTA", CELESTE)
    fila += 1
    ws.cell(row=fila, column=2, value="DESCRIPCIÓN").font = Font(bold=True)
    ws.cell(row=fila, column=5, value="TRIBUTO").font = Font(bold=True)
    fila += 1
    renta_result = calc["renta_resultante"] if calc["renta_resultante"] is not None else 0
    _fila_igv("(-)", "Impuesto a pagar", renta_bruta)
    _fila_igv("(-)", "Saldo a favor del periodo anterior", calc["renta_saldo_favor_anterior"])
    _fila_igv("", "Saldo a favor" if renta_result < 0 else "Tributo a pagar", renta_result, negrita=True, gris=True)
    _fila_igv("(-)", "Otros créditos permitido por ley", calc["renta_otros_creditos"])
    _fila_igv("(-)", "Pagos previos", calc["renta_pagos_previos"])

    _franja(ws, fila, 2, 4, "TOTAL DEUDA TRIBUTARIA DE RENTA DE 3ERA CATEGORIA", CELESTE)
    ws.cell(row=fila, column=5).fill = PatternFill("solid", fgColor=CELESTE)
    ws.cell(row=fila, column=5, value=_fmt(max(0, renta_result)))
    ws.cell(row=fila, column=5).font = Font(bold=True, color="FFFFFFFF")
    ws.cell(row=fila, column=5).number_format = "#,##0.00"
    fila += 2

    _franja(ws, fila, 2, 5, "RESUMEN DEUDA TRIBUTARIA", CELESTE_CLARO, blanco=False)
    fila += 1
    for c_idx, txt in [(2, "TRIBUTO"), (3, "CÓDIGO"), (4, "IMPORTE"), (5, "ESTADO")]:
        ws.cell(row=fila, column=c_idx, value=txt).font = Font(bold=True)
    fila += 1
    igv_pagar = max(0, calc["igv_resultante"])
    renta_pagar = max(0, renta_result)
    ws.cell(row=fila, column=2, value="IGV"); ws.cell(row=fila, column=3, value="1011")
    ws.cell(row=fila, column=4, value=_fmt(igv_pagar)).number_format = "#,##0.00"
    ws.cell(row=fila, column=5, value="POR PAGAR" if igv_pagar > 0 else "SALDO A FAVOR")
    fila += 1
    ws.cell(row=fila, column=2, value="RENTA"); ws.cell(row=fila, column=3, value="3121")
    ws.cell(row=fila, column=4, value=_fmt(renta_pagar) if renta_pagar else "-")
    if renta_pagar:
        ws.cell(row=fila, column=4).number_format = "#,##0.00"
    ws.cell(row=fila, column=5, value="POR PAGAR" if renta_pagar > 0 else "SALDO A FAVOR")
    fila += 1
    ws.cell(row=fila, column=2, value="TOTAL").font = Font(bold=True)
    ws.cell(row=fila, column=4, value=_fmt(igv_pagar + renta_pagar)).font = Font(bold=True)
    ws.cell(row=fila, column=4).number_format = "#,##0.00"

    return ws


def _set_categorias_texto(chart, ws_ref_formula):
    for serie in chart.series:
        serie.cat = AxDataSource(strRef=StrRef(f=ws_ref_formula))


def _hoja_evolucion(wb, branding, calc_ruc, anio_base, anio_anterior, datos_por_anio):
    """EXACTO como las capturas: bloque EVOLUCION ANUAL (tabla Ventas +
    tabla Compras lado a lado + 1 grafico de BARRAS con 4 series) y bloque
    EVOLUCION MENSUAL (tabla Ventas + tabla Compras + 1 grafico de LINEAS
    con 2 series). Las columnas 'Variacion'/'%' usan FORMULAS reales de
    Excel (pedido explicito del usuario 2026-09-23: 'ten en cuenta que en
    algunos hay formula'), no valores ya calculados en Python -- asi si
    alguien corrige un numero a mano, el resto se recalcula solo."""
    ws = wb.create_sheet("AF")
    for col, w in [("A", 12), ("B", 13), ("C", 13), ("D", 13), ("E", 13),
                   ("F", 13), ("G", 13), ("H", 11), ("I", 11)]:
        ws.column_dimensions[col].width = w

    ws.cell(row=1, column=1, value="RAZÓN SOCIAL:").font = Font(bold=True, size=10)
    ws.cell(row=1, column=2, value=branding["nombre"]).font = Font(bold=True, size=10, color=_hex_a_argb("1E88C7"))
    ws.cell(row=2, column=1, value="RUC:").font = Font(bold=True, size=10)
    ws.cell(row=2, column=2, value=calc_ruc).font = Font(bold=True, size=10)
    ws.cell(row=3, column=1, value="RÉGIMEN TRIBUTARIO:").font = Font(bold=True, size=10)
    ws.cell(row=3, column=2, value=branding.get("regimen_tributario") or "—").font = Font(bold=True, size=10)

    def sval(anio, mes_idx, clave):
        d = datos_por_anio.get(anio, [None] * 12)[mes_idx]
        return _fmt(d[clave]) if d else 0

    fila = 5
    _franja(ws, fila, 1, 9, "EVOLUCIÓN ANUAL", NARANJA)
    fila += 1
    fila_hdr1 = fila
    ws.cell(row=fila, column=1, value="Mes").font = Font(bold=True)
    ws.cell(row=fila, column=2, value=f"Ventas {anio_anterior or ''}").font = Font(bold=True)
    ws.cell(row=fila, column=3, value=f"Ventas {anio_base}").font = Font(bold=True)
    ws.cell(row=fila, column=4, value="Variación").font = Font(bold=True)
    ws.cell(row=fila, column=5, value="%").font = Font(bold=True)
    fila += 1
    fila_datos1_ini = fila
    for i, mes in enumerate(MESES):
        ws.cell(row=fila, column=1, value=mes)
        ws.cell(row=fila, column=2, value=sval(anio_anterior, i, "ventas") if anio_anterior else 0)
        ws.cell(row=fila, column=3, value=sval(anio_base, i, "ventas"))
        r = fila
        ws.cell(row=r, column=4, value=f"=C{r}-B{r}").number_format = "#,##0.00"
        ws.cell(row=r, column=5, value='=IF(B{0}=0,"",D{0}/B{0})'.format(r)).number_format = "0.0%"
        fila += 1
    fila_datos1_fin = fila - 1
    ws.cell(row=fila, column=1, value="Total General").font = Font(bold=True)
    ws.cell(row=fila, column=2, value=f"=SUM(B{fila_datos1_ini}:B{fila_datos1_fin})").font = Font(bold=True)
    ws.cell(row=fila, column=3, value=f"=SUM(C{fila_datos1_ini}:C{fila_datos1_fin})").font = Font(bold=True)
    for cc in (2, 3):
        ws.cell(row=fila, column=cc).number_format = "#,##0.00"
    fila_total1 = fila

    fh = fila_hdr1
    ws.cell(row=fh, column=7, value=f"Compras {anio_anterior or ''}").font = Font(bold=True)
    ws.cell(row=fh, column=8, value=f"Compras {anio_base}").font = Font(bold=True)
    ws.cell(row=fh, column=9, value="Variación").font = Font(bold=True)
    for i in range(12):
        r = fila_datos1_ini + i
        ws.cell(row=r, column=7, value=sval(anio_anterior, i, "compras") if anio_anterior else 0)
        ws.cell(row=r, column=8, value=sval(anio_base, i, "compras"))
        ws.cell(row=r, column=9, value=f"=H{r}-G{r}").number_format = "#,##0.00"
    ws.cell(row=fila_total1, column=7, value=f"=SUM(G{fila_datos1_ini}:G{fila_datos1_fin})").font = Font(bold=True)
    ws.cell(row=fila_total1, column=8, value=f"=SUM(H{fila_datos1_ini}:H{fila_datos1_fin})").font = Font(bold=True)
    for cc in (7, 8):
        ws.cell(row=fila_total1, column=cc).number_format = "#,##0.00"

    cat_formula = f"'AF'!$A${fila_datos1_ini}:$A${fila_datos1_fin}"
    chart_anual = BarChart()
    chart_anual.type = "col"
    chart_anual.title = f"EVOLUCIÓN DE VENTAS - COMPRAS\nPERIODO: {anio_anterior or ''}-{anio_base}".strip("- ")
    chart_anual.height, chart_anual.width = 10, 20
    for col in (2, 3, 7, 8):
        ref = Reference(ws, min_col=col, min_row=fila_hdr1, max_row=fila_datos1_fin)
        chart_anual.add_data(ref, titles_from_data=True)
    _set_categorias_texto(chart_anual, cat_formula)
    ws.add_chart(chart_anual, "K5")

    fila = fila_total1 + 2

    _franja(ws, fila, 1, 9, "EVOLUCIÓN MENSUAL", NARANJA)
    fila += 1
    fila_hdr2 = fila
    ws.cell(row=fila, column=1, value="Mes").font = Font(bold=True)
    ws.cell(row=fila, column=2, value="Ventas").font = Font(bold=True)
    ws.cell(row=fila, column=3, value="Variación").font = Font(bold=True)
    ws.cell(row=fila, column=4, value="%").font = Font(bold=True)
    ws.cell(row=fila, column=7, value="Compras").font = Font(bold=True)
    ws.cell(row=fila, column=8, value="Variación").font = Font(bold=True)
    ws.cell(row=fila, column=9, value="%").font = Font(bold=True)
    fila += 1
    fila_datos2_ini = fila
    for i, mes in enumerate(MESES):
        r = fila
        ws.cell(row=r, column=1, value=mes)
        ws.cell(row=r, column=2, value=sval(anio_base, i, "ventas"))
        ws.cell(row=r, column=7, value=sval(anio_base, i, "compras"))
        if i > 0:
            ws.cell(row=r, column=3, value=f"=B{r}-B{r-1}").number_format = "#,##0.00"
            ws.cell(row=r, column=4, value='=IF(B{0}=0,"",C{1}/B{0})'.format(r - 1, r)).number_format = "0.0%"
            ws.cell(row=r, column=8, value=f"=G{r}-G{r-1}").number_format = "#,##0.00"
            ws.cell(row=r, column=9, value='=IF(G{0}=0,"",H{1}/G{0})'.format(r - 1, r)).number_format = "0.0%"
        fila += 1
    fila_datos2_fin = fila - 1
    ws.cell(row=fila, column=1, value="Total General").font = Font(bold=True)
    ws.cell(row=fila, column=2, value=f"=SUM(B{fila_datos2_ini}:B{fila_datos2_fin})").font = Font(bold=True)
    ws.cell(row=fila, column=7, value=f"=SUM(G{fila_datos2_ini}:G{fila_datos2_fin})").font = Font(bold=True)
    for cc in (2, 7):
        ws.cell(row=fila, column=cc).number_format = "#,##0.00"

    cat_formula2 = f"'AF'!$A${fila_datos2_ini}:$A${fila_datos2_fin}"
    chart_mensual = LineChart()
    chart_mensual.title = f"EVOLUCIÓN DE VENTAS - COMPRAS\nPERIODO: {anio_base}"
    chart_mensual.height, chart_mensual.width = 10, 20
    ref_v = Reference(ws, min_col=2, min_row=fila_hdr2, max_row=fila_datos2_fin)
    ref_c = Reference(ws, min_col=7, min_row=fila_hdr2, max_row=fila_datos2_fin)
    chart_mensual.add_data(ref_v, titles_from_data=True)
    chart_mensual.add_data(ref_c, titles_from_data=True)
    _set_categorias_texto(chart_mensual, cat_formula2)
    for s in chart_mensual.series:
        s.marker = Marker(symbol="circle", size=5)
        s.smooth = False
    ws.add_chart(chart_mensual, "K27")

    return ws


def construir_workbook(ruc, calc, branding, logo_bytes, anio_base, anio_anterior, datos_por_anio):
    wb = Workbook()
    _hoja_igv_renta(wb, calc, branding, logo_bytes)
    _hoja_evolucion(wb, branding, ruc, anio_base, anio_anterior, datos_por_anio)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
