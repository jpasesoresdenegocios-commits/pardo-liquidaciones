"""Arma el .xlsx final: hoja IGV-RENTA (calcada del formato que ya usaba
Sistema Pardo) + hoja AF con los 5 graficos de evolucion, EN EL MISMO ORDEN
que el Excel real del despacho (Ventas Anual, Compras Anual, Ventas
Mensual, Compras Mensual, Ventas-Compras Mensual), con el color/logo de
cada empresa."""
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import BarChart, Reference
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage

MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Set", "Oct", "Nov", "Dic"]


def _fmt(n):
    return round(float(n or 0), 2)


def _hex_a_argb(hex6):
    return "FF" + hex6.upper()


def _hoja_igv_renta(wb, calc, branding, logo_bytes):
    ws = wb.active
    ws.title = "IGV-RENTA"
    color = _hex_a_argb(branding["color"])
    fill_header = PatternFill("solid", fgColor=color)
    font_header = Font(bold=True, color="FFFFFFFF", size=13)

    ws.column_dimensions["A"].width = 42
    for col in "BCD":
        ws.column_dimensions[col].width = 16

    if logo_bytes:
        try:
            img = PILImage.open(io.BytesIO(logo_bytes))
            img.thumbnail((140, 70))
            buf = io.BytesIO()
            img.convert("RGB").save(buf, format="PNG")
            buf.seek(0)
            xlimg = XLImage(buf)
            ws.add_image(xlimg, "F1")
        except Exception:
            pass

    ws["A1"] = f"LIQUIDACIÓN DE IMPUESTOS MENSUAL — {branding['nombre']}"
    ws["A1"].font = Font(bold=True, size=14, color=color)
    ws["A2"] = f"RUC {calc['ruc']}"
    ws["C2"] = f"Período {calc['periodo'][:4]}-{calc['periodo'][4:]}"

    r = 4
    v, c = calc["ventas"], calc["compras"]
    filas = [
        ("VENTAS", "Base", "IGV", "Total"),
        ("Gravadas", v["base_gravada"], v["igv"], v["base_gravada"] + v["igv"]),
        ("No Gravadas", v["no_gravado"], "-", v["no_gravado"]),
        ("TOTAL VENTAS", v["base_total"], v["igv"], v["total"]),
        ("", "", "", ""),
        ("COMPRAS", "Base", "IGV", "Total"),
        ("Gravadas", c["base_gravada"], c["igv"], c["base_gravada"] + c["igv"]),
        ("No Gravadas", c["no_gravado"], "-", c["no_gravado"]),
        ("TOTAL COMPRAS", c["base_total"], c["igv"], c["total"]),
        ("", "", "", ""),
        ("IGV — DETERMINACIÓN DE LA DEUDA", "", "", ""),
        ("IGV Ventas", "", _fmt(v["igv"]), ""),
        ("IGV Compras (crédito fiscal)", "", -_fmt(c["igv"]), ""),
        ("Saldo a favor del período anterior", "", -_fmt(calc["igv_saldo_favor_anterior"]), ""),
        ("Percepciones del período", "", -_fmt(calc["igv_percepciones_periodo"]), ""),
        ("Percepciones de períodos anteriores", "", -_fmt(calc["igv_percepciones_anteriores"]), ""),
        ("Retenciones del período", "", -_fmt(calc["igv_retenciones_periodo"]), ""),
        ("Retenciones de períodos anteriores", "", -_fmt(calc["igv_retenciones_anteriores"]), ""),
        ("ITAN", "", -_fmt(calc["igv_itan"]), ""),
        ("Otros créditos", "", -_fmt(calc["igv_otros_creditos"]), ""),
        ("Pagos previos (rectificatorias)", "", -_fmt(calc["igv_pagos_previos"]), ""),
        ("IGV POR PAGAR" if calc["igv_resultante"] > 0 else "SALDO A FAVOR (IGV)",
         "", abs(_fmt(calc["igv_resultante"])), ""),
        ("", "", "", ""),
        ("RENTA 3RA CATEGORÍA — DETERMINACIÓN", "", "", ""),
        ("Ingresos netos (base)", "", _fmt(v["base_total"]), ""),
        ("Tasa / coeficiente (%)", "", calc["renta_tasa_pct"] or "-", ""),
        ("Saldo a favor del período anterior", "", -_fmt(calc["renta_saldo_favor_anterior"]), ""),
        ("Otros créditos", "", -_fmt(calc["renta_otros_creditos"]), ""),
        ("Pagos previos (rectificatorias)", "", -_fmt(calc["renta_pagos_previos"]), ""),
    ]
    if calc["renta_resultante"] is not None:
        filas.append((
            "RENTA POR PAGAR" if calc["renta_resultante"] > 0 else "SALDO A FAVOR (RENTA)",
            "", abs(_fmt(calc["renta_resultante"])), ""))

    filas_encabezado = {4, 9, 15, 28}
    for offset, fila in enumerate(filas):
        for col_idx, valor in enumerate(fila, start=1):
            celda = ws.cell(row=r + offset, column=col_idx, value=valor)
            if col_idx in (2, 3, 4) and isinstance(valor, (int, float)):
                celda.number_format = "#,##0.00"
                celda.alignment = Alignment(horizontal="right")
        fila_num = r + offset
        if fila_num - r in {0, 5, 10, 23} or fila[0] in (
                "VENTAS", "COMPRAS", "IGV — DETERMINACIÓN DE LA DEUDA",
                "RENTA 3RA CATEGORÍA — DETERMINACIÓN"):
            for col_idx in range(1, 5):
                ws.cell(row=fila_num, column=col_idx).fill = fill_header
                ws.cell(row=fila_num, column=col_idx).font = Font(bold=True, color="FFFFFFFF")

    return ws


def _serie_mes(datos_mes, clave):
    return [datos_mes[m][clave] if datos_mes[m] else 0 for m in range(12)]


def _agregar_grafico(ws, titulo, celda_ancla, categorias_ref, series_defs, alto=15, ancho=9):
    chart = BarChart()
    chart.type = "col"
    chart.title = titulo
    chart.y_axis.title = "S/"
    chart.height, chart.width = alto, ancho
    for ref, nombre in series_defs:
        chart.add_data(ref, titles_from_data=True)
    chart.set_categories(categorias_ref)
    if series_defs:
        chart.series[-1].tx = None
    ws.add_chart(chart, celda_ancla)
    return chart


def _hoja_evolucion(wb, ruc, branding, anio_base, anio_anterior, datos_por_anio):
    ws = wb.create_sheet("AF")
    ws["A1"] = f"EVOLUCIÓN COMPRAS-VENTAS — {branding['nombre']} (declarado en PDT)"
    ws["A1"].font = Font(bold=True, size=13, color=_hex_a_argb(branding["color"]))

    encabezado = ["Mes"]
    if anio_anterior:
        encabezado += [f"Ventas {anio_anterior}", f"Compras {anio_anterior}"]
    encabezado += [f"Ventas {anio_base}", f"Compras {anio_base}"]
    ws.append([])
    fila_encabezado = 3
    for i, h in enumerate(encabezado, start=1):
        ws.cell(row=fila_encabezado, column=i, value=h).font = Font(bold=True)

    for i, mes in enumerate(MESES):
        fila = fila_encabezado + 1 + i
        col = 1
        ws.cell(row=fila, column=col, value=mes); col += 1
        if anio_anterior:
            dm_ant = datos_por_anio[anio_anterior][i]
            ws.cell(row=fila, column=col, value=_fmt(dm_ant["ventas"]) if dm_ant else 0); col += 1
            ws.cell(row=fila, column=col, value=_fmt(dm_ant["compras"]) if dm_ant else 0); col += 1
        dm_base = datos_por_anio[anio_base][i]
        ws.cell(row=fila, column=col, value=_fmt(dm_base["ventas"]) if dm_base else 0); col += 1
        ws.cell(row=fila, column=col, value=_fmt(dm_base["compras"]) if dm_base else 0); col += 1

    fin = fila_encabezado + 12
    cat_ref = Reference(ws, min_col=1, min_row=fila_encabezado + 1, max_row=fin)

    col_v_ant = 2 if anio_anterior else None
    col_c_ant = 3 if anio_anterior else None
    col_v_base = 4 if anio_anterior else 2
    col_c_base = 5 if anio_anterior else 3

    # 1) Ventas anual (comparando 2 años, si hay año de comparación)
    if anio_anterior:
        chart1 = BarChart(); chart1.type = "col"
        chart1.title = f"EVOLUCIÓN DE VENTAS ANUAL {anio_anterior}-{anio_base}"
        chart1.add_data(Reference(ws, min_col=col_v_ant, min_row=fila_encabezado, max_row=fin), titles_from_data=True)
        chart1.add_data(Reference(ws, min_col=col_v_base, min_row=fila_encabezado, max_row=fin), titles_from_data=True)
        chart1.set_categories(cat_ref)
        ws.add_chart(chart1, "A19")
    else:
        chart1 = BarChart(); chart1.type = "col"
        chart1.title = f"EVOLUCIÓN DE VENTAS ANUAL {anio_base}"
        chart1.add_data(Reference(ws, min_col=col_v_base, min_row=fila_encabezado, max_row=fin), titles_from_data=True)
        chart1.set_categories(cat_ref)
        ws.add_chart(chart1, "A19")

    # 2) Compras anual
    if anio_anterior:
        chart2 = BarChart(); chart2.type = "col"
        chart2.title = f"EVOLUCIÓN DE COMPRAS ANUAL {anio_anterior}-{anio_base}"
        chart2.add_data(Reference(ws, min_col=col_c_ant, min_row=fila_encabezado, max_row=fin), titles_from_data=True)
        chart2.add_data(Reference(ws, min_col=col_c_base, min_row=fila_encabezado, max_row=fin), titles_from_data=True)
        chart2.set_categories(cat_ref)
        ws.add_chart(chart2, "A38")
    else:
        chart2 = BarChart(); chart2.type = "col"
        chart2.title = f"EVOLUCIÓN DE COMPRAS ANUAL {anio_base}"
        chart2.add_data(Reference(ws, min_col=col_c_base, min_row=fila_encabezado, max_row=fin), titles_from_data=True)
        chart2.set_categories(cat_ref)
        ws.add_chart(chart2, "A38")

    # 3) Ventas mensual (solo año base)
    chart3 = BarChart(); chart3.type = "col"
    chart3.title = f"EVOLUCIÓN DE VENTAS MENSUAL {anio_base}"
    chart3.add_data(Reference(ws, min_col=col_v_base, min_row=fila_encabezado, max_row=fin), titles_from_data=True)
    chart3.set_categories(cat_ref)
    ws.add_chart(chart3, "M19")

    # 4) Compras mensual (solo año base)
    chart4 = BarChart(); chart4.type = "col"
    chart4.title = f"EVOLUCIÓN DE COMPRAS MENSUAL {anio_base}"
    chart4.add_data(Reference(ws, min_col=col_c_base, min_row=fila_encabezado, max_row=fin), titles_from_data=True)
    chart4.set_categories(cat_ref)
    ws.add_chart(chart4, "M38")

    # 5) Ventas vs Compras mensual combinado (año base)
    chart5 = BarChart(); chart5.type = "col"
    chart5.title = f"EVOLUCIÓN DE VENTAS - COMPRAS MENSUAL - {anio_base}"
    chart5.add_data(Reference(ws, min_col=col_v_base, min_row=fila_encabezado, max_row=fin), titles_from_data=True)
    chart5.add_data(Reference(ws, min_col=col_c_base, min_row=fila_encabezado, max_row=fin), titles_from_data=True)
    chart5.set_categories(cat_ref)
    ws.add_chart(chart5, "A57")

    return ws


def construir_workbook(ruc, calc, branding, logo_bytes, anio_base, anio_anterior, datos_por_anio):
    wb = Workbook()
    _hoja_igv_renta(wb, calc, branding, logo_bytes)
    _hoja_evolucion(wb, ruc, branding, anio_base, anio_anterior, datos_por_anio)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
