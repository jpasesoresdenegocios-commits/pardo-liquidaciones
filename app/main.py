import re
from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse

from .supabase_client import sb
from . import igv_renta, branding as branding_mod
from .pdt_parser import obtener_evolucion_mes
from .excel_builder import construir_workbook

app = FastAPI(title="Pardo Liquidaciones")


@app.get("/")
def salud():
    return {"ok": True, "servicio": "pardo-liquidaciones"}


def _empresa_id(ruc):
    r = sb.table("empresas").select("id").eq("ruc", ruc).maybe_single().execute()
    if not r or not r.data:
        raise HTTPException(404, f"No existe la empresa con RUC {ruc}")
    return r.data["id"]


@app.get("/liquidacion/{ruc}/{periodo}")
def generar_liquidacion(ruc: str, periodo: str, anio_comparar: str | None = None):
    if not re.fullmatch(r"\d{11}", ruc):
        raise HTTPException(400, "RUC inválido (deben ser 11 dígitos)")
    if not re.fullmatch(r"\d{6}", periodo):
        raise HTTPException(400, "Período inválido (formato AAAAMM, ej. 202608)")

    empresa_id = _empresa_id(ruc)
    calc = igv_renta.calcular(ruc, empresa_id, periodo)
    br = branding_mod.obtener_branding(ruc)
    logo_bytes = branding_mod.descargar_logo(br["logo_url"])

    anio_base = periodo[:4]
    anios = [anio_base] + ([anio_comparar] if anio_comparar and anio_comparar != anio_base else [])
    datos_por_anio = {}
    for anio in anios:
        datos_por_anio[anio] = []
        for mes in range(1, 13):
            p = f"{anio}{mes:02d}"
            try:
                datos_por_anio[anio].append(obtener_evolucion_mes(empresa_id, p))
            except Exception:
                datos_por_anio[anio].append(None)

    anio_anterior = anio_comparar if (anio_comparar and anio_comparar != anio_base) else None
    buf = construir_workbook(ruc, calc, br, logo_bytes, anio_base, anio_anterior, datos_por_anio)

    nombre_archivo = f"Liquidacion_{br['nombre'][:30].replace(' ', '_')}_{periodo}.xlsx"
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{nombre_archivo}"'},
    )
