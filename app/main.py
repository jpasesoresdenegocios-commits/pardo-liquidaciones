import re
from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse

from .supabase_client import sb
from . import igv_renta, branding as branding_mod
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
def generar_liquidacion(ruc: str, periodo: str):
    """2026-09-23: por pedido explicito, solo genera la hoja IGV-RENTA
    (preliquidacion) -- la hoja de Evolucion queda pausada (no se llama a
    construir_workbook con datos_por_anio) hasta que se retome ese tema."""
    if not re.fullmatch(r"\d{11}", ruc):
        raise HTTPException(400, "RUC inválido (deben ser 11 dígitos)")
    if not re.fullmatch(r"\d{6}", periodo):
        raise HTTPException(400, "Período inválido (formato AAAAMM, ej. 202608)")

    empresa_id = _empresa_id(ruc)
    calc = igv_renta.calcular(ruc, empresa_id, periodo)
    br = branding_mod.obtener_branding(ruc)
    logo_bytes = branding_mod.descargar_logo(br["logo_url"])

    buf = construir_workbook(ruc, calc, br, logo_bytes)

    nombre_archivo = f"Liquidacion_{br['nombre'][:30].replace(' ', '_')}_{periodo}.xlsx"
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{nombre_archivo}"'},
    )
