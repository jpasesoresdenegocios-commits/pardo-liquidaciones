import re
from typing import Optional
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
    br = branding_mod.obtener_branding(ruc)
    calc = igv_renta.calcular(ruc, empresa_id, periodo, regimen=br["regimen_tributario"])
    logo_bytes = branding_mod.descargar_logo(br["logo_url"])

    buf = construir_workbook(ruc, calc, br, logo_bytes)

    nombre_archivo = f"Liquidacion_{br['nombre'][:30].replace(' ', '_')}_{periodo}.xlsx"
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{nombre_archivo}"'},
    )


@app.get("/calcular/{ruc}/{periodo}")
def calcular_json(ruc: str, periodo: str, regimen: Optional[str] = None):
    """Igual que /liquidacion pero sin el Excel -- devuelve el dict crudo de
    igv_renta.calcular() como JSON. Pensado para ser consumido por scripts
    (ej. estimar_tributos_igv_renta_pdt.py) que necesitan el monto sin
    descargar/parsear un xlsx por cada empresa/periodo.

    'regimen' es un query param OPCIONAL (?regimen=...) que, si se manda,
    PISA el regimen_tributario de branding.py para esta llamada puntual.
    Por que: branding.py (almacen_datos.empresas_conciliador) solo tiene el
    campo 'regimen' cargado a mano para ~17 de las 67 empresas (se cargo una
    sola vez el 2026-09-24); para el resto, igv_renta.calcular(regimen=None)
    asume RMT por default -- correcto para la mayoria, pero confirmado
    INCORRECTO para varias empresas sin tag cuyo propio historial de
    tributos_periodo muestra pagos confirmados bajo el codigo de RER (3111)
    o Regimen General (3031), no RMT (3121). estimar_tributos_igv_renta_pdt.py
    resuelve el regimen real por empresa usando ese historial (mas confiable
    que el tag manual) y lo manda aqui para no heredar el supuesto RMT por
    default cuando hay evidencia mejor. Sin este parametro, el endpoint se
    comporta exactamente igual que antes (usa branding.py)."""
    if not re.fullmatch(r"\d{11}", ruc):
        raise HTTPException(400, "RUC inválido (deben ser 11 dígitos)")
    if not re.fullmatch(r"\d{6}", periodo):
        raise HTTPException(400, "Período inválido (formato AAAAMM, ej. 202608)")

    empresa_id = _empresa_id(ruc)
    br = branding_mod.obtener_branding(ruc)
    regimen_usado = regimen or br["regimen_tributario"]
    return igv_renta.calcular(ruc, empresa_id, periodo, regimen=regimen_usado)
