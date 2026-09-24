"""Logo/color/regimen por empresa -- ya vive en almacen_datos.empresas_conciliador
(mismo blob que usan todos los scripts de Python del despacho), no hay que
armar ninguna tabla nueva ni extraer color de la imagen.

'regimen' se cargo el 2026-09-24 desde la hoja BD del Excel real de
Diabetes (RUC->Regimen Tributario, convertida a JSON por el usuario) --
OJO: esa hoja BD es una plantilla COMPARTIDA entre archivos y puede quedar
desactualizada para alguna empresa puntual si cambio de regimen (caso real
encontrado: EVEADAM aparecia como RER en esa hoja pero su propio Excel de
julio 2026 ya dice RMT -- se corrigio a mano con la fuente mas fresca).
Si una empresa nueva no tiene 'regimen' cargado todavia, se debe agregar
aqui mismo (mismo patron que 'color'/'logoUrl') en vez de inventar una
tabla nueva."""
import httpx
from .supabase_client import sb

_DEFAULT_COLOR = "1E3A5F"  # azul oscuro del template si la empresa no tiene color propio


def obtener_branding(ruc):
    r = sb.table("almacen_datos").select("contenido").eq(
        "id_archivo", "empresas_conciliador").maybe_single().execute()
    empresas = (r.data or {}).get("contenido") or {}
    info = empresas.get(ruc) or {}
    color = (info.get("color") or _DEFAULT_COLOR).lstrip("#").upper()
    paleta = info.get("paleta") or None
    return {
        "nombre": info.get("nombre", ruc),
        "color": color,
        "paleta": paleta,
        "logo_url": info.get("logoUrl") or "",
        "regimen_tributario": info.get("regimen") or None,
    }


def descargar_logo(logo_url):
    if not logo_url:
        return None
    try:
        resp = httpx.get(logo_url, timeout=15)
        resp.raise_for_status()
        return resp.content
    except Exception:
        return None
