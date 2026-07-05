# -*- coding: utf-8 -*-
"""Cliente de la API de Ahrefs v3 (Site Explorer -> Organic Keywords)."""
from __future__ import annotations
import datetime as _dt
import json
import requests

from modelo import KeywordRow, Dominio, sanear_texto

BASE_URL = "https://api.ahrefs.com/v3/site-explorer/organic-keywords"
LIMITS_URL = "https://api.ahrefs.com/v3/subscription-info/limits-and-usage"
PAGE = 250  # Ahrefs limita a 250 filas por peticion en este plan.

# Unidades de API que cobra Ahrefs por keyword con las columnas que pedimos
# (volume, keyword_difficulty, sum_traffic -> 10 unidades cada una = ~40/fila).
# Se usa para estimar coste y frenar ejecuciones que agotarian la cuota
# mensual compartida antes de lanzar las peticiones.
UNIDADES_POR_KEYWORD = 40

# Columnas que pedimos a Ahrefs.
SELECT = ",".join([
    "keyword", "volume", "keyword_difficulty", "best_position",
    "best_position_url", "sum_traffic", "is_branded",
    "is_informational", "is_commercial", "is_transactional",
    "is_navigational", "is_local",
])


class AhrefsError(RuntimeError):
    pass


class AhrefsUnitsError(AhrefsError):
    """Se lanza cuando la cuenta se ha quedado sin unidades de API."""


def _headers(api_key: str) -> dict:
    return {"Authorization": f"Bearer {api_key}", "Accept": "application/json"}


def unidades_disponibles(api_key: str) -> dict:
    """Devuelve el estado de cuota de la cuenta (no consume unidades)."""
    r = requests.get(LIMITS_URL, headers=_headers(api_key), timeout=60)
    r.raise_for_status()
    return r.json().get("limits_and_usage", {})


def ultima_fecha_con_datos(api_key: str, target: str, country: str = "es",
                           mode: str = "subdomains", max_atras_dias: int = 210,
                           paso: int = 10) -> str | None:
    """Busca hacia atras la fecha mas reciente con datos para el dominio.

    Usa un `select` gratuito (keyword,best_position = 0 unidades) para no gastar
    cuota. Ahrefs no indexa el mes en curso, asi que hay que retroceder.
    """
    hoy = _dt.date.today()
    for d in range(0, max_atras_dias + 1, paso):
        fecha = (hoy - _dt.timedelta(days=d)).isoformat()
        params = {"target": target, "country": country.lower(), "mode": mode,
                  "select": "keyword,best_position", "date": fecha,
                  "limit": 1, "output": "json"}
        r = requests.get(BASE_URL, headers=_headers(api_key), params=params, timeout=60)
        if r.status_code == 403 and "units" in r.text.lower():
            raise AhrefsUnitsError(r.json().get("error", r.text))
        if r.status_code == 200 and r.json().get("keywords"):
            return fecha
    return None


def fetch_dominio(
    api_key: str,
    etiqueta: str,
    url: str,
    country: str = "es",
    limit: int = 1000,
    date: str | None = None,
    mode: str = "subdomains",
) -> Dominio:
    """Descarga las keywords orgánicas de un dominio y devuelve un `Dominio`."""
    target = _target(url)
    if date is None:
        date = ultima_fecha_con_datos(api_key, target, country=country, mode=mode)
        if date is None:
            # Dominio sin datos: devolvemos vacio (el resto de la auditoria sigue).
            return Dominio(etiqueta=etiqueta, url=url, filas=[])

    # Paginacion por tráfico descendente para superar el tope de 250 filas/peticion,
    # obteniendo siempre primero las keywords de mas tráfico.
    filas: list[KeywordRow] = []
    vistos: set[str] = set()
    boundary = None
    while len(filas) < limit:
        params = {
            "target": target, "country": country.lower(), "mode": mode,
            "select": SELECT, "date": date, "limit": PAGE,
            "order_by": "sum_traffic:desc", "output": "json",
        }
        if boundary is not None:
            params["where"] = json.dumps(
                {"field": "sum_traffic", "is": ["lte", boundary]})
        r = requests.get(BASE_URL, headers=_headers(api_key), params=params, timeout=120)
        if r.status_code == 403 and "units" in r.text.lower():
            raise AhrefsUnitsError(r.json().get("error", r.text))
        if r.status_code == 401:
            raise AhrefsError("API key de Ahrefs invalida o caducada (401).")
        r.raise_for_status()

        kws = r.json().get("keywords", [])
        nuevos = [k for k in kws if k.get("keyword") and k["keyword"] not in vistos]
        for k in nuevos:
            vistos.add(k["keyword"])
            filas.append(_fila(k))
        if len(kws) < PAGE or not nuevos:
            break  # ultima pagina o sin avances (grupo de empates)
        boundary = kws[-1].get("sum_traffic") or 0

    filas = filas[:limit]
    dom = Dominio(etiqueta=etiqueta, url=url, filas=filas)
    dom.calcular_pct_trafico()
    return dom


def _fila(k: dict) -> KeywordRow:
    return KeywordRow(
        # keyword/url vienen de una API externa: se sanean por si contuvieran
        # una cadena que Excel pudiera interpretar como formula al abrir el
        # archivo (ver modelo.sanear_texto).
        keyword=sanear_texto(k.get("keyword", "")),
        volumen=k.get("volume"),
        kd=k.get("keyword_difficulty"),
        posicion=k.get("best_position"),
        url=sanear_texto(k.get("best_position_url") or ""),
        sum_traffic=k.get("sum_traffic") or 0,
        marca="Marca" if k.get("is_branded") else "No marca",
        intencion="",  # vacío por defecto (los flags is_* estan disponibles si se quieren)
    )


def _target(url: str) -> str:
    """Normaliza una URL a dominio raiz.

    Importante: se quita el "www." porque con mode=subdomains el dominio raiz
    cubre www y todos los subdominios, mientras que "www.dominio.com" puede
    devolver 0 resultados si el sitio rankea en el dominio raiz (caso real:
    www.poppacool.com -> 0 filas, poppacool.com -> datos).
    """
    t = url.strip().lower()
    for pref in ("https://", "http://"):
        if t.startswith(pref):
            t = t[len(pref):]
    t = t.split("/")[0].split("?")[0]   # fuera ruta y query
    if t.startswith("www."):
        t = t[4:]
    return t
