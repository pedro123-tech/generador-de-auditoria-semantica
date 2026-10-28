# -*- coding: utf-8 -*-
"""Cliente de la API de DataForSEO (Labs -> Google -> Ranked Keywords)."""
from __future__ import annotations
import math
import re
import unicodedata

import requests

from modelo import KeywordRow, Dominio, sanear_texto

BASE = "https://api.dataforseo.com/v3"
RANKED_URL = f"{BASE}/dataforseo_labs/google/ranked_keywords/live"
USER_DATA_URL = f"{BASE}/appendix/user_data"
PAGE = 1000  # maximo de filas por peticion en Ranked Keywords.

# Precio de DataForSEO Labs en modo Live (USD): se cobra por peticion y por
# keyword devuelta. Se usa para estimar el coste antes de lanzar la auditoria.
COSTE_POR_PETICION = 0.012
COSTE_POR_KEYWORD = 0.00012

# Pais de la interfaz -> (location_code, language_code) de DataForSEO.
LOCALIZACIONES = {
    "es": (2724, "es"), "mx": (2484, "es"), "ar": (2032, "es"),
    "co": (2170, "es"), "us": (2840, "en"), "gb": (2826, "en"),
    "fr": (2250, "fr"), "de": (2276, "de"), "it": (2380, "it"),
    "pt": (2620, "pt"),
}

# Primeras palabras de un nombre de marca demasiado genericas para usarlas
# solas como termino de marca ("Grupo X" no debe marcar toda keyword con
# "grupo").
_GENERICAS = {"grupo", "group", "hotel", "hoteles", "clinica", "tienda",
              "casa", "the", "los", "las", "seguros", "abogados", "centro"}

# Segundos niveles de dominio que no son la marca (marca.co.uk, marca.com.mx).
_SLD_GENERICOS = {"co", "com", "org", "net", "gob", "gov", "edu", "ac"}


class DataForSEOError(RuntimeError):
    pass


class DataForSEOSaldoError(DataForSEOError):
    """Se lanza cuando la cuenta no tiene saldo suficiente."""


def coste_estimado(limit: int, n_dominios: int) -> float:
    """Coste maximo en USD de descargar `limit` keywords de `n_dominios`."""
    peticiones = math.ceil(limit / PAGE)
    return n_dominios * (peticiones * COSTE_POR_PETICION + limit * COSTE_POR_KEYWORD)


def saldo_disponible(login: str, password: str) -> float:
    """Devuelve el saldo de la cuenta en USD (no consume saldo)."""
    r = requests.get(USER_DATA_URL, auth=(login, password), timeout=60)
    datos = _comprobar(r)
    return float(datos["tasks"][0]["result"][0]["money"]["balance"])


def fetch_dominio(
    login: str,
    password: str,
    etiqueta: str,
    url: str,
    country: str = "es",
    limit: int = 1000,
) -> Dominio:
    """Descarga las keywords orgánicas de un dominio y devuelve un `Dominio`."""
    target = _target(url)
    location_code, language_code = LOCALIZACIONES[country.lower()]
    terminos = terminos_marca(etiqueta, url)

    # Paginacion por tráfico estimado descendente: si se corta por `limit`,
    # siempre se quedan las keywords que mas tráfico aportan.
    filas: list[KeywordRow] = []
    vistos: set[str] = set()
    offset = 0
    while len(filas) < limit:
        tarea = {
            "target": target, "location_code": location_code,
            "language_code": language_code, "item_types": ["organic"],
            "order_by": ["ranked_serp_element.serp_item.etv,desc"],
            "limit": min(PAGE, limit - len(filas)), "offset": offset,
        }
        r = requests.post(RANKED_URL, auth=(login, password), json=[tarea],
                          timeout=180)
        task = _comprobar(r)["tasks"][0]
        resultado = (task.get("result") or [{}])[0] or {}
        items = resultado.get("items") or []

        # Un mismo dominio puede posicionar varias URLs para una keyword: nos
        # quedamos con la primera (la de mas tráfico, por el orden pedido).
        for it in items:
            fila = _fila(it, terminos)
            if fila.keyword and fila.keyword not in vistos:
                vistos.add(fila.keyword)
                filas.append(fila)
        if len(items) < tarea["limit"]:
            break  # ultima pagina
        offset += len(items)

    filas = filas[:limit]
    dom = Dominio(etiqueta=etiqueta, url=url, filas=filas)
    dom.calcular_pct_trafico()
    return dom


def _comprobar(r: requests.Response) -> dict:
    """Valida la respuesta HTTP y los status_code de DataForSEO (global y de
    la tarea). Los errores HTTP (401, 403...) tambien traen un JSON con el
    motivo real, asi que se mira antes de recurrir a raise_for_status."""
    if r.status_code == 401:
        raise DataForSEOError("Credenciales de DataForSEO invalidas (401).")
    try:
        respuesta = r.json()
    except ValueError:
        r.raise_for_status()
        raise DataForSEOError(f"Respuesta no valida de DataForSEO ({r.status_code}).")
    codigos = [(respuesta.get("status_code"), respuesta.get("status_message", ""))]
    for t in respuesta.get("tasks") or []:
        codigos.append((t.get("status_code"), t.get("status_message", "")))
    for codigo, mensaje in codigos:
        if codigo == 20000:
            continue
        if codigo == 40104:
            raise DataForSEOError(
                "La cuenta de DataForSEO no esta verificada: completa la "
                "verificacion en https://app.dataforseo.com/")
        if any(p in mensaje.lower() for p in ("balance", "funds", "payment")):
            raise DataForSEOSaldoError(mensaje)
        raise DataForSEOError(f"{codigo}: {mensaje}")
    r.raise_for_status()
    return respuesta


def _fila(it: dict, terminos: list[str]) -> KeywordRow:
    kd = it.get("keyword_data") or {}
    info = kd.get("keyword_info") or {}
    props = kd.get("keyword_properties") or {}
    serp = (it.get("ranked_serp_element") or {}).get("serp_item") or {}
    keyword = kd.get("keyword", "")
    return KeywordRow(
        # keyword/url vienen de una API externa: se sanean por si contuvieran
        # una cadena que Excel pudiera interpretar como formula al abrir el
        # archivo (ver modelo.sanear_texto).
        keyword=sanear_texto(keyword),
        volumen=info.get("search_volume"),
        kd=props.get("keyword_difficulty"),
        posicion=serp.get("rank_group"),
        url=sanear_texto(serp.get("url") or ""),
        sum_traffic=serp.get("etv") or 0,
        marca="Marca" if es_marca(keyword, terminos) else "No marca",
        intencion="",  # vacío por defecto (search_intent_info disponible si se quiere)
    )


def _normalizar(texto: str) -> str:
    """Minusculas y sin tildes."""
    t = unicodedata.normalize("NFKD", texto.lower())
    return "".join(c for c in t if not unicodedata.combining(c))


def _compacto(texto: str) -> str:
    """Normalizado y sin nada que no sea letra o numero ("Dupla-Logistics" -> "duplalogistics")."""
    return re.sub(r"[^a-z0-9]", "", _normalizar(texto))


def terminos_marca(etiqueta: str, url: str) -> list[str]:
    """Terminos que identifican la marca de un dominio, para la regex de marca.

    - El nombre completo compactado ("Dupla Logistics" -> "duplalogistics").
    - La etiqueta del dominio ("duplalogistics.com" -> "duplalogistics").
    - La primera palabra del nombre si es distintiva (>= 4 letras y no
      generica), para captar busquedas cortas como "dupla".
    """
    terminos = {_compacto(etiqueta)}
    partes = _target(url).split(".")
    if len(partes) >= 3 and partes[-2] in _SLD_GENERICOS:
        terminos.add(_compacto(partes[-3]))
    elif len(partes) >= 2:
        terminos.add(_compacto(partes[-2]))
    palabras = _normalizar(etiqueta).split()
    if palabras:
        primera = _compacto(palabras[0])
        if len(primera) >= 4 and primera not in _GENERICAS:
            terminos.add(primera)
    return sorted(t for t in terminos if len(t) >= 3)


def es_marca(keyword: str, terminos: list[str]) -> bool:
    """True si la keyword contiene algun termino de marca.

    Los terminos compuestos se buscan en la keyword compactada (asi
    "dupla logistics" y "duplalogistics" coinciden); los de una sola palabra
    corta se buscan como palabra completa para no marcar falsos positivos.
    """
    compacta = _compacto(keyword)
    palabras = set(re.findall(r"[a-z0-9]+", _normalizar(keyword)))
    for t in terminos:
        if len(t) >= 6 and t in compacta:
            return True
        if t in palabras:
            return True
    return False


def _target(url: str) -> str:
    """Normaliza una URL a dominio raiz.

    DataForSEO exige el dominio sin "https://" ni "www." (con ellos lo
    interpretaria como una URL concreta en vez del dominio entero).
    """
    t = url.strip().lower()
    for pref in ("https://", "http://"):
        if t.startswith(pref):
            t = t[len(pref):]
    t = t.split("/")[0].split("?")[0]   # fuera ruta y query
    if t.startswith("www."):
        t = t[4:]
    return t
