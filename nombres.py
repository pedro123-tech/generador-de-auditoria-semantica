# -*- coding: utf-8 -*-
"""Deteccion del nombre de marca de un dominio a partir de su URL."""
from __future__ import annotations
import html
import ipaddress
import re
import socket
from urllib.parse import urljoin, urlparse

import requests

from dataforseo import compacto, dominio_raiz, etiqueta_dominio

_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
       "(KHTML, like Gecko) Chrome/130.0 Safari/537.36")
_MAX_BYTES = 300_000
_MAX_REDIRECCIONES = 3
# Separadores habituales en un <title> ("Inicio | Dupla Logistics").
_SEPARADORES = re.compile(r"\s+[|\-–—:·•»]\s+|\s*[|–—·•»]\s*")


def detectar_nombre(url: str) -> str:
    """Nombre de marca de un dominio ("duplalogistics.com" -> "Dupla Logistics").

    Primero intenta leer el nombre que declara la propia web (og:site_name,
    application-name o el <title>); si la web no responde o lo que declara
    no se parece al dominio, lo construye a partir del dominio.
    """
    dominio = dominio_raiz(url)
    if not dominio or "." not in dominio or dominio.replace(".", "").isdigit():
        return dominio  # no es un dominio de verdad (IP, "localhost"...)
    etiqueta = etiqueta_dominio(url)
    try:
        nombre = _nombre_desde_web(dominio, etiqueta)
    except Exception:  # noqa: BLE001 — cualquier fallo de red cae al respaldo
        nombre = ""
    return nombre or _nombre_desde_dominio(etiqueta)


def _nombre_desde_dominio(etiqueta: str) -> str:
    """"dupla-logistics" -> "Dupla Logistics"; "duplalogistics" -> "Duplalogistics"."""
    return " ".join(p.capitalize() for p in re.split(r"[-_]+", etiqueta) if p)


def _nombre_desde_web(dominio: str, etiqueta: str) -> str:
    pagina = _descargar(f"https://{dominio}/")
    candidatos = [
        _meta(pagina, "og:site_name"),
        _meta(pagina, "application-name"),
    ]
    titulo = re.search(r"<title[^>]*>(.*?)</title>", pagina, re.I | re.S)
    if titulo:
        candidatos += _SEPARADORES.split(_limpiar(titulo.group(1)))

    clave = compacto(etiqueta)
    for c in candidatos:
        c = _limpiar(c or "")
        if 2 <= len(c) <= 40 and _se_parece(compacto(c), clave):
            return c
    return ""


def _se_parece(nombre: str, clave: str) -> bool:
    """El nombre detectado tiene que corresponder al dominio, para no
    quedarse con un eslogan ("Envíos en 24h") o un "Inicio"."""
    if not nombre or not clave:
        return False
    return nombre in clave or clave in nombre or (len(nombre) >= 4 and clave.startswith(nombre[:4]))


def _meta(pagina: str, nombre: str) -> str:
    for patron in (
        rf'<meta[^>]+(?:property|name)=["\']{re.escape(nombre)}["\'][^>]*content=["\']([^"\']*)',
        rf'<meta[^>]+content=["\']([^"\']*)["\'][^>]*(?:property|name)=["\']{re.escape(nombre)}["\']',
    ):
        m = re.search(patron, pagina, re.I)
        if m:
            return m.group(1)
    return ""


def _limpiar(texto: str) -> str:
    return re.sub(r"\s+", " ", html.unescape(texto)).strip()


def _descargar(url: str) -> str:
    """GET con timeout, tope de tamaño y sin salir a redes internas.

    La URL la teclea quien usa la app compartida, asi que se comprueba que
    cada salto (incluidas las redirecciones) apunta a una IP publica, para
    que no se pueda usar el servidor para sondear su red interna (SSRF).
    """
    for _ in range(_MAX_REDIRECCIONES + 1):
        partes = urlparse(url)
        if partes.scheme not in ("http", "https") or not partes.hostname:
            return ""
        if not _es_publica(partes.hostname):
            return ""
        r = requests.get(url, headers={"User-Agent": _UA}, timeout=6,
                         allow_redirects=False, stream=True)
        if r.is_redirect:
            url = urljoin(url, r.headers.get("Location", ""))
            r.close()
            continue
        if r.status_code != 200:
            return ""
        contenido = r.raw.read(_MAX_BYTES, decode_content=True)
        r.close()
        # Sin charset explicito requests supone ISO-8859-1; casi toda la web
        # actual es UTF-8 y asi no se rompen las tildes del nombre.
        charset = r.encoding if "charset" in r.headers.get("Content-Type", "").lower() else "utf-8"
        return contenido.decode(charset, errors="replace")
    return ""


def _es_publica(host: str) -> bool:
    try:
        ips = {info[4][0] for info in socket.getaddrinfo(host, None)}
    except OSError:
        return False
    return bool(ips) and all(ipaddress.ip_address(ip).is_global for ip in ips)
