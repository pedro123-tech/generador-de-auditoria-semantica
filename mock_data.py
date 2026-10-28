# -*- coding: utf-8 -*-
"""Datos de ejemplo (demo) para probar el generador sin gastar saldo de DataForSEO.

Genera keywords sinteticas pero realistas del sector logistico para los 5
dominios, con la misma forma que devolveria DataForSEO.
"""
from __future__ import annotations
import random

from modelo import KeywordRow, Dominio

_BASE_KEYWORDS = [
    # (keyword, volumen, kd, intencion_placeholder)
    ("operador logistico", 1000, 22), ("logistica integral", 260, 24),
    ("almacenaje y distribucion", 210, 18), ("servicio de almacenaje", 70, 8),
    ("fulfillment ecommerce", 40, 15), ("logistica 3pl", 50, 19),
    ("operador logistico 3pl", 50, 16), ("ultima milla", 480, 20),
    ("transporte de mercancias", 1600, 27), ("logistica aduanera", 20, 5),
    ("picking y packing", 90, 12), ("cross docking", 320, 25),
    ("logistica inversa", 210, 21), ("almacen logistico madrid", 90, 17),
    ("distribucion capilar", 40, 9), ("transporte frigorifico", 720, 29),
    ("logistica para ecommerce", 110, 23), ("gestion de stock", 260, 14),
    ("paqueteria industrial", 30, 7), ("logistica farmaceutica", 170, 26),
]

_CATEGORIAS = {
    "operador logistico": ("Servicios", "Logística integral"),
    "logistica integral": ("Servicios", "Logística integral"),
    "almacenaje y distribucion": ("Servicios", "Almacenaje"),
    "servicio de almacenaje": ("Servicios", "Almacenaje"),
    "fulfillment ecommerce": ("Ecommerce", "Fulfillment"),
    "logistica 3pl": ("Servicios", "3PL"),
    "operador logistico 3pl": ("Servicios", "3PL"),
    "ultima milla": ("Servicios", "Última milla"),
    "transporte de mercancias": ("Transporte", "Mercancías"),
    "logistica aduanera": ("Servicios", "Aduanas"),
    "picking y packing": ("Servicios", "Almacenaje"),
    "cross docking": ("Servicios", "Cross docking"),
    "logistica inversa": ("Servicios", "Logística inversa"),
    "almacen logistico madrid": ("Servicios", "Almacenaje"),
    "distribucion capilar": ("Transporte", "Distribución"),
    "transporte frigorifico": ("Transporte", "Frigorífico"),
    "logistica para ecommerce": ("Ecommerce", "Fulfillment"),
    "gestion de stock": ("Servicios", "Gestión de stock"),
    "paqueteria industrial": ("Transporte", "Paquetería"),
    "logistica farmaceutica": ("Servicios", "Farmacéutica"),
}


def _dominio_mock(etiqueta: str, url: str, seed: int, con_categorias: bool) -> Dominio:
    rnd = random.Random(seed)
    marca_tokens = etiqueta.lower().split()
    filas: list[KeywordRow] = []

    # Cada dominio rankea para un subconjunto distinto (para que haya huecos).
    kws = _BASE_KEYWORDS[:]
    rnd.shuffle(kws)
    n = rnd.randint(11, len(kws))
    for kw, vol, kd in kws[:n]:
        pos = rnd.choices(
            [rnd.randint(1, 3), rnd.randint(4, 10), rnd.randint(11, 20),
             rnd.randint(21, 50), rnd.randint(51, 100)],
            weights=[3, 4, 3, 2, 1])[0]
        traffic = max(1, int(vol * max(0.01, (30 - pos) / 30) * rnd.uniform(0.3, 1.0)))
        cat, sub = _CATEGORIAS.get(kw, ("", "")) if con_categorias else ("", "")
        filas.append(KeywordRow(
            keyword=kw, volumen=vol, kd=kd, posicion=pos,
            url=f"{url.rstrip('/')}/{kw.replace(' ', '-')}/",
            sum_traffic=traffic, marca="No marca",
            categoria=cat, subcategoria=sub,
        ))

    # Añade algunas keywords de marca propias.
    for i in range(rnd.randint(2, 4)):
        brand_kw = " ".join(marca_tokens) + ("" if i == 0 else f" {['contacto','opiniones','telefono'][i%3]}")
        filas.append(KeywordRow(
            keyword=brand_kw, volumen=rnd.choice([110, 210, 320]), kd=rnd.randint(0, 5),
            posicion=rnd.randint(1, 3), url=url,
            sum_traffic=rnd.randint(50, 300), marca="Marca",
            categoria=("Home" if con_categorias else ""), subcategoria="",
        ))

    dom = Dominio(etiqueta=etiqueta, url=url, filas=filas)
    dom.calcular_pct_trafico()
    return dom


def dominios_demo(cliente: tuple[str, str], competidores: list[tuple[str, str]],
                  con_categorias: bool = True) -> list[Dominio]:
    """Devuelve los 5 Dominios con datos de ejemplo."""
    doms = [_dominio_mock(cliente[0], cliente[1], seed=1, con_categorias=con_categorias)]
    for i, (etq, url) in enumerate(competidores, start=2):
        doms.append(_dominio_mock(etq, url, seed=i * 7, con_categorias=con_categorias))
    return doms
