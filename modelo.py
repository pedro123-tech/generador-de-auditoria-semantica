# -*- coding: utf-8 -*-
"""Modelo de datos compartido por todo el generador."""
from __future__ import annotations
from dataclasses import dataclass, field
from typing import Optional

# Caracteres con los que Excel puede interpretar una celda de texto como el
# inicio de una formula (o, en un CSV, como el inicio de una formula al
# importarlo). Mitigacion estandar de "formula/CSV injection": si un texto
# que viene de fuera (API externa, IA, o el propio usuario) empieza por uno
# de estos caracteres, se antepone un apostrofe para forzar texto literal.
_PREFIJOS_PELIGROSOS = ("=", "+", "-", "@", "\t", "\r")


def sanear_texto(valor):
    """Neutraliza una posible inyeccion de formula antes de que el valor
    acabe en una celda o se interpole en una formula del Excel generado.

    Se aplica en las fronteras de entrada de datos no confiables: la
    respuesta de la API de DataForSEO, la respuesta del proveedor de IA, y los
    nombres/URLs que teclea el usuario en la interfaz.
    """
    if isinstance(valor, str) and valor.startswith(_PREFIJOS_PELIGROSOS):
        return "'" + valor
    return valor


@dataclass
class KeywordRow:
    """Una fila de keyword orgánica, ya normalizada (columnas del Excel)."""
    keyword: str
    volumen: Optional[int]
    kd: Optional[int]              # keyword difficulty
    posicion: Optional[int]       # mejor posición orgánica
    url: str                      # URL que posiciona
    sum_traffic: float            # tráfico estimado (para calcular %)
    marca: str                    # "Marca" / "No marca"
    intencion: str = ""           # vacío por defecto
    categoria: str = ""           # rellena la IA
    subcategoria: str = ""        # rellena la IA
    pct_trafico: float = 0.0      # calculado = sum_traffic / total * 100


@dataclass
class Dominio:
    """Un dominio con su etiqueta, url y sus filas de keywords."""
    etiqueta: str                 # nombre visible (p.ej. "Dupla Logistics")
    url: str
    filas: list[KeywordRow] = field(default_factory=list)
    # Nombres de hoja de Excel ya saneados y garantizados unicos dentro del
    # libro. Se rellenan una sola vez en build_excel.construir_workbook()
    # antes de crear ninguna hoja; el resto del motor los usa tal cual en
    # lugar de recalcularlos (evita duplicados/nombres invalidos).
    hoja: str = ""
    hoja_analisis: str = ""

    def calcular_pct_trafico(self) -> None:
        total = sum(f.sum_traffic for f in self.filas) or 1
        for f in self.filas:
            f.pct_trafico = round(f.sum_traffic / total * 100, 2)
