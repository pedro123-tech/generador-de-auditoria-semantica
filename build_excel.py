# -*- coding: utf-8 -*-
"""Motor openpyxl: reconstruye la auditoria semantica, fiel al original.

Diseño (extraido celda a celda del Excel original):
- Cabeceras: relleno navy #0F1820, Calibri blanco negrita (14 titulos / 12 cols).
- AMARILLO #FCE940 (texto oscuro) en toda columna que refiere al CLIENTE
  dentro de un contexto comparativo.
- "NO POS": fuente roja C00000 negrita, sin relleno. Posicionado: relleno
  verde EAF1DD + fuente verde 00823B.
- Nº competidores: 3 -> naranja FFC000, 4 -> verde 00B050.
- Pestañas: cliente amarilla, competidores azules 8DB3E2.
- Autofiltro en las cabeceras de las hojas de datos.

Layout de datos (fiel al original, sin Intencion):
- CLIENTE (fila 7 cabecera, datos desde 8):
  A Marca | B Cat.principal | C Cat.secundaria | D Palabras Claves |
  E Volumen | F Posicion | G %Trafico | H URL
- COMPETIDOR: sin categorias; con columnas comparativas del cliente:
  A Marca | B Palabras Claves | C Volumen | D Posicion actual |
  E Posicion <cliente>* | F %Trafico | G %Trafico <cliente>* |
  H URL | I URL <cliente>*        (* = VLOOKUP vivo, cabecera amarilla)

Todos los numeros son formulas vivas (COUNTIF/SUMIF/COUNTIFS/VLOOKUP).
"""
from __future__ import annotations
import os
from collections import OrderedDict
from io import BytesIO

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import PieChart, BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.workbook.properties import CalcProperties

from modelo import Dominio, sanear_texto

# ------------------------------------------------------------- estilos -------
NAVY = "0F1820"
DARK = "1F262F"
WHITE = "FFFFFF"
YELLOW = "FCE940"
GREEN_FILL = "EAF1DD"
GREEN_FONT = "00823B"
RED_FONT = "C00000"
AMBER = "FFC000"
GREEN4 = "00B050"
DUP_FILL = "FF0000"
URLBLUE = "1155CC"
TAB_CLIENT = "FCE940"
TAB_COMP = "8DB3E2"

_ASSETS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")

_FILL_NAVY = PatternFill("solid", fgColor=NAVY)
_FILL_YELLOW = PatternFill("solid", fgColor=YELLOW)
_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
_CENTER_NW = Alignment(horizontal="center", vertical="center")
_THIN = Side(style="thin", color="BFBFBF")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

SEGMENTOS = ["TOP 3", "Posiciones 4-10", "Posiciones 11-20",
             "Posiciones 21-50", "Posiciones 51-100"]
_SEG_RANGO = [(1, 3), (4, 10), (11, 20), (21, 50), (51, 100)]


def _title(ws, coord, text, span=1, size=14):
    """Titulo de bloque: navy, blanco, negrita."""
    c = ws[coord]
    c.value = text
    c.font = Font(name="Calibri", size=size, bold=True, color=WHITE)
    c.fill = _FILL_NAVY
    c.alignment = _CENTER
    c.border = _BORDER
    if span > 1:
        ws.merge_cells(start_row=c.row, start_column=c.column,
                       end_row=c.row, end_column=c.column + span - 1)
        for i in range(span):
            cc = ws.cell(row=c.row, column=c.column + i)
            cc.fill = _FILL_NAVY
            cc.border = _BORDER


def _head(cell, text=None, size=12, yellow=False):
    """Cabecera de columna. yellow=True -> columna del cliente."""
    if text is not None:
        cell.value = text
    if yellow:
        cell.font = Font(name="Calibri", size=size, bold=True, color=NAVY)
        cell.fill = _FILL_YELLOW
    else:
        cell.font = Font(name="Calibri", size=size, bold=True, color=WHITE)
        cell.fill = _FILL_NAVY
    cell.alignment = _CENTER
    cell.border = _BORDER


def _label(cell, text):
    """Etiqueta de fila (TOP 3, Marca, ...): negrita 12 oscura."""
    cell.value = text
    cell.font = Font(name="Calibri", size=12, bold=True, color=DARK)
    cell.alignment = _CENTER_NW
    cell.border = _BORDER


def _cell(ws, r, c, value=None, fmt=None, bold=False, align=_CENTER_NW):
    cell = ws.cell(row=r, column=c, value=value)
    cell.border = _BORDER
    if fmt:
        cell.number_format = fmt
    if bold:
        cell.font = Font(bold=True)
    if align:
        cell.alignment = align
    return cell


# ---------------------------------------------------------------- helpers ----
def _cols(is_client: bool) -> dict:
    if is_client:
        return dict(marca="A", cat="B", keyword="C", volumen="D",
                    pos="E", pct="F", url="G")
    return dict(marca="A", keyword="B", volumen="C", pos="D",
                pos_cli="E", pct="F", pct_cli="G", url="H", url_cli="I")


_HOJAS_RESERVADAS = {"history"}  # nombre de hoja reservado por Excel


def _safe(nombre: str) -> str:
    """Sanitiza un nombre de hoja de Excel.

    Quita caracteres invalidos (incluido el apostrofe: aunque Excel lo
    permite en el interior de un nombre, complica el escapado al referenciar
    la hoja desde formulas, así que se elimina directamente), evita nombres
    vacios y la palabra reservada "History", y trunca a 31 caracteres (limite
    de Excel). No garantiza unicidad por si sola: usa `_nombre_hoja_unico`
    para eso.
    """
    nombre = (nombre or "").strip()
    for ch in "[]:*?/\\'!":
        nombre = nombre.replace(ch, "")
    nombre = nombre.strip()[:31].strip()
    if not nombre or nombre.lower() in _HOJAS_RESERVADAS:
        nombre = "Dominio"
    return nombre


def _nombre_hoja_unico(deseado: str, usados: set[str]) -> str:
    """Devuelve un nombre de hoja valido que no colisiona (case-insensitive,
    como hace Excel) con ningún nombre ya usado en `usados`. Añade " (2)",
    " (3)"... si hace falta. Registra el resultado en `usados`.
    """
    base = _safe(deseado)
    candidato = base
    i = 2
    while candidato.lower() in usados:
        sufijo = f" ({i})"
        candidato = base[:max(1, 31 - len(sufijo))].rstrip() + sufijo
        i += 1
    usados.add(candidato.lower())
    return candidato


def _qs(nombre_hoja: str) -> str:
    """Entrecomilla un nombre de hoja YA saneado/único para usarlo en una
    fórmula (p.ej. 'Dupla Logistics'!A1)."""
    return f"'{nombre_hoja}'"


def _q(sheet: str, cell: str) -> str:
    return f"'{sheet}'!{cell}"


def _idx(a, b):
    return column_index_from_string(b) - column_index_from_string(a) + 1


# Celdas de Inicio con nombre/url por dominio (cliente + 4 competidores)
_INICIO_NOMBRE = ["G11", "B27", "F27", "J27", "N27"]
_INICIO_URL = ["G13", "B29", "F29", "J29", "N29"]


# ------------------------------------------------------------ hoja Inicio ----
def _hoja_inicio(wb, cliente, competidores):
    ws = wb.active
    ws.title = "Inicio"
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = NAVY
    for col in "ABCDEFGHIJKLMNOPQRST":
        ws.column_dimensions[col].width = 11.5

    # Fondo completo navy (como el original)
    for row in ws.iter_rows(min_row=1, max_row=60, min_col=1, max_col=22):
        for c in row:
            c.fill = _FILL_NAVY

    def bloque(rng, text, size=12, url=False):
        first = rng.split(":")[0]
        ws.merge_cells(rng)
        c = ws[first]
        c.value = text
        c.alignment = _CENTER
        if url:
            c.font = Font(size=11, bold=True, color=URLBLUE, underline="single")
            fill = PatternFill("solid", fgColor="FFFFFF")
            r0, c0 = c.row, c.column
            ncols = ws[rng][0].__len__()
            for i in range(ncols):
                ws.cell(row=r0, column=c0 + i).fill = fill
        else:
            c.font = Font(size=size, bold=True, color=WHITE)
        return c

    def caja_blanca(rng):
        ws.merge_cells(rng)
        fill = PatternFill("solid", fgColor="FFFFFF")
        for fila in ws[rng]:
            for c in fila:
                c.fill = fill

    # Logo Mediapower arriba, centrado (blanco sobre fondo oscuro)
    logo = os.path.join(_ASSETS, "mediapower.png")
    if os.path.exists(logo):
        try:
            img = XLImage(logo)
            ratio = img.height / img.width if img.width else 0.25
            img.width = 280
            img.height = int(280 * ratio)
            ws.add_image(img, "G2")
        except Exception:
            pass  # sin Pillow o logo corrupto: se omite sin romper nada

    bloque("B8:Q9", "AUDITORÍA SEMÁNTICA", size=18)
    # cliente/competidores: texto tecleado por el usuario -> se sanea contra
    # inyeccion de formula antes de escribirlo en la celda.
    bloque("G11:L12", sanear_texto(cliente[0]), size=14)
    bloque("G13:L13", sanear_texto(cliente[1]), url=True)
    caja_blanca("G14:L21")  # hueco logo cliente (blanco por defecto)
    bloque("B26:Q26", "COMPETIDORES", size=14)
    anclas = [("B", "E"), ("F", "I"), ("J", "M"), ("N", "Q")]
    for i, (etq, url) in enumerate(competidores):
        c1, c2 = anclas[i]
        bloque(f"{c1}27:{c2}28", sanear_texto(etq), size=12)
        bloque(f"{c1}29:{c2}29", sanear_texto(url), url=True)
        caja_blanca(f"{c1}30:{c2}37")  # hueco logo (blanco por defecto)
    return ws


# --------------------------------------------------------- hojas de datos ----
def _cabecera_raw(ws, idx_dom, ncols):
    """Bloque superior C2/C3/C5 como el original (refs vivas a Inicio)."""
    _title(ws, "C2", "ANÁLISIS SEMÁNTICO", span=3)
    ws.merge_cells("C3:E3")
    c = ws["C3"]
    c.value = f"=Inicio!{_INICIO_NOMBRE[idx_dom]}"
    c.font = Font(size=12, bold=True)
    c.alignment = _CENTER
    c.border = _BORDER
    ws.merge_cells("C5:E5")
    c = ws["C5"]
    c.value = f"=Inicio!{_INICIO_URL[idx_dom]}"
    c.font = Font(size=11, bold=True, color=URLBLUE, underline="single")
    c.alignment = _CENTER
    c.border = _BORDER


def _hoja_raw_cliente(wb, dom: Dominio):
    ws = wb.create_sheet(dom.hoja)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = TAB_CLIENT
    headers = ["Marca / No marca", "Categoría principal", "Palabras Claves",
               "Volumen de búsqueda", "Posición", "%Trafico estimado",
               "URL posicionada"]
    widths = [20.4, 32.9, 41.7, 29.7, 17, 26.3, 95.6]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    _cabecera_raw(ws, 0, len(headers))
    for i, h in enumerate(headers, start=1):
        _head(ws.cell(row=7, column=i), h, size=14)
    r = 8
    for f in dom.filas:
        vals = [f.marca, f.categoria, f.keyword, f.volumen,
                f.posicion, f.pct_trafico, f.url]
        for i, v in enumerate(vals, start=1):
            c = _cell(ws, r, i, v)
            if i == 6 and isinstance(v, (int, float)):
                c.number_format = '0.00'
        r += 1
    ws.auto_filter.ref = f"A7:G{max(r - 1, 7)}"
    return ws


def _hoja_raw_comp(wb, dom: Dominio, cliente: Dominio, idx_dom: int):
    """Hoja de datos de competidor con columnas comparativas del cliente."""
    ws = wb.create_sheet(dom.hoja)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = TAB_COMP
    rawc = _qs(cliente.hoja)
    widths = [24, 40.9, 22, 18, 24, 22, 26, 95, 66]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    _cabecera_raw(ws, idx_dom, 9)

    _head(ws.cell(row=7, column=1), "Marca / No marca", size=14)
    _head(ws.cell(row=7, column=2), "Palabras Claves", size=14)
    _head(ws.cell(row=7, column=3), "Volumen de búsqueda", size=14)
    _head(ws.cell(row=7, column=4), "Posición actual", size=14)
    _head(ws.cell(row=7, column=5), '=CONCATENATE("Posicion ",Inicio!G11)',
          size=14, yellow=True)
    _head(ws.cell(row=7, column=6), "%Trafico estimado", size=14)
    _head(ws.cell(row=7, column=7), '=CONCATENATE("%Tráfico estimado ",Inicio!G11)',
          size=14, yellow=True)
    _head(ws.cell(row=7, column=8), "URL posicionada", size=14)
    _head(ws.cell(row=7, column=9), '=CONCATENATE("URL posicionada ",Inicio!G11)',
          size=14, yellow=True)

    cc = _cols(True)
    kw0 = cc["keyword"]
    r = 8
    for f in dom.filas:
        _cell(ws, r, 1, f.marca)
        _cell(ws, r, 2, f.keyword)
        _cell(ws, r, 3, f.volumen)
        _cell(ws, r, 4, f.posicion)
        # comparativas contra el cliente (VLOOKUP vivo sobre su hoja)
        _cell(ws, r, 5, f'=IFERROR(VLOOKUP($B{r},{rawc}!${kw0}:${cc["pos"]},'
                        f'{_idx(kw0, cc["pos"])},FALSE),"NO POS")')
        c = _cell(ws, r, 6, f.pct_trafico)
        if isinstance(f.pct_trafico, (int, float)):
            c.number_format = '0.00'
        _cell(ws, r, 7, f'=IFERROR(VLOOKUP($B{r},{rawc}!${kw0}:${cc["pct"]},'
                        f'{_idx(kw0, cc["pct"])},FALSE),"NO POS")', fmt='0.00')
        _cell(ws, r, 8, f.url)
        _cell(ws, r, 9, f'=IFERROR(VLOOKUP($B{r},{rawc}!${kw0}:${cc["url"]},'
                        f'{_idx(kw0, cc["url"])},FALSE),"NO POS")')
        r += 1
    fin = max(r - 1, 8)
    ws.auto_filter.ref = f"A7:I{fin}"

    # CF de las columnas comparativas (rojo NO POS / verde posicionado)
    _cf_nopos(ws, "E", 8, fin)
    _cf_nopos(ws, "G", 8, fin)
    _cf_nopos(ws, "I", 8, fin, urls=True)
    return ws


def _cf_nopos(ws, col, r0, r1, urls=False):
    rng = f"{col}{r0}:{col}{r1}"
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'ISNUMBER(SEARCH("NO POS",{col}{r0}))'],
        font=Font(color=RED_FONT, bold=True), stopIfTrue=True))
    if urls:
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f'AND({col}{r0}<>"",ISERROR(SEARCH("NO POS",{col}{r0})))'],
            fill=PatternFill("solid", fgColor=GREEN_FILL),
            font=Font(color=GREEN_FONT)))
    else:
        ws.conditional_formatting.add(rng, CellIsRule(
            operator="greaterThan", formula=["0"],
            fill=PatternFill("solid", fgColor=GREEN_FILL),
            font=Font(color=GREEN_FONT)))


# ---------------------------------------------- agregaciones (solo labels) ---
def _unicos_por_pct(dom: Dominio, campo: str):
    acc: "OrderedDict[str, float]" = OrderedDict()
    for f in dom.filas:
        v = getattr(f, campo, "")
        if v:
            acc[v] = acc.get(v, 0.0) + f.pct_trafico
    return [k for k, _ in sorted(acc.items(), key=lambda x: x[1], reverse=True)]


def _keywords_por_pct(dom: Dominio, cap=100):
    return [f.keyword for f in sorted(dom.filas, key=lambda x: x.pct_trafico,
                                      reverse=True)[:cap]]


# ------------------------------------------------ hoja analisis (cliente) ----
def _hoja_analisis_cliente(wb, dom: Dominio, con_cat: bool):
    ws = wb.create_sheet(dom.hoja_analisis)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = TAB_CLIENT
    raw = _qs(dom.hoja)
    c = _cols(True)
    for col in "ABDEFHIJLMNPQ":
        ws.column_dimensions[col].width = 17
    ws.column_dimensions["L"].width = 46
    ws.column_dimensions["P"].width = 32

    # Bloque segmentos (A-B)
    _title(ws, "A2", "KEYWORDS POR SEGMENTO DE POSICIONES", span=2)
    _head(ws["A3"], "SEGMENTO")
    _head(ws["B3"], "TOTAL DE KEYWORDS")
    for i, (s, (lo, hi)) in enumerate(zip(SEGMENTOS, _SEG_RANGO)):
        r = 4 + i
        _label(ws.cell(row=r, column=1), s)
        _cell(ws, r, 2, f'=COUNTIFS({raw}!{c["pos"]}:{c["pos"]},">="&{lo},'
                        f'{raw}!{c["pos"]}:{c["pos"]},"<="&{hi})')
    _label(ws.cell(row=9, column=1), "Total keywords")
    _cell(ws, 9, 2, "=SUM(B4:B8)", bold=True)

    # Bloque marca (D-F)
    _title(ws, "D2", "KEYWORDS POR MARCA", span=3)
    _head(ws["D3"], "MARCA / NO MARCA")
    _head(ws["E3"], "TOTAL DE KEYWORDS")
    _head(ws["F3"], "% TRÁFICO")
    for i, m in enumerate(["Marca", "No marca"]):
        r = 4 + i
        _label(ws.cell(row=r, column=4), m)
        _cell(ws, r, 5, f'=COUNTIF({raw}!{c["marca"]}:{c["marca"]},"{m}")')
        _cell(ws, r, 6, f'=SUMIF({raw}!{c["marca"]}:{c["marca"]},"{m}",'
                        f'{raw}!{c["pct"]}:{c["pct"]})', fmt='0.00')
    _label(ws.cell(row=6, column=4), "Total")
    _cell(ws, 6, 5, "=SUM(E4:E5)", bold=True)
    _cell(ws, 6, 6, "=SUM(F4:F5)", bold=True, fmt='0.00')

    # Bloque categoria principal (H-J) -- solo con IA
    n_cat = 0
    if con_cat:
        cats = _unicos_por_pct(dom, "categoria")
        n_cat = _bloque_campo(ws, "H", "KEYWORDS POR CATEGORÍA PRINCIPAL",
                              "CATEGORÍA", cats, raw, c["cat"], c["pct"])
    else:
        _title(ws, "H2", "CATEGORÍAS · no disponible (sin IA)", span=3)

    # Bloque URL (L-N)
    urls = _unicos_por_pct(dom, "url")
    _bloque_campo(ws, "L", "KEYWORDS POR URL", "URL", urls, raw, c["url"], c["pct"])

    # Bloque % trafico por keyword (P-Q)
    _title(ws, "P2", "% DE TRÁFICO POR KEYWORD", span=2)
    _head(ws["P3"], "KEYWORDS")
    _head(ws["Q3"], "% TRÁFICO")
    for i, kw in enumerate(_keywords_por_pct(dom)):
        r = 4 + i
        safe = str(kw).replace('"', '""')
        _cell(ws, r, 16, kw)
        _cell(ws, r, 17, f'=SUMIF({raw}!{c["keyword"]}:{c["keyword"]},"{safe}",'
                         f'{raw}!{c["pct"]}:{c["pct"]})', fmt='0.00')

    # Graficos
    _pie(ws, "Segmentos de posición", cat=("A", 4, 8), val=("B", 4, 8), anchor="A11")
    _pie(ws, "Marca / No marca", cat=("D", 4, 5), val=("E", 4, 5), anchor="D11")
    if con_cat and n_cat:
        _bar_simple(ws, "Categorías principales", cat=("H", 4, 3 + n_cat),
                    val=("I", 4, 3 + n_cat), anchor="H18")
    return ws


def _bloque_campo(ws, col0, titulo, etiqueta, labels, raw, campo_col, pct_col, cap=25):
    c0 = column_index_from_string(col0)
    _title(ws, f"{col0}2", titulo, span=3)
    _head(ws.cell(row=3, column=c0), etiqueta)
    _head(ws.cell(row=3, column=c0 + 1), "TOTAL DE KEYWORDS")
    _head(ws.cell(row=3, column=c0 + 2), "% TRÁFICO")
    for i, lab in enumerate(labels[:cap]):
        r = 4 + i
        safe = str(lab).replace('"', '""')
        _cell(ws, r, c0, lab)
        _cell(ws, r, c0 + 1, f'=COUNTIF({raw}!{campo_col}:{campo_col},"{safe}")')
        _cell(ws, r, c0 + 2, f'=SUMIF({raw}!{campo_col}:{campo_col},"{safe}",'
                             f'{raw}!{pct_col}:{pct_col})', fmt='0.00')
    return min(len(labels), cap)


# --------------------------------------------- hoja analisis (competidor) ----
def _hoja_analisis_comp(wb, dom: Dominio, cliente: Dominio, idx_dom: int):
    ws = wb.create_sheet(dom.hoja_analisis)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = TAB_COMP
    raw = _qs(dom.hoja)
    rawc = _qs(cliente.hoja)
    ansC = cliente.hoja_analisis
    c = _cols(False)
    cc = _cols(True)
    for col in "ABCEFGHIKLMOPRST":
        ws.column_dimensions[col].width = 17
    ws.column_dimensions["K"].width = 46
    ws.column_dimensions["O"].width = 32
    ws.column_dimensions["R"].width = 40

    ancla = _INICIO_NOMBRE[idx_dom]
    hdr_kw_comp = f'=CONCATENATE("KEYWORDS ",Inicio!{ancla})'
    hdr_kw_cli = '=CONCATENATE("KEYWORDS ",Inicio!G11)'

    # --- Segmentos comparativo (A-C) ---
    _title(ws, "A2", "KEYWORDS POR SEGMENTO DE POSICIONES", span=3)
    _head(ws["A3"], "SEGMENTO")
    _head(ws["B3"], hdr_kw_comp)
    _head(ws["C3"], hdr_kw_cli, yellow=True)
    for i, (s, (lo, hi)) in enumerate(zip(SEGMENTOS, _SEG_RANGO)):
        r = 4 + i
        _label(ws.cell(row=r, column=1), s)
        _cell(ws, r, 2, f'=COUNTIFS({raw}!{c["pos"]}:{c["pos"]},">="&{lo},'
                        f'{raw}!{c["pos"]}:{c["pos"]},"<="&{hi})')
        _cell(ws, r, 3, f'={_q(ansC, f"B{r}")}')
    _label(ws.cell(row=9, column=1), "Total keywords")
    _cell(ws, 9, 2, "=SUM(B4:B8)", bold=True)
    _cell(ws, 9, 3, f'={_q(ansC, "B9")}', bold=True)

    # --- Marca comparativo (E-I) ---
    _title(ws, "E2", "KEYWORDS POR MARCA", span=5)
    _head(ws["E3"], "MARCA / NO MARCA")
    _head(ws["F3"], hdr_kw_comp)
    _head(ws["G3"], hdr_kw_cli, yellow=True)
    _head(ws["H3"], f'=CONCATENATE("% TRÁFICO ",Inicio!{ancla})')
    _head(ws["I3"], '=CONCATENATE("% TRÁFICO ",Inicio!G11)', yellow=True)
    for i, m in enumerate(["Marca", "No marca"]):
        r = 4 + i
        _label(ws.cell(row=r, column=5), m)
        _cell(ws, r, 6, f'=COUNTIF({raw}!{c["marca"]}:{c["marca"]},"{m}")')
        _cell(ws, r, 7, f'={_q(ansC, f"E{r}")}')
        _cell(ws, r, 8, f'=SUMIF({raw}!{c["marca"]}:{c["marca"]},"{m}",'
                        f'{raw}!{c["pct"]}:{c["pct"]})', fmt='0.00')
        _cell(ws, r, 9, f'={_q(ansC, f"F{r}")}', fmt='0.00')
    _label(ws.cell(row=6, column=5), "Total")
    _cell(ws, 6, 6, "=SUM(F4:F5)", bold=True)
    _cell(ws, 6, 7, f'={_q(ansC, "E6")}', bold=True)
    _cell(ws, 6, 8, "=SUM(H4:H5)", bold=True, fmt='0.00')
    _cell(ws, 6, 9, f'={_q(ansC, "F6")}', bold=True, fmt='0.00')

    # --- URL (K-M) ---
    urls = _unicos_por_pct(dom, "url")
    _bloque_campo(ws, "K", "KEYWORDS POR URL", "URL", urls, raw, c["url"], c["pct"])

    # --- % trafico por keyword (O-P) ---
    _title(ws, "O2", "% DE TRÁFICO POR KEYWORD", span=2)
    _head(ws["O3"], "KEYWORDS")
    _head(ws["P3"], "% TRÁFICO")
    for i, kw in enumerate(_keywords_por_pct(dom)):
        r = 4 + i
        safe = str(kw).replace('"', '""')
        _cell(ws, r, 15, kw)
        _cell(ws, r, 16, f'=SUMIF({raw}!{c["keyword"]}:{c["keyword"]},"{safe}",'
                         f'{raw}!{c["pct"]}:{c["pct"]})', fmt='0.00')

    # --- Keywords compartidas (R-T) ---
    _title(ws, "R2", "KEYWORDS COMPARTIDAS", span=3)
    _head(ws["R3"], "KEYWORD")
    _head(ws["S3"], f'=CONCATENATE("POSICIÓN ",Inicio!{ancla})')
    _head(ws["T3"], '=CONCATENATE("POSICIÓN ",Inicio!G11)', yellow=True)
    kws_cli = {f.keyword for f in cliente.filas}
    compartidas = [f.keyword for f in sorted(dom.filas, key=lambda x: x.pct_trafico,
                                             reverse=True) if f.keyword in kws_cli]
    for i, kw in enumerate(compartidas[:60]):
        r = 4 + i
        safe = str(kw).replace('"', '""')
        _cell(ws, r, 18, kw)
        _cell(ws, r, 19, f'=IFERROR(VLOOKUP("{safe}",{raw}!{c["keyword"]}:{c["pos"]},'
                         f'{_idx(c["keyword"], c["pos"])},FALSE),"-")')
        _cell(ws, r, 20, f'=IFERROR(VLOOKUP("{safe}",{rawc}!{cc["keyword"]}:{cc["pos"]},'
                         f'{_idx(cc["keyword"], cc["pos"])},FALSE),"-")')
    n_comp = min(len(compartidas), 60)
    if n_comp:
        fin = 3 + n_comp
        ws.conditional_formatting.add(
            f"S4:S{fin}", FormulaRule(formula=["AND(ISNUMBER(S4),ISNUMBER(T4),S4<T4)"],
                                      fill=PatternFill("solid", fgColor="B7E1CD")))
        ws.conditional_formatting.add(
            f"T4:T{fin}", FormulaRule(formula=["AND(ISNUMBER(S4),ISNUMBER(T4),T4<S4)"],
                                      fill=PatternFill("solid", fgColor="B7E1CD")))

    # --- Graficos comparativos ---
    _bar_comp(ws, "Segmentos de posición", cat=("A", 4, 8), cols=(2, 3), anchor="A11")
    _bar_comp(ws, "Keywords por marca", cat=("E", 4, 5), cols=(6, 7), anchor="E8")
    _bar_comp(ws, "% Tráfico por marca", cat=("E", 4, 5), cols=(8, 9), anchor="E24")
    return ws


# ------------------------------------------------------------- graficos ------
def _pie(ws, titulo, cat, val, anchor):
    ch = PieChart(); ch.title = titulo; ch.height, ch.width = 7.5, 11
    data = Reference(ws, min_col=column_index_from_string(val[0]),
                     min_row=val[1], max_row=val[2])
    cats = Reference(ws, min_col=column_index_from_string(cat[0]),
                     min_row=cat[1], max_row=cat[2])
    ch.add_data(data, titles_from_data=False)
    ch.set_categories(cats)
    dl = DataLabelList(); dl.showPercent = True
    ch.dataLabels = dl
    ws.add_chart(ch, anchor)


def _bar_simple(ws, titulo, cat, val, anchor):
    ch = BarChart(); ch.type = "col"; ch.title = titulo; ch.height, ch.width = 8, 13
    data = Reference(ws, min_col=column_index_from_string(val[0]),
                     min_row=val[1], max_row=val[2])
    cats = Reference(ws, min_col=column_index_from_string(cat[0]),
                     min_row=cat[1], max_row=cat[2])
    ch.add_data(data, titles_from_data=False)
    ch.set_categories(cats)
    dl = DataLabelList(); dl.showVal = True; ch.dataLabels = dl
    ch.legend = None
    ws.add_chart(ch, anchor)


def _bar_comp(ws, titulo, cat, cols, anchor):
    ch = BarChart(); ch.type = "col"; ch.grouping = "clustered"
    ch.title = titulo; ch.height, ch.width = 7.5, 13
    data = Reference(ws, min_col=cols[0], max_col=cols[1], min_row=3, max_row=cat[2])
    cats = Reference(ws, min_col=column_index_from_string(cat[0]),
                     min_row=cat[1], max_row=cat[2])
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    dl = DataLabelList(); dl.showVal = True; ch.dataLabels = dl
    ch.legend.position = "b"
    ws.add_chart(ch, anchor)


def _bar_global(ws, titulo, cats, head_row, anchor):
    ch = BarChart(); ch.type = "col"; ch.grouping = "clustered"
    ch.title = titulo; ch.height, ch.width = 8, 17
    data = Reference(ws, min_col=2, max_col=6, min_row=head_row, max_row=cats[2])
    catref = Reference(ws, min_col=column_index_from_string(cats[0]),
                       min_row=cats[1], max_row=cats[2])
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(catref)
    dl = DataLabelList(); dl.showVal = True; ch.dataLabels = dl
    ch.legend.position = "b"
    ws.add_chart(ch, anchor)


# ---------------------------------------------------- hoja analisis global ---
def _hoja_global(wb, dominios: list[Dominio], con_cat: bool):
    ws = wb.create_sheet("Análisis global")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 46
    for col in "BCDEFGH":
        ws.column_dimensions[col].width = 20
    cli = dominios[0]
    comps = dominios[1:]
    ansC = cli.hoja_analisis
    ansX = [d.hoja_analisis for d in comps]

    def _fila_dom_hdr(row, prefijo):
        # columna B = cliente (amarilla), C-F competidores
        for i, ref in enumerate(_INICIO_NOMBRE):
            cell = ws.cell(row=row, column=2 + i)
            _head(cell, f'=CONCATENATE("{prefijo} ",Inicio!{ref})',
                  yellow=(i == 0))

    # Tabla 1 (filas 1-6)
    _title(ws, "A1", "COMPARATIVA DE TOTAL DE KEYWORDS DE MARCA / NO MARCA", span=6)
    _head(ws["A3"], "MARCA / NO MARCA")
    _fila_dom_hdr(3, "Keywords")
    _label(ws.cell(row=4, column=1), "Marca")
    _label(ws.cell(row=5, column=1), "No marca")
    _label(ws.cell(row=6, column=1), "Total Keywords")
    _cell(ws, 4, 2, f'={_q(ansC, "E4")}'); _cell(ws, 5, 2, f'={_q(ansC, "E5")}')
    _cell(ws, 6, 2, f'={_q(ansC, "E6")}', bold=True)
    for j, an in enumerate(ansX):
        col = 3 + j
        _cell(ws, 4, col, f'={_q(an, "F4")}')
        _cell(ws, 5, col, f'={_q(an, "F5")}')
        _cell(ws, 6, col, f'={_q(an, "F6")}', bold=True)

    # Tabla 2 (filas 8-13)
    _title(ws, "A8", "COMPARATIVA % DE TRÁFICO DE MARCA / NO MARCA", span=6)
    _head(ws["A10"], "MARCA / NO MARCA")
    _fila_dom_hdr(10, "% Tráfico")
    _label(ws.cell(row=11, column=1), "Marca")
    _label(ws.cell(row=12, column=1), "No marca")
    _label(ws.cell(row=13, column=1), "Total Keywords")
    _cell(ws, 11, 2, f'={_q(ansC, "F4")}', fmt='0.00')
    _cell(ws, 12, 2, f'={_q(ansC, "F5")}', fmt='0.00')
    _cell(ws, 13, 2, f'={_q(ansC, "F6")}', fmt='0.00', bold=True)
    for j, an in enumerate(ansX):
        col = 3 + j
        _cell(ws, 11, col, f'={_q(an, "H4")}', fmt='0.00')
        _cell(ws, 12, col, f'={_q(an, "H5")}', fmt='0.00')
        _cell(ws, 13, col, f'={_q(an, "H6")}', fmt='0.00', bold=True)

    # Tabla 3 (filas 15-23)
    _title(ws, "A15", "COMPARATIVA DE TOTAL DE KEYWORDS POR SEGMENTO DE POSICIONES", span=6)
    _head(ws["A17"], "SEGMENTO")
    _fila_dom_hdr(17, "Keywords")
    for i, s in enumerate(SEGMENTOS):
        r = 18 + i
        _label(ws.cell(row=r, column=1), s)
        _cell(ws, r, 2, f'={_q(ansC, f"B{4+i}")}')
        for j, an in enumerate(ansX):
            _cell(ws, r, 3 + j, f'={_q(an, f"B{4+i}")}')
    _label(ws.cell(row=23, column=1), "Total keywords")
    _cell(ws, 23, 2, f'={_q(ansC, "B9")}', bold=True)
    for j, an in enumerate(ansX):
        _cell(ws, 23, 3 + j, f'={_q(an, "B9")}', bold=True)

    # Matriz: posiciones y rango
    universo = _universo_con_meta(dominios)
    hdr = 37
    data0 = hdr + 1
    fin = data0 + max(len(universo), 1) - 1

    # Tabla 4 (filas 25-33): unicas por segmento, sobre la matriz
    _title(ws, "A25", "COMPARATIVA DE TOTAL DE KEYWORDS ÚNICAS POR SEGMENTO DE POSICIONES", span=6)
    _head(ws["A27"], "SEGMENTO")
    _fila_dom_hdr(27, "Keywords")
    lims = ["<4", "<11", "<21", "<51", "<101"]
    for i, s in enumerate(SEGMENTOS):
        r = 28 + i
        _label(ws.cell(row=r, column=1), s)
        for j in range(5):
            mcol = get_column_letter(3 + j)   # C-G de la matriz
            base = (f'COUNTIFS({mcol}${data0}:{mcol}${fin},"<>NO POS",'
                    f'{mcol}${data0}:{mcol}${fin},"{lims[i]}")')
            if i == 0:
                f = f"={base}"
            else:
                prev = "+".join(f"{get_column_letter(2+j)}{28+k}" for k in range(i))
                f = f"={base}-({prev})"
            _cell(ws, r, 2 + j, f)
    _label(ws.cell(row=33, column=1), "Total keywords")
    for j in range(5):
        cl = get_column_letter(2 + j)
        _cell(ws, 33, 2 + j, f"=SUM({cl}28:{cl}32)", bold=True)

    # Graficos
    _bar_global(ws, "Keywords Marca / No marca", cats=("A", 4, 5), head_row=3, anchor="J1")
    _bar_global(ws, "% Tráfico Marca / No marca", cats=("A", 11, 12), head_row=10, anchor="J16")
    _bar_global(ws, "Keywords por segmento de posiciones", cats=("A", 18, 22), head_row=17, anchor="J31")

    # Cabecera de la matriz
    _title(ws, f"A{hdr-2}", "LISTADO Y COMPARACIÓN DE POSICIONES DE KEYWORDS ÚNICAS", span=8)
    _head(ws.cell(row=hdr, column=1), "Palabra clave")
    _head(ws.cell(row=hdr, column=2), "Volumen de búsqueda")
    for i, ref in enumerate(_INICIO_NOMBRE):
        _head(ws.cell(row=hdr, column=3 + i),
              f'=CONCATENATE("Posición ",Inicio!{ref})', yellow=(i == 0))
    _head(ws.cell(row=hdr, column=8), "Nº COMPETIDORES POSICIONADOS")

    raws = [(_qs(d.hoja), _cols(k == 0)) for k, d in enumerate(dominios)]
    for j, (kw, vol) in enumerate(universo):
        r = data0 + j
        safe = str(kw).replace('"', '""')
        _cell(ws, r, 1, kw)
        _cell(ws, r, 2, vol)
        for k, (rw, cm) in enumerate(raws):
            _cell(ws, r, 3 + k,
                  f'=IFERROR(VLOOKUP("{safe}",{rw}!{cm["keyword"]}:{cm["pos"]},'
                  f'{_idx(cm["keyword"], cm["pos"])},FALSE),"NO POS")')
        _cell(ws, r, 8, f'=COUNTIF(D{r}:G{r},"<>"&"NO POS")')
    if universo:
        for col in "CDEFG":
            _cf_nopos(ws, col, data0, fin)
        rng = f"H{data0}:H{fin}"
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal",
                                       formula=["3"], fill=PatternFill("solid", fgColor=AMBER)))
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal",
                                       formula=["4"], fill=PatternFill("solid", fgColor=GREEN4)))
    return (hdr, data0, fin)


def _universo_con_meta(dominios):
    traf = {}
    vol = {}
    for d in dominios:
        for f in d.filas:
            traf[f.keyword] = traf.get(f.keyword, 0) + f.sum_traffic
            if (f.volumen or 0) > (vol.get(f.keyword) or 0):
                vol[f.keyword] = f.volumen
    orden = sorted(traf.items(), key=lambda x: x[1], reverse=True)
    return [(k, vol.get(k)) for k, _ in orden]


# --------------------------------------------------- hoja keyword research ---
def _hoja_keyword_research(wb, dominios, con_cat, matriz):
    ws = wb.create_sheet("Keyword Research")
    ws.sheet_view.showGridLines = False
    anchos = [54, 29.9, 27.1, 32.4, 29.4, 35.3, 33.7, 33, 32.7, 39.3]
    for i, w in enumerate(anchos, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    _title(ws, "A1", "KEYWORD RESEARCH", span=10)
    _head(ws["A3"], "Palabra clave", size=14)
    _head(ws["B3"], "Volumen de búsqueda", size=14)
    _head(ws["C3"], "Dificultad (KD)", size=14)
    _head(ws["D3"], "Categoría principal", size=14)
    for i, ref in enumerate(_INICIO_NOMBRE):
        _head(ws.cell(row=3, column=5 + i),
              f'=CONCATENATE("Posicion ",Inicio!{ref})', size=14, yellow=(i == 0))
    _head(ws.cell(row=3, column=10), "Nº competidores posicionados", size=14)

    opps = _oportunidades(dominios)
    rng = "'Análisis global'!$A:$G"
    n_fill = len(opps)
    total = 3 + n_fill + 40  # buffer de plantilla para pegar mas keywords
    for i in range(total - 3):
        r = 4 + i
        if i < n_fill:
            kw, meta = opps[i]
            _cell(ws, r, 1, kw)
            _cell(ws, r, 2, meta.get("volumen"))
            _cell(ws, r, 3, meta.get("kd"))
            _cell(ws, r, 4, meta.get("categoria", ""))
        else:
            for cidx in range(1, 5):
                _cell(ws, r, cidx, None)
        for k in range(5):
            _cell(ws, r, 5 + k,
                  f'=IF($A{r}="","",IFERROR(VLOOKUP($A{r},{rng},{3 + k},FALSE),"NO POS"))')
        _cell(ws, r, 10, f'=IF($A{r}="","",COUNTIF(F{r}:I{r},"<>"&"NO POS"))')

    ult = total
    for col in "EFGHI":
        _cf_nopos(ws, col, 4, ult)
    lrng = f"J4:J{ult}"
    ws.conditional_formatting.add(lrng, CellIsRule(operator="equal", formula=["3"],
                                   fill=PatternFill("solid", fgColor=AMBER)))
    ws.conditional_formatting.add(lrng, CellIsRule(operator="equal", formula=["4"],
                                   fill=PatternFill("solid", fgColor=GREEN4)))
    ws.conditional_formatting.add(f"A4:A{ult}", FormulaRule(
        formula=[f"COUNTIF($A$4:$A${ult},A4)>1"],
        fill=PatternFill("solid", fgColor=DUP_FILL)))
    return ws


def _oportunidades(dominios, cap=80):
    cli = dominios[0]
    comps = dominios[1:]
    pos_cli = {f.keyword: f.posicion for f in cli.filas}
    branded = set()
    meta = {}
    for d in dominios:
        for f in d.filas:
            if f.marca == "Marca":
                branded.add(f.keyword)
            m = meta.setdefault(f.keyword, {"volumen": f.volumen, "kd": f.kd,
                                            "categoria": f.categoria})
            if (f.volumen or 0) > (m["volumen"] or 0):
                m["volumen"] = f.volumen
            if not m["categoria"] and f.categoria:
                m["categoria"] = f.categoria
    comp_count = {}
    for d in comps:
        for f in d.filas:
            if f.posicion is not None:
                comp_count[f.keyword] = comp_count.get(f.keyword, 0) + 1
    out = []
    for kw, n in comp_count.items():
        if kw in branded or n < 2:
            continue
        pc = pos_cli.get(kw)
        if pc is not None and pc <= 10:
            continue
        out.append((kw, n, meta.get(kw, {})))
    out.sort(key=lambda x: (x[1], x[2].get("volumen") or 0), reverse=True)
    return [(kw, m) for kw, n, m in out[:cap]]


# --------------------------------------------------------------- publico -----
def construir_workbook(cliente, competidores, dominios, con_categorias=True):
    wb = Workbook()
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    _hoja_inicio(wb, cliente, competidores)

    # Nombres de hoja finales: saneados y sin colisiones entre si ni con las
    # hojas fijas del libro. Se calculan ANTES de crear ninguna hoja de
    # dominio para que todas las formulas cruzadas (VLOOKUP, referencias
    # entre 'Analisis X' y su hoja de datos) usen siempre el mismo nombre
    # resuelto. Sin esto, dos dominios con nombres parecidos (o que colisionen
    # tras sanear caracteres invalidos) romperian el libro (nombres de hoja
    # duplicados = Excel lo marca como corrupto al abrir).
    usados = {"inicio", "análisis global", "keyword research"}
    for dom in dominios:
        dom.hoja = _nombre_hoja_unico(dom.etiqueta, usados)
        dom.hoja_analisis = _nombre_hoja_unico(f"Análisis {dom.etiqueta}", usados)

    # Orden del original: cliente, analisis cliente, comp1, analisis comp1, ...
    cli = dominios[0]
    _hoja_raw_cliente(wb, cli)
    _hoja_analisis_cliente(wb, cli, con_categorias)
    for k, dom in enumerate(dominios[1:]):
        _hoja_raw_comp(wb, dom, cli, idx_dom=k + 1)
        _hoja_analisis_comp(wb, dom, cli, idx_dom=k + 1)
    matriz = _hoja_global(wb, dominios, con_categorias)
    _hoja_keyword_research(wb, dominios, con_categorias, matriz)
    return wb


def construir_bytes(*args, **kwargs) -> bytes:
    wb = construir_workbook(*args, **kwargs)
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()
