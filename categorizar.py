# -*- coding: utf-8 -*-
"""Categorizacion semantica con IA (Anthropic / OpenAI / Gemini).

Se envia, por lotes (una llamada por dominio), la lista de keywords + su URL y se
pide categoria principal y subcategoria en JSON. Si no hay key, se omite.
"""
from __future__ import annotations
import json

from modelo import Dominio, sanear_texto

SYSTEM = (
    "Eres un consultor SEO. Clasificas keywords de un mismo dominio en una "
    "CATEGORIA principal (temática de producto/servicio, pocas y consistentes) "
    "y una SUBCATEGORIA mas concreta, fijandote en la keyword y su URL. "
    "Responde SOLO con JSON valido."
)


def detectar_proveedor(api_key: str) -> str:
    """Detecta el proveedor por el prefijo de la key."""
    k = api_key.strip()
    if k.startswith("sk-ant-"):
        return "anthropic"
    if k.startswith("AIza"):
        return "gemini"
    if k.startswith("sk-"):
        return "openai"
    return "desconocido"


def _prompt(dom: Dominio) -> str:
    items = [{"keyword": f.keyword, "url": f.url} for f in dom.filas]
    return (
        f"Dominio: {dom.etiqueta} ({dom.url}).\n"
        f"Clasifica estas {len(items)} keywords. Devuelve un JSON con la forma "
        '{"resultados":[{"keyword":"...","categoria":"...","subcategoria":"..."}]} '
        "usando exactamente las mismas keywords.\n\n"
        + json.dumps(items, ensure_ascii=False)
    )


def _parse_y_aplicar(dom: Dominio, texto: str) -> None:
    texto = texto.strip()
    if texto.startswith("```"):
        texto = texto.strip("`")
        texto = texto[texto.find("{"):texto.rfind("}") + 1]
    data = json.loads(texto)
    mapa = {r["keyword"]: r for r in data.get("resultados", [])}
    for f in dom.filas:
        r = mapa.get(f.keyword)
        if r:
            # La categoria la genera un LLM: se sanea igual que cualquier
            # texto externo antes de que acabe en una celda del Excel.
            f.categoria = sanear_texto(r.get("categoria", "") or "")
            f.subcategoria = sanear_texto(r.get("subcategoria", "") or "")


def _llamar(proveedor: str, api_key: str, prompt: str) -> str:
    if proveedor == "anthropic":
        import anthropic
        cli = anthropic.Anthropic(api_key=api_key)
        msg = cli.messages.create(
            model="claude-sonnet-5", max_tokens=8000,
            system=SYSTEM, messages=[{"role": "user", "content": prompt}])
        return msg.content[0].text
    if proveedor == "openai":
        from openai import OpenAI
        cli = OpenAI(api_key=api_key)
        resp = cli.chat.completions.create(
            model="gpt-4o-mini", response_format={"type": "json_object"},
            messages=[{"role": "system", "content": SYSTEM},
                      {"role": "user", "content": prompt}])
        return resp.choices[0].message.content
    if proveedor == "gemini":
        import google.generativeai as genai
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel("gemini-1.5-flash", system_instruction=SYSTEM)
        return model.generate_content(prompt).text
    raise ValueError(f"Proveedor no soportado: {proveedor}")


def categorizar_dominios(dominios: list[Dominio], api_key: str) -> tuple[bool, str]:
    """Rellena categoria/subcategoria in-place. Devuelve (ok, mensaje)."""
    proveedor = detectar_proveedor(api_key)
    if proveedor == "desconocido":
        return False, "No se reconoce el proveedor de la API key de IA."
    errores = []
    for dom in dominios:
        if not dom.filas:
            continue
        try:
            texto = _llamar(proveedor, api_key, _prompt(dom))
            _parse_y_aplicar(dom, texto)
        except Exception as e:  # noqa: BLE001
            errores.append(f"{dom.etiqueta}: {e}")
    if errores:
        return False, f"Categorizacion parcial ({proveedor}). Errores: " + "; ".join(errores)
    return True, f"Categorizacion completada con {proveedor}."
