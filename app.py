# -*- coding: utf-8 -*-
"""Interfaz Streamlit del Generador de Auditoría Semántica SEO."""
from __future__ import annotations
import hmac
import logging
import os
import datetime as dt

import streamlit as st
from dotenv import load_dotenv

import build_excel
import categorizar
import dataforseo

load_dotenv()

logger = logging.getLogger("auditoria_semantica")
logging.basicConfig(level=logging.INFO)

st.set_page_config(page_title="Generador de Auditoría Semántica SEO",
                   layout="wide")

# Compactar el layout: menos padding y menos aire entre widgets.
st.markdown("""
<style>
.block-container {padding-top: 1.6rem; padding-bottom: 0.5rem; max-width: 1250px;}
div[data-testid="stVerticalBlock"] {gap: 0.45rem;}
h1 {font-size: 1.55rem; margin-bottom: 0;}
div[data-testid="stWidgetLabel"] {margin-bottom: 0.1rem;}
</style>
""", unsafe_allow_html=True)


def _get_secret(name: str, default: str = "") -> str:
    """Lee una configuracion SOLO de variables de entorno o st.secrets.

    Nunca hay valores por defecto con datos reales en el código: si no está
    configurada, se devuelve `default` (normalmente cadena vacía).
    """
    val = os.getenv(name, "")
    if val:
        return val
    try:
        return str(st.secrets[name])
    except Exception:
        return default


DFS_LOGIN = _get_secret("DATAFORSEO_LOGIN")
DFS_PASSWORD = _get_secret("DATAFORSEO_PASSWORD")
APP_PASSWORD = _get_secret("APP_PASSWORD")


def _check_password() -> bool:
    """Gate de acceso simple con una contraseña compartida por el equipo.

    La contraseña se compara con hmac.compare_digest (evita timing attacks
    triviales) y nunca se guarda en el código: viene de APP_PASSWORD en
    variables de entorno o st.secrets. Si no está configurada, se avisa y se
    deja pasar (pensado solo para desarrollo local).
    """
    if st.session_state.get("_auth_ok", False):
        return True
    if not APP_PASSWORD:
        st.warning("Sin contraseña de acceso configurada (APP_PASSWORD). "
                   "Esto es aceptable en local, pero NO despliegues así una "
                   "instancia compartida con el equipo.", icon="⚠️")
        return True

    st.title("Generador de Auditoría Semántica SEO")
    with st.form("login"):
        pwd = st.text_input("Contraseña de acceso", type="password")
        entrar = st.form_submit_button("Entrar")
    if entrar:
        if hmac.compare_digest(pwd, APP_PASSWORD):
            st.session_state["_auth_ok"] = True
            st.rerun()
        else:
            st.error("Contraseña incorrecta.")
    return False


if not _check_password():
    st.stop()

PAISES = {"España (ES)": "es", "México (MX)": "mx", "Argentina (AR)": "ar",
          "Colombia (CO)": "co", "Estados Unidos (US)": "us",
          "Reino Unido (GB)": "gb", "Francia (FR)": "fr", "Alemania (DE)": "de",
          "Italia (IT)": "it", "Portugal (PT)": "pt"}

SIN_LIMITE = 100000

st.title("Generador de Auditoría Semántica SEO")
st.caption("1 cliente + 4 competidores · Excel de auditoría automático")

col_dom, col_opts = st.columns([2.1, 1], gap="large")

with col_dom:
    st.markdown("**Dominios**")
    etiquetas = ["Cliente", "Comp. 1", "Comp. 2", "Comp. 3", "Comp. 4"]
    entradas = []
    for i, etq in enumerate(etiquetas):
        c0, c1, c2 = st.columns([0.55, 1, 1.6])
        c0.markdown(f"<div style='padding-top:0.55rem;font-size:0.9rem'>{etq}</div>",
                    unsafe_allow_html=True)
        n = c1.text_input("Nombre", key=f"n{i}", placeholder="Nombre",
                          label_visibility="collapsed")
        u = c2.text_input("URL", key=f"u{i}", placeholder="https://…",
                          label_visibility="collapsed")
        entradas.append((n, u))

with col_opts:
    st.markdown("**Opciones**")
    pais = st.selectbox("País", list(PAISES.keys()), index=0)
    o1, o2 = st.columns([1, 1])
    limite = o1.number_input("Máx. keywords", 100, SIN_LIMITE, 2000, step=500,
                             help="Solo controla coste y velocidad (~0,12 $ por "
                                  "cada 1000 keywords y dominio en DataForSEO). "
                                  "La herramienta pagina automáticamente por "
                                  "encima de 1000.")
    sin_limite = o2.checkbox("Sin límite", value=False,
                             help="Trae todas las keywords del dominio.")
    ia_key = st.text_input("API key de IA (opcional)", type="password",
                           help="Anthropic, OpenAI o Gemini. Solo para categorizar; "
                                "sin key el Excel se genera sin categorías.")
    generar = st.button("Crear auditoría semántica", type="primary",
                        use_container_width=True)
    resultado = st.container()

if generar:
    cliente = (entradas[0][0].strip(), entradas[0][1].strip())
    competidores = [(n.strip(), u.strip()) for n, u in entradas[1:]]
    country = PAISES[pais]
    con_categorias = bool(ia_key.strip())
    limite_efectivo = SIN_LIMITE if sin_limite else int(limite)

    with resultado:
        if not cliente[0] or not cliente[1]:
            st.error("Falta el nombre o la URL del cliente.")
            st.stop()
        if not DFS_LOGIN or not DFS_PASSWORD:
            st.error("Faltan DATAFORSEO_LOGIN / DATAFORSEO_PASSWORD (.env o secrets).")
            st.stop()

        todos = [cliente] + competidores
        n_dom = sum(1 for etq, url in todos if etq and url)

        # Aviso de gasto (NO bloquea, solo informa): antes de lanzar las
        # peticiones se estima el coste maximo y se compara con el saldo que
        # de verdad queda en la cuenta de DataForSEO. Por decision explicita
        # del usuario, esto nunca detiene la ejecucion — solo avisa para que
        # se pueda decidir con informacion.
        coste = dataforseo.coste_estimado(limite_efectivo, n_dom)
        try:
            saldo = dataforseo.saldo_disponible(DFS_LOGIN, DFS_PASSWORD)
        except Exception:
            logger.exception("No se pudo consultar el saldo de DataForSEO")
            saldo = None

        if saldo is not None and coste > saldo:
            st.warning(
                f"Esta auditoría podría costar hasta ~{coste:.2f} $ y solo "
                f"quedan {saldo:.2f} $ de saldo en DataForSEO. Continuando "
                "de todas formas.")

        try:
            dominios = []
            prog = st.progress(0.0, text="Descargando keywords de DataForSEO…")
            hechos = 0
            for etq, url in todos:
                if not etq or not url:
                    continue
                dominios.append(dataforseo.fetch_dominio(
                    DFS_LOGIN, DFS_PASSWORD, etq, url, country=country,
                    limit=limite_efectivo))
                hechos += 1
                prog.progress(hechos / max(n_dom, 1), text=f"Descargado: {etq}")
            prog.empty()

            vacios = [d.etiqueta for d in dominios if not d.filas]
            if vacios:
                st.warning("Sin datos de DataForSEO para: " + ", ".join(vacios))

            if con_categorias:
                with st.spinner("Categorizando con IA…"):
                    ok, msg = categorizar.categorizar_dominios(dominios[:1], ia_key.strip())
                    if not ok:
                        logger.warning("Categorización con incidencias: %s", msg)
                        st.warning("La categorización con IA tuvo incidencias; "
                                  "el resto del Excel se generó igualmente.")

            with st.spinner("Construyendo el Excel…"):
                data = build_excel.construir_bytes(cliente, competidores, dominios,
                                                   con_categorias=con_categorias)

            total_kw = sum(len(d.filas) for d in dominios)
            st.success(f"Auditoría generada · {len(dominios)} dominios · {total_kw} keywords")
            fecha = dt.date.today().isoformat()
            st.download_button(
                "Descargar auditoría (.xlsx)", data=data,
                file_name=f"Auditoria_Semantica_{build_excel._safe(cliente[0])}_{fecha}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)

        except dataforseo.DataForSEOSaldoError as e:
            logger.warning("DataForSEO sin saldo: %s", e)
            st.error("La cuenta de DataForSEO no tiene saldo suficiente. "
                     "Recarga en app.dataforseo.com.")
        except dataforseo.DataForSEOError as e:
            logger.warning("Error de DataForSEO: %s", e)
            st.error(f"No se ha podido obtener datos de DataForSEO: {e}")
        except Exception:  # noqa: BLE001
            # No se muestra el traceback crudo a quien usa la app compartida
            # (podría exponer rutas internas, versiones de librerías, etc.);
            # queda registrado en el log del servidor para depurar.
            logger.exception("Error inesperado generando la auditoría")
            st.error("Ha ocurrido un error inesperado generando la auditoría. "
                     "Si persiste, contacta con el administrador de la herramienta.")
