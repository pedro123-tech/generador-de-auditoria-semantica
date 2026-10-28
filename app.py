# -*- coding: utf-8 -*-
"""Interfaz Streamlit del Generador de Auditoría Semántica SEO."""
from __future__ import annotations
import hmac
import html
import logging
import os
import datetime as dt

import streamlit as st
from dotenv import load_dotenv

import build_excel
import categorizar
import dataforseo
import nombres

load_dotenv()

logger = logging.getLogger("auditoria_semantica")
logging.basicConfig(level=logging.INFO)

st.set_page_config(page_title="Auditoría Semántica SEO – Media Power",
                   page_icon="⚡", layout="wide")

# Estilo del panel "Equipo SEO – Media Power": los colores base vienen de
# .streamlit/config.toml; aqui solo la barra superior, las tarjetas y los
# detalles que el tema de Streamlit no cubre.
AMARILLO = "#FFD43B"
st.markdown(f"""
<style>
.block-container {{padding-top: 1.2rem; padding-bottom: 0.5rem; max-width: 1250px;}}
div[data-testid="stVerticalBlock"] {{gap: 0.5rem;}}
div[data-testid="stWidgetLabel"] {{margin-bottom: 0.1rem;}}
header[data-testid="stHeader"] {{background: transparent;}}
.mp-top {{display: flex; align-items: center; gap: 0.7rem; padding-bottom: 0.9rem;
          margin-bottom: 1.4rem; border-bottom: 1px solid #1C2232;}}
.mp-logo {{width: 34px; height: 34px; border-radius: 8px; background: {AMARILLO};
           display: flex; align-items: center; justify-content: center;}}
.mp-marca {{font-weight: 700; font-size: 1.02rem; color: #F4F5F7;}}
.amarillo {{color: {AMARILLO};}}
.mp-titulo {{font-size: 1.7rem; font-weight: 700; margin: 0; color: #F4F5F7;}}
.mp-sub {{color: #8B93A7; font-size: 0.92rem; margin: 0.15rem 0 1.1rem 0;}}
.mp-card-titulo {{font-weight: 700; font-size: 1.05rem; color: #F4F5F7; margin-bottom: 0.2rem;}}
.mp-etq {{padding-top: 0.55rem; font-size: 0.9rem; color: #C9CEDA;}}
.mp-nombre {{padding-top: 0.55rem; font-size: 0.9rem; color: {AMARILLO};
             white-space: nowrap; overflow: hidden; text-overflow: ellipsis;}}
.mp-nombre.vacio {{color: #4A5266;}}
div[data-testid="stVerticalBlockBorderWrapper"] {{border-color: #2A3142;}}
button[kind="primary"], button[data-testid="stBaseButton-primary"],
div[data-testid="stDownloadButton"] button {{
    background: {AMARILLO}; border-color: {AMARILLO}; color: #0B0F19; font-weight: 600;}}
button[kind="primary"] p, button[data-testid="stBaseButton-primary"] p,
div[data-testid="stDownloadButton"] button p {{color: #0B0F19; font-weight: 600;}}
</style>
""", unsafe_allow_html=True)

_RAYO = ('<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#0B0F19" '
         'stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round">'
         '<path d="M13 2 3 14h9l-1 8 10-12h-9l1-8z"/></svg>')


def _cabecera() -> None:
    """Barra superior y titulo con la identidad del panel del equipo SEO."""
    st.markdown(f"""
<div class="mp-top"><div class="mp-logo">{_RAYO}</div>
<span class="mp-marca">Equipo <span class="amarillo">SEO</span> – Media Power</span></div>
<div class="mp-titulo">Auditoría semántica – <span class="amarillo">Media Power</span></div>
<div class="mp-sub">1 cliente + 4 competidores · Excel de auditoría automático</div>
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

    _cabecera()
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


@st.cache_data(ttl=24 * 3600, show_spinner="Detectando nombre…")
def _nombre_de(url: str) -> str:
    return nombres.detectar_nombre(url)


_cabecera()

col_dom, col_opts = st.columns([2.1, 1], gap="large")

with col_dom, st.container(border=True):
    st.markdown("<div class='mp-card-titulo'>Dominios</div>", unsafe_allow_html=True)
    st.caption("Pega solo el dominio: el nombre de cada marca (y de sus hojas "
               "en el Excel) se detecta automáticamente.")
    etiquetas = ["Cliente", "Comp. 1", "Comp. 2", "Comp. 3", "Comp. 4"]
    entradas = []
    for i, etq in enumerate(etiquetas):
        c0, c1, c2 = st.columns([0.5, 1.7, 1])
        c0.markdown(f"<div class='mp-etq'>{etq}</div>", unsafe_allow_html=True)
        u = c1.text_input("URL", key=f"u{i}", placeholder="dominio.com",
                          label_visibility="collapsed").strip()
        n = _nombre_de(u) if "." in u else ""
        # El nombre viene de la web del dominio: se escapa antes de pintarlo
        # como HTML.
        c2.markdown(
            f"<div class='mp-nombre'>→ {html.escape(n)}</div>" if n
            else "<div class='mp-nombre vacio'>→ nombre</div>",
            unsafe_allow_html=True)
        entradas.append((n, u))

with col_opts, st.container(border=True):
    st.markdown("<div class='mp-card-titulo'>Opciones</div>", unsafe_allow_html=True)
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
    cliente = entradas[0]
    competidores = entradas[1:]
    country = PAISES[pais]
    con_categorias = bool(ia_key.strip())
    limite_efectivo = SIN_LIMITE if sin_limite else int(limite)

    with resultado:
        if not cliente[0] or not cliente[1]:
            st.error("Falta el dominio del cliente.")
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
                "Descargar auditoría (.xlsx)", data=data, type="primary",
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
