# -*- coding: utf-8 -*-
"""Interfaz Streamlit del Generador de Auditoría Semántica SEO."""
from __future__ import annotations
import hmac
import logging
import os
import datetime as dt

import streamlit as st
from dotenv import load_dotenv

import ahrefs
import build_excel
import categorizar

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


AHREFS_KEY = _get_secret("AHREFS_API_KEY")
APP_PASSWORD = _get_secret("APP_PASSWORD")
# Tope de unidades de Ahrefs que se permite consumir en una sola ejecución,
# para que ningún uso individual agote la cuota mensual compartida por todo
# el equipo. Configurable vía entorno/secrets; con un valor conservador por
# defecto.
MAX_UNITS_PER_RUN = int(_get_secret("MAX_UNITS_PER_RUN", "150000"))


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
                             help="Solo controla coste y velocidad (~40 unidades de "
                                  "Ahrefs por keyword). La herramienta pagina "
                                  "automáticamente por encima de 250.")
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
        if not AHREFS_KEY:
            st.error("No hay AHREFS_API_KEY configurada (.env o secrets).")
            st.stop()

        todos = [cliente] + competidores
        n_dom = sum(1 for etq, url in todos if etq and url)

        # Aviso de gasto (NO bloquea, solo informa): antes de lanzar las
        # peticiones se estima el coste maximo y se compara con (a) el umbral
        # configurado y (b) las unidades que de verdad quedan este mes en la
        # cuenta de Ahrefs. Por decision explicita del usuario, esto nunca
        # detiene la ejecucion — solo avisa para que se pueda decidir con
        # informacion, ya que el limite de keywords lo controla el propio
        # usuario y es su cuota la que está en juego.
        coste_estimado = limite_efectivo * n_dom * ahrefs.UNIDADES_POR_KEYWORD
        if coste_estimado > MAX_UNITS_PER_RUN:
            st.warning(
                f"Esta ejecución podría consumir hasta ~{coste_estimado:,} "
                f"unidades de Ahrefs (umbral de aviso: {MAX_UNITS_PER_RUN:,}). "
                "Continuando de todas formas.")

        try:
            cuota = ahrefs.unidades_disponibles(AHREFS_KEY)
            restante = (cuota.get("units_limit_workspace", 0)
                       - cuota.get("units_usage_workspace", 0))
        except Exception:
            logger.exception("No se pudo consultar la cuota de Ahrefs")
            restante = None

        if restante is not None and coste_estimado > restante:
            st.warning(
                f"Esta auditoría podría necesitar hasta ~{coste_estimado:,} "
                f"unidades y solo quedan {restante:,} disponibles este mes en "
                "la cuenta de Ahrefs. Continuando de todas formas.")

        try:
            dominios = []
            prog = st.progress(0.0, text="Descargando keywords de Ahrefs…")
            hechos = 0
            for etq, url in todos:
                if not etq or not url:
                    continue
                dominios.append(ahrefs.fetch_dominio(
                    AHREFS_KEY, etq, url, country=country, limit=limite_efectivo))
                hechos += 1
                prog.progress(hechos / max(n_dom, 1), text=f"Descargado: {etq}")
            prog.empty()

            vacios = [d.etiqueta for d in dominios if not d.filas]
            if vacios:
                st.warning("Sin datos de Ahrefs para: " + ", ".join(vacios))

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

        except ahrefs.AhrefsUnitsError as e:
            logger.warning("Ahrefs sin unidades: %s", e)
            st.error("La cuenta de Ahrefs se ha quedado sin unidades de API este mes.")
        except ahrefs.AhrefsError as e:
            logger.warning("Error de Ahrefs: %s", e)
            st.error("No se ha podido conectar con Ahrefs. Revisa la API key o "
                     "inténtalo de nuevo en unos minutos.")
        except Exception:  # noqa: BLE001
            # No se muestra el traceback crudo a quien usa la app compartida
            # (podría exponer rutas internas, versiones de librerías, etc.);
            # queda registrado en el log del servidor para depurar.
            logger.exception("Error inesperado generando la auditoría")
            st.error("Ha ocurrido un error inesperado generando la auditoría. "
                     "Si persiste, contacta con el administrador de la herramienta.")
