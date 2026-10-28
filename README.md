# Generador de Auditoría Semántica SEO

Herramienta web que, a partir de **5 dominios** (1 cliente + 4 competidores),
genera un **Excel de auditoría semántica SEO** con la misma estructura y estilo
que la plantilla de referencia, de forma automática.

- **Datos:** DataForSEO Labs API · Google `ranked_keywords/live` (Competitor Research).
- **Marca / No marca:** regex con el nombre del dominio y de la marca (`dataforseo.terminos_marca`).
- **Categorización:** IA opcional (Anthropic / OpenAI / Gemini, autodetectada por la key).
- **Excel:** `openpyxl` — hojas por dominio, análisis comparativo con gráficos,
  matriz global y hoja *Keyword Research* con `VLOOKUP` vivo y formato condicional.

## Instalación

```bash
pip install -r requirements.txt
cp .env.example .env      # y rellena DATAFORSEO_LOGIN / DATAFORSEO_PASSWORD
```

## Uso

```bash
streamlit run app.py
```

Abre <http://localhost:8501>. Pega el dominio del cliente y de hasta 4
competidores (el nombre de cada marca, que da nombre a sus hojas del Excel, se
detecta solo a partir de la web: `og:site_name` o `<title>`, y si no, del
propio dominio), elige país y (opcional) la API key de IA para las categorías. Pulsa **Crear auditoría semántica** y descarga el `.xlsx`.

- **Máx. keywords:** controla velocidad y coste (DataForSEO cobra 0,012 $ por
  petición + 0,00012 $ por keyword: ~0,13 $ por cada 1000 keywords y dominio).
  La herramienta pagina automáticamente por encima de 1000 filas.

## Estructura

| Archivo | Rol |
|---|---|
| `app.py` | Interfaz Streamlit + control de acceso y aviso de saldo |
| `dataforseo.py` | Cliente de DataForSEO (paginación, regex de marca, saneo de entrada) |
| `nombres.py` | Detección del nombre de marca a partir del dominio |
| `.streamlit/config.toml` | Tema visual (colores del panel Equipo SEO – Media Power) |
| `categorizar.py` | Categorización IA por lotes (JSON), 3 proveedores |
| `build_excel.py` | Motor openpyxl que reconstruye toda la plantilla |
| `modelo.py` | Modelo de datos + saneo anti-inyección de fórmulas |

## Seguridad — leer antes de compartir la app con el equipo

Este proyecto está pensado para desplegarse y ser usado por varias personas a
la vez. Ten en cuenta lo siguiente:

### Secretos

- Las claves (`DATAFORSEO_LOGIN`/`DATAFORSEO_PASSWORD`, key de IA, `APP_PASSWORD`) se leen **solo**
  de variables de entorno o de `st.secrets`. El código nunca tiene una clave
  ni un valor por defecto hardcodeado.
- En local: `.env` (a partir de `.env.example`). En despliegue (Streamlit
  Community Cloud u otro host compatible): `.streamlit/secrets.toml` (a
  partir de `.streamlit/secrets.toml.example`).
- **Ninguno de los dos archivos reales (`.env`, `secrets.toml`) se sube a
  git** — están en `.gitignore`. Solo se versionan los `.example`, que no
  contienen secretos reales.
- Si alguna vez pegas una API key real en un chat, un issue, un commit o
  cualquier sitio fuera de tu gestor de secretos: **considérala expuesta y
  rótala** en el panel del proveedor (DataForSEO / Anthropic / OpenAI / Google),
  aunque solo haya estado ahí un momento.

### Contraseña de acceso compartido

- Configura `APP_PASSWORD` antes de compartir el link con el equipo. Sin
  ella, cualquiera con la URL puede usar la herramienta y gastar el saldo
  de DataForSEO. La app avisa explícitamente si detecta que no está configurada.

### Saldo de DataForSEO compartido

- DataForSEO es prepago: el saldo de la cuenta es **compartido por todo el
  equipo** y cada auditoría lo descuenta.
- Antes de lanzar una ejecución, la app estima el coste máximo y lo compara
  con el saldo real de la cuenta (consultado en vivo, gratis). Si no llega,
  avisa, pero no bloquea la ejecución.

### Excel generado

- Todo texto que llega de fuentes externas (keywords y URLs de DataForSEO,
  categorías generadas por IA, nombres/URLs que teclea el usuario) se pasa
  por `modelo.sanear_texto()` antes de escribirse en una celda: si empieza
  por `=`, `+`, `-`, `@`, tabulador o retorno de carro, se neutraliza con un
  apostrofe inicial para que Excel nunca pueda interpretarlo como una
  fórmula ("formula/CSV injection").
- Los nombres de hoja se sanean y **deduplican** antes de crear ninguna hoja
  (`build_excel._nombre_hoja_unico`), para que dos dominios con nombres
  parecidos —o que colisionen tras quitar caracteres inválidos— no generen
  un archivo corrupto.

### Errores

- Si algo falla generando una auditoría, la persona que usa la app ve un
  mensaje genérico (sin traceback ni rutas internas); el detalle completo
  queda en el log del servidor (`logging`) para quien administre la
  herramienta.

## Notas

- Marca / No marca: una keyword es de marca si contiene el nombre del dominio
  ("duplalogistics"), el nombre completo de la marca compactado o su primera
  palabra si es distintiva ("dupla"; no se usan genéricas como "grupo").
- *Keyword Research* y las hojas de análisis usan fórmulas vivas
  (`COUNTIF`/`SUMIF`/`VLOOKUP`), no valores precalculados.

## Despliegue

Pensado para **Streamlit Community Cloud** o **Hugging Face Spaces** (no
Netlify: no ejecuta procesos Python persistentes).

1. Sube el repo (sin `.env` ni `secrets.toml`, ya excluidos por `.gitignore`).
2. En el panel de *Secrets* del servicio, pega el contenido de
   `.streamlit/secrets.toml.example` relleno con los valores reales,
   incluyendo `APP_PASSWORD`.
3. Comparte el link con el equipo junto con la contraseña, por un canal
   que no sea el propio repositorio (Slack/1Password/etc.).
