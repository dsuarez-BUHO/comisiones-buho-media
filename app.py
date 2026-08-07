"""
Comisiones — Buho Media
=======================
Interfaz web (Streamlit Community Cloud). Descarga las 4 fuentes desde una carpeta
privada de Google Drive (Service Account), calcula las comisiones del período y ofrece
los reportes en Excel para descarga directa. Sin terminal, sin subir archivos a mano.
"""

import hmac
import urllib.parse
from datetime import date

import streamlit as st

from src.comisiones import (
    CANONICAL_NAMES,
    FILE_PATTERNS,
    MAX_FILE_BYTES,
    MAX_FILE_MB,
    DataValidationError,
    clear_drive_cache,
    get_payment_date_range,
    load_data_sources_from_drive,
    match_prefix,
    run_pipeline,
    upload_file_to_drive,
)

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# Detalle breve de cada insumo (se muestra en popovers, a un clic — no invade la pantalla).
INSUMOS: list[tuple[str, str]] = [
    ("📋 Margen Tarifario",
     "**Catálogo maestro.** Asocia cada cliente con su **KAM** y define la "
     "**tasa de comisión** según su plan de ventas."),
    ("📦 Catálogo BWMS",
     "**Respaldo.** Recupera la asignación de **clientes nuevos** que aún no figuran "
     "en el Margen Tarifario."),
    ("💵 Concentrado de Facturas",
     "**Facturación.** Aporta los **montos facturados** para calcular la "
     "**base comisionable** sin IVA (÷1.16)."),
    ("🏦 Asiento Contable",
     "**Cobranza.** Aporta la **fecha real de cobro** para aplicar las "
     "**penalizaciones** por pago tardío (100% / 70% / 30% / 0%)."),
]

# Paleta de marca (para reutilizar en HTML controlado como el hero/stepper).
BUHO_NAVY = "#1B2A44"
BUHO_ORANGE = "#F5821F"

# Logotipo oficial de Buho (isologo con los búhos en las "oo"), inline y recoloreable
# vía `currentColor` (el color lo fija el contenedor .buho-hero-brand en el CSS).
BUHO_LOGO_SVG = (
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 223 45" fill="none">'
    '<g clip-path="url(#buhoclip)">'
    '<path d="M52.752 0.203125H22.0583C9.87565 0.203125 0 10.1192 0 22.3511C0 33.7792 '
    '9.67039 44.7621 21.457 44.392C28.6787 44.1636 35.7732 39.1767 36.5567 31.5705H37.2534L42.8851 '
    '44.4007L53.3505 44.392C65.14 44.765 74.8075 33.7821 74.8075 22.3511C74.8075 10.1192 64.9347 '
    '0.203125 52.752 0.203125ZM22.0149 32.0243C16.6435 32.0243 12.2925 27.6705 12.2925 22.3019C12.2925 '
    '16.9333 16.6464 12.5795 22.0149 12.5795C27.3835 12.5795 31.7374 16.9333 31.7374 22.3019C31.7374 '
    '27.6705 27.3835 32.0243 22.0149 32.0243ZM52.6104 32.0243C47.2389 32.0243 42.888 27.6705 42.888 '
    '22.3019C42.888 16.9333 47.2418 12.5795 52.6104 12.5795C57.979 12.5795 62.3328 16.9333 62.3328 '
    '22.3019C62.3328 27.6705 57.979 32.0243 52.6104 32.0243Z" fill="currentColor"></path>'
    '<path d="M207.195 12.8181C198.129 12.7776 191.297 19.5454 191.266 28.6C191.231 37.9813 197.851 '
    '44.778 207.045 44.8012C216.195 44.8214 222.983 38.0565 223 28.8949C223.017 19.69 216.302 12.8585 '
    '207.195 12.8181ZM205.35 36.1426C202.37 35.4228 200.514 32.8064 200.335 29.0828C200.089 23.9513 '
    '203.671 20.5948 208.516 21.4188C211.832 21.9825 213.885 24.8475 213.87 28.8891C213.85 34.0495 '
    '210.054 37.2817 205.353 36.1455L205.35 36.1426Z" fill="currentColor"></path>'
    '<path d="M168.73 0V15.5622C171.549 13.3911 174.607 12.6886 177.921 13.0413C183.306 13.6108 186.828 '
    '17.0106 187.47 22.3648C187.686 24.1659 187.698 25.9987 187.709 27.8201C187.741 32.6162 187.721 '
    '37.4153 187.721 42.2115C187.721 42.7261 187.721 43.2407 187.721 43.8275H178.571C178.571 43.3361 '
    '178.571 42.8301 178.571 42.3271C178.571 37.0019 178.6 31.6767 178.559 26.3514C178.531 22.703 '
    '175.897 20.6331 172.419 21.4281C170.155 21.9456 168.727 23.9028 168.716 26.6087C168.695 31.7894 '
    '168.71 36.9701 168.71 42.1507C168.71 42.6711 168.71 43.1915 168.71 43.8044H159.652V0H168.733H168.73Z" '
    'fill="currentColor"></path>'
    '<path d="M126.73 13.7993H135.75C135.75 14.3891 135.75 14.9008 135.75 15.4154C135.75 20.6423 135.724 '
    '25.8692 135.758 31.099C135.784 34.9267 138.86 37.0574 142.471 35.8056C144.428 35.1262 145.582 '
    '33.3078 145.588 30.7839C145.602 25.7016 145.594 20.6163 145.594 15.5339C145.594 15.0078 145.594 '
    '14.4845 145.594 13.8398H154.692V43.8223H146.276C146.152 42.9897 146.033 42.2034 145.891 '
    '41.2551C142.506 44.7474 138.487 45.21 134.215 44.0247C129.751 42.7874 126.733 38.6764 126.629 '
    '33.5795C126.496 27.1557 126.6 20.7261 126.605 14.3024C126.605 14.1694 126.669 14.0335 126.73 '
    '13.8022V13.7993Z" fill="currentColor"></path>'
    '<path d="M135.2 10.0985C137.056 7.03111 138.701 4.28177 140.396 1.56423C140.581 1.26935 141.063 '
    '0.991813 141.413 0.988922C144.437 0.942666 147.461 0.962903 150.852 0.962903C149.517 2.43153 '
    '148.389 3.68333 147.247 4.92068C145.883 6.40087 144.521 7.88684 143.122 9.33523C142.81 9.65903 '
    '142.361 10.0522 141.962 10.0667C139.817 10.1389 137.672 10.1013 135.198 10.1013L135.2 10.0985Z" '
    'fill="currentColor"></path>'
    '<path d="M122.735 24.9464C121.315 16.3255 113.937 11.3096 105.692 13.4692C103.871 13.9462 102.22 '
    '15.0679 100.41 15.9323V0H91.6448V43.7813H100.292C100.341 42.7781 100.384 41.8935 100.425 '
    '41.0811C101.708 41.9599 102.83 42.9718 104.143 43.5818C108.913 45.805 115.504 44.2612 119.04 '
    '40.1907C122.896 35.7501 123.654 30.5289 122.737 24.9464H122.735ZM104.663 35.8686C101.561 34.6544 '
    '100.457 32.0814 100.387 29.0083C100.324 26.1375 101.095 23.5241 103.871 22.0901C106.302 20.8325 '
    '108.812 20.8904 111.116 22.4573C113.359 23.9837 113.978 26.3052 113.952 28.8695C113.903 34.4347 '
    '109.508 37.7651 104.663 35.8686Z" fill="currentColor"></path>'
    '</g><defs><clipPath id="buhoclip"><rect width="223" height="44.8018" fill="white"></rect>'
    '</clipPath></defs></svg>'
)

# Color del logotipo en el hero (navy). Cambiar a "#FFFFFF" para blanco. Naranja de marca:
LOGO_COLOR = "#F5821F"

# <img> con data-URI: robusto ante la sanitización de Streamlit (que puede eliminar <svg>).
BUHO_LOGO_IMG = (
    '<img class="buho-logo" alt="Buho" src="data:image/svg+xml;utf8,'
    + urllib.parse.quote(BUHO_LOGO_SVG.replace("currentColor", LOGO_COLOR))
    + '">'
)

# CSS único (Nivel 1+2+3). Elementos compartidos usan variables nativas de Streamlit
# (var(--text-color), etc.); el hex de marca se reserva a hero/stepper/badge.
_STYLES = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

/* Tipografía de marca (sin tocar los íconos, que fijan su propia font-family) */
.stApp, .stApp p, .stApp label, .stApp h1, .stApp h2, .stApp h3, .stApp h4,
.stApp button, .stApp input, .stApp .stMarkdown {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
}

/* Chrome más limpio: header transparente, sin footer */
[data-testid="stHeader"] { background: transparent; }
footer { visibility: hidden; }

/* Ancho centrado — arregla la dispersión de layout=wide */
.block-container { max-width: 960px; padding-top: 1.6rem; padding-bottom: 3rem; }

/* ── Hero de marca ─────────────────────────────────────────────────────────── */
.buho-hero {
    background: linear-gradient(135deg, #16233A 0%, #22314F 55%, #2C3E60 100%);
    border-radius: 20px; padding: 28px 32px; margin: 0 0 24px;
    box-shadow: 0 14px 34px rgba(22,35,58,0.28);
}
.buho-hero-brand { margin-bottom: 14px; }
.buho-logo { height: 30px; width: auto; display: block; }
.buho-hero-title { color: #FFFFFF; font-size: 1.95rem; font-weight: 800;
                   letter-spacing: -0.02em; line-height: 1.12; }
.buho-hero-sub { color: #C3D0E4; font-size: 1rem; margin-top: 6px; line-height: 1.4; }

/* ── Eyebrow (etiquetas de sección) ────────────────────────────────────────── */
.buho-eyebrow { text-transform: uppercase; letter-spacing: 0.07em; font-size: 0.76rem;
                font-weight: 700; color: var(--text-color); opacity: 0.5;
                margin: 0.6rem 0 0.5rem; }

/* ── Stepper conectado ─────────────────────────────────────────────────────── */
.buho-stepper { display: flex; align-items: center; flex-wrap: wrap; row-gap: 8px;
                margin: 0.2rem 0 1.4rem; }
.buho-step { display: flex; align-items: center; gap: 9px; }
.buho-step .num { width: 26px; height: 26px; border-radius: 50%; background: #F5821F;
                  color: #fff; font-weight: 700; font-size: 0.82rem; display: flex;
                  align-items: center; justify-content: center; flex: 0 0 auto; }
.buho-step .lbl { font-weight: 600; color: var(--text-color); font-size: 0.92rem; }
.buho-line { width: 34px; height: 2px; background: #E3E8EF; margin: 0 12px; }

/* ── Badge "Datos disponibles" ─────────────────────────────────────────────── */
.buho-badge { display: inline-block; background: #FFF3E6; color: #8A4B0A;
              border: 1px solid #F6C99B; border-radius: 999px; padding: 7px 15px;
              font-size: 0.92rem; font-weight: 500; margin-bottom: 0.6rem; }

/* ── Botones ───────────────────────────────────────────────────────────────── */
.stButton > button, .stDownloadButton > button { border-radius: 9px; font-weight: 600; }
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {
    box-shadow: 0 6px 16px rgba(245,130,31,0.34); border: none; }

/* ── Tarjetas con profundidad ──────────────────────────────────────────────── */
[data-testid="stVerticalBlockBorderWrapper"] {
    border-radius: 16px; box-shadow: 0 6px 20px rgba(22,35,58,0.07);
    border-color: rgba(22,35,58,0.08); }

/* ── Métricas con acento (valor completo, sin recorte) ─────────────────────── */
[data-testid="stMetric"] { background: var(--secondary-background-color);
                           border-left: 3px solid #F5821F; border-radius: 12px;
                           padding: 14px 16px; overflow: visible; }
[data-testid="stMetricValue"] { font-size: 1.45rem; font-weight: 700;
                                white-space: nowrap; overflow: visible; }
[data-testid="stMetricValue"] > div { overflow: visible; }

/* ── Alertas y popovers ────────────────────────────────────────────────────── */
[data-testid="stAlert"] { border-radius: 12px; }
/* Popovers de insumos: misma altura aunque el texto ocupe 2 líneas */
[data-testid="stPopover"] button { border-radius: 9px; min-height: 54px;
    display: flex; align-items: center; justify-content: center;
    text-align: center; line-height: 1.2; }
</style>
"""


def _inject_styles() -> None:
    st.markdown(_STYLES, unsafe_allow_html=True)


def _render_hero(subtitle: str) -> None:
    """Header de marca (HTML controlado): banda navy + logotipo Buho + título + subtítulo."""
    st.markdown(
        f'<div class="buho-hero">'
        f'  <div class="buho-hero-brand">{BUHO_LOGO_IMG}</div>'
        f'  <div class="buho-hero-title">Comisiones</div>'
        f'  <div class="buho-hero-sub">{subtitle}</div>'
        f'</div>',
        unsafe_allow_html=True)


def _check_login() -> bool:
    """Gate de acceso simple (usuario/contraseña en st.secrets["auth"]).

    Nota: es una barrera básica de un solo usuario compartido, adecuada para uso
    interno. Para control por persona, usa el allowlist de Streamlit Community Cloud.
    """
    if st.session_state.get("auth_ok"):
        return True

    _render_hero("Acceso restringido. Ingresa tus credenciales para continuar.")
    with st.form("login"):
        usuario = st.text_input("Usuario")
        password = st.text_input("Contraseña", type="password")
        entrar = st.form_submit_button("Entrar", type="primary")

    if not entrar:
        return False

    try:
        esperado_u = st.secrets["auth"]["username"]
        esperado_p = st.secrets["auth"]["password"]
    except KeyError:
        st.error("⚠️ Login no configurado: falta la sección `[auth]` en *secrets*.")
        return False

    # Comparación en tiempo constante para no filtrar info por timing.
    if hmac.compare_digest(usuario, esperado_u) and hmac.compare_digest(password, esperado_p):
        st.session_state["auth_ok"] = True
        st.rerun()
    st.error("Usuario o contraseña incorrectos.")
    return False


def _reset_resultado() -> None:
    """Invalida el resultado previo al cambiar el período (evita descargar un corte viejo)."""
    st.session_state.pop("resultado", None)


def _render_header() -> None:
    _render_hero("Calcula comisiones por <b>fecha de cobro</b> (base sin IVA) y aplica "
                 "las penalizaciones por cobranza tardía.")
    st.markdown(
        '<div class="buho-stepper">'
        '<span class="buho-step"><span class="num">1</span>'
        '<span class="lbl">Elige el período</span></span>'
        '<span class="buho-line"></span>'
        '<span class="buho-step"><span class="num">2</span>'
        '<span class="lbl">Procesar Comisiones</span></span>'
        '<span class="buho-line"></span>'
        '<span class="buho-step"><span class="num">3</span>'
        '<span class="lbl">Descarga los reportes</span></span>'
        '</div>', unsafe_allow_html=True)

    # Documentación mínima: 4 popovers en una fila; el detalle está a un clic.
    st.markdown('<div class="buho-eyebrow">¿Qué insumos usa el sistema?</div>',
                unsafe_allow_html=True)
    for col, (titulo, detalle) in zip(st.columns(len(INSUMOS)), INSUMOS):
        with col.popover(titulo, use_container_width=True):
            st.markdown(detalle)


def _render_admin_uploader() -> None:
    """Sección de administración: carga/actualiza las fuentes en Google Drive desde la app."""
    with st.expander("🗂️ Cargar / actualizar fuentes en Google Drive (administración)"):
        st.caption(
            "Arrastra los Excel del período. Cada archivo se guarda con su **nombre "
            "canónico** según el insumo reconocido y **actualiza** el existente en Drive. "
            f"Tamaño máximo por archivo: **{MAX_FILE_MB} MB**."
        )
        archivos = st.file_uploader(
            "Archivos .xlsx", type=["xlsx"], accept_multiple_files=True, key="uploader")
        if not archivos:
            return

        # Plan de subida: cada archivo → nombre canónico del insumo que reconoce por prefijo.
        planes = []  # (archivo, nombre_destino_canónico)
        for archivo in archivos:
            if archivo.size > MAX_FILE_BYTES:
                st.write(f"⛔ `{archivo.name}` pesa {archivo.size / 1024 / 1024:.1f} MB "
                         f"(máx {MAX_FILE_MB} MB); se omite.")
                continue
            clave = match_prefix(archivo.name)
            if clave:
                destino = CANONICAL_NAMES[clave]
                planes.append((archivo, destino))
                st.write(f"✅ `{archivo.name}` → **{clave}** → se guardará como `{destino}`")
            else:
                st.write(f"⚠️ `{archivo.name}` no coincide con ningún prefijo conocido; se omite.")

        if not planes:
            st.info("Ningún archivo válido para subir. Prefijos esperados: "
                    + ", ".join(f"`{p}`" for p in FILE_PATTERNS.values()))
            return

        if not st.button("⬆️ Subir a Google Drive", type="primary", key="btn_upload"):
            return
        try:
            folder_id = st.secrets["drive"]["folder_id"]
            with st.spinner("Subiendo archivos a Google Drive…"):
                for archivo, destino in planes:
                    accion = upload_file_to_drive(folder_id, destino, archivo.getvalue())
                    st.success(f"`{destino}`: {accion}")
            clear_drive_cache()          # el próximo "Procesar" verá los archivos nuevos
            _reset_resultado()
        except DataValidationError as exc:
            st.warning(f"⚠️ {exc}")
        except KeyError:
            st.error("❌ Falta `drive.folder_id` en *secrets*.")
        except Exception as exc:
            st.error(f"❌ No se pudieron subir los archivos a Drive. Detalle: {exc}")


def _rango_disponible() -> tuple[date, date] | None:
    """Rango de fechas de pago disponible en Drive (silencioso). None si aún no está listo."""
    try:
        sources = load_data_sources_from_drive(st.secrets["drive"]["folder_id"])
        return get_payment_date_range(sources)
    except Exception:
        return None


def _sync_rango_bounds(rango_disp: tuple[date, date] | None) -> None:
    """Reinicia el selector si el rango disponible cambió (evita valores fuera de límites)."""
    if st.session_state.get("_rango_disp") != rango_disp:
        st.session_state["_rango_disp"] = rango_disp
        st.session_state.pop("rango", None)


def _render_controles(rango_disp: tuple[date, date] | None) -> None:
    """Tarjeta de generación: rango disponible + selector de período + botón de procesar."""
    with st.container(border=True):
        col_badge, col_refresh = st.columns([8, 2])
        if rango_disp:
            dmin, dmax = rango_disp
            col_badge.markdown(
                f'<div class="buho-badge">📅 Datos disponibles: pagos del '
                f'<b>{dmin:%d/%m/%Y}</b> al <b>{dmax:%d/%m/%Y}</b></div>',
                unsafe_allow_html=True)
            st.date_input(
                "Período de corte (fecha de pago)", value=(dmin, dmax),
                format="DD/MM/YYYY",
                key="rango", on_change=_reset_resultado)
        else:
            today = date.today()
            st.date_input(
                "Período de corte (fecha de pago)", value=(today.replace(day=1), today),
                format="DD/MM/YYYY", key="rango", on_change=_reset_resultado)
        if col_refresh.button("🔄 Refrescar datos", help="Limpia la caché y vuelve a leer los archivos desde Google Drive."):
            clear_drive_cache()
            _reset_resultado()
            st.rerun()

        if st.button("Procesar Comisiones", type="primary"):
            rango = st.session_state["rango"]
            if not isinstance(rango, (list, tuple)) or len(rango) != 2:
                st.warning("Selecciona un rango completo: fecha de inicio y fecha de fin.")
                st.stop()
            _procesar(rango[0], rango[1])


def _procesar(inicio: date, fin: date) -> None:
    """Ejecuta el pipeline con indicador de micro-pasos y guarda el resultado en session_state."""
    try:
        folder_id = st.secrets["drive"]["folder_id"]
    except KeyError:
        st.error("❌ Falta configuración en *secrets*: revisa `gcp_service_account` y "
                 "`drive.folder_id` en los ajustes de la app.")
        st.stop()

    try:
        with st.status("Procesando comisiones…", expanded=True) as status:
            status.write("🔗 Conectando a Google Drive…")
            sources = load_data_sources_from_drive(folder_id)
            status.write("⬇️ Insumos descargados.")
            # run_pipeline reporta sus fases reales vía el callback progress.
            resultado = run_pipeline(inicio, fin, sources, progress=status.write)
            status.update(label="✅ Comisiones procesadas", state="complete", expanded=False)
        st.session_state["resultado"] = resultado
    except DataValidationError as exc:
        st.error(f"⚠️ Problema con los datos: {exc}")
        st.stop()
    except Exception as exc:  # HttpError de Drive, credenciales, red, etc.
        st.error("❌ No se pudo conectar con Google Drive o procesar la información. "
                 "Verifica que la carpeta esté compartida con la cuenta de servicio. "
                 f"Detalle: {exc}")
        st.stop()


def _render_resultados() -> None:
    """Muestra métricas y botones de descarga a partir del resultado en session_state."""
    resultado = st.session_state.get("resultado")
    if resultado is None:
        st.info("Selecciona un período y presiona **Procesar Comisiones** para comenzar.")
        return

    st.success(f"Período procesado: {resultado.periodo}")

    with st.container(border=True):
        col_fac, col_base, col_com = st.columns(3)
        col_fac.metric("Total Facturas Procesadas", f"{resultado.total_facturas:,}")
        col_base.metric("Base Comisionable Global",
                        f"${resultado.base_comisionable_global:,.2f} MXN")
        col_com.metric("Total Comisiones a Pagar",
                       f"${resultado.total_comisiones:,.2f} MXN")

        global_buffer, global_name = resultado.resumen_global
        global_buffer.seek(0)
        st.download_button(
            "⬇️ Descargar Resumen Global",
            data=global_buffer, file_name=global_name, mime=XLSX_MIME, type="primary",
        )

    # ── Resumen por vendedor (vista previa formateada) ────────────────────────
    with st.expander("📋 Resumen por vendedor"):
        st.dataframe(
            resultado.resumen_vendedores, hide_index=True, use_container_width=True,
            column_config={
                "Clientes": st.column_config.NumberColumn(format="%d"),
                "Base Comisionable MXN": st.column_config.NumberColumn(
                    "Base Comisionable", format="dollar"),
                "Comisión MXN": st.column_config.NumberColumn("Comisión", format="dollar"),
                "Impacto Cobranza MXN": st.column_config.NumberColumn(
                    "Impacto Cobranza", format="dollar"),
                "Tasa Efectiva": st.column_config.NumberColumn(format="%.2f%%"),
            },
        )

    # ── Descarga individual por vendedor (buscador) ───────────────────────────
    st.subheader("Reporte individual por vendedor")
    vendedor = st.selectbox("Buscar KAM", resultado.vendedores, index=None,
                            placeholder="Escribe o selecciona un vendedor…")
    if vendedor:
        m = resultado.metricas_vendedor[vendedor]
        with st.container(border=True):
            # Micro-resumen: confirmación visual antes de descargar el reporte.
            c_base, c_com, c_imp, c_cli = st.columns(4)
            c_base.metric("Base Comisionable", f"${m['base']:,.2f} MXN")
            c_com.metric("Total Comisión", f"${m['comision']:,.2f} MXN")
            c_imp.metric("Impacto Cobranza", f"${m['impacto']:,.2f} MXN")
            c_cli.metric("Clientes · Tasa", f"{m['clientes']} · {m['tasa']:.2%}")

            vend_buffer, vend_name = resultado.reportes_vendedor[vendedor]
            vend_buffer.seek(0)
            st.download_button(
                f"⬇️ Descargar reporte de {vendedor}",
                data=vend_buffer, file_name=vend_name, mime=XLSX_MIME, key=f"dl_{vendedor}",
            )

    if resultado.sin_match:
        st.warning(
            "Empresas sin match en catálogo (no se les calculó comisión): "
            + ", ".join(sorted(resultado.sin_match))
        )


def main() -> None:
    st.set_page_config(
        page_title="Comisiones — Buho Media", page_icon="🦉", layout="wide")
    _inject_styles()
    if not _check_login():
        return
    _render_header()
    _render_admin_uploader()
    st.divider()

    rango_disp = _rango_disponible()
    _sync_rango_bounds(rango_disp)
    _render_controles(rango_disp)

    _render_resultados()


if __name__ == "__main__":
    main()
