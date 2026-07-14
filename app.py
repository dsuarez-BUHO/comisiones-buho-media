"""
Sistema Central de Comisiones — Buho Media
==========================================
Interfaz web (Streamlit Community Cloud). Descarga las 4 fuentes desde una carpeta
privada de Google Drive (Service Account), calcula las comisiones del período y ofrece
los reportes en Excel para descarga directa. Sin terminal, sin subir archivos a mano.
"""

import hmac
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


def _check_login() -> bool:
    """Gate de acceso simple (usuario/contraseña en st.secrets["auth"]).

    Nota: es una barrera básica de un solo usuario compartido, adecuada para uso
    interno. Para control por persona, usa el allowlist de Streamlit Community Cloud.
    """
    if st.session_state.get("auth_ok"):
        return True

    st.title("🦉 Comisiones — Buho Media")
    st.caption("Acceso restringido. Ingresa tus credenciales para continuar.")
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
    st.title("🦉 Comisiones — Buho Media")
    st.caption(
        "Calcula comisiones por **fecha de cobro** (base sin IVA) y aplica las "
        "penalizaciones por cobranza tardía."
    )
    st.markdown(
        "**Flujo:**  ①  Elige el período  →  ②  Procesar Comisiones  →  ③  Descarga los reportes"
    )
    # Documentación mínima: 4 popovers en una sola fila; el detalle está a un clic.
    st.write("ℹ️ ¿Qué insumos usa el sistema?")
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
    """Selector de período (acotado al rango disponible) + botón de procesar."""
    if rango_disp:
        dmin, dmax = rango_disp
        st.caption(f"📅 **Datos disponibles:** pagos del {dmin:%d/%m/%Y} al {dmax:%d/%m/%Y}.")
        st.date_input(
            "Período de corte (fecha de pago)", value=(dmin, dmax),
            min_value=dmin, max_value=dmax, format="DD/MM/YYYY",
            key="rango", on_change=_reset_resultado)
    else:
        today = date.today()
        st.date_input(
            "Período de corte (fecha de pago)", value=(today.replace(day=1), today),
            format="DD/MM/YYYY", key="rango", on_change=_reset_resultado)

    if st.button("Procesar Comisiones", type="primary"):
        rango = st.session_state["rango"]
        if not isinstance(rango, (list, tuple)) or len(rango) != 2:
            st.warning("Selecciona un rango completo: fecha de inicio y fecha de fin.")
            st.stop()
        _procesar(rango[0], rango[1])


def _procesar(inicio: date, fin: date) -> None:
    """Ejecuta el pipeline y guarda el resultado en session_state, con errores amigables."""
    try:
        with st.spinner("Descargando y procesando información…"):
            folder_id = st.secrets["drive"]["folder_id"]
            sources = load_data_sources_from_drive(folder_id)
            resultado = run_pipeline(inicio, fin, sources)
        st.session_state["resultado"] = resultado
    except DataValidationError as exc:
        st.error(f"⚠️ Problema con los datos: {exc}")
        st.stop()
    except KeyError:
        st.error(
            "❌ Falta configuración en *secrets*: revisa `gcp_service_account` y "
            "`drive.folder_id` en los ajustes de la app."
        )
        st.stop()
    except Exception as exc:  # HttpError de Drive, credenciales, red, etc.
        st.error(
            "❌ No se pudo conectar con Google Drive o procesar la información. "
            f"Verifica que la carpeta esté compartida con la cuenta de servicio. Detalle: {exc}"
        )
        st.stop()


def _render_resultados() -> None:
    """Muestra métricas y botones de descarga a partir del resultado en session_state."""
    resultado = st.session_state.get("resultado")
    if resultado is None:
        st.info("Selecciona un período y presiona **Procesar Comisiones** para comenzar.")
        return

    st.success(f"Período procesado: {resultado.periodo}")

    col_fac, col_base, col_com = st.columns(3)
    col_fac.metric("Total Facturas Procesadas", f"{resultado.total_facturas:,}")
    col_base.metric("Base Comisionable Global", f"${resultado.base_comisionable_global:,.2f}")
    col_com.metric("Total Comisiones a Pagar", f"${resultado.total_comisiones:,.2f}")

    st.divider()

    # ── Descarga del Resumen Global ───────────────────────────────────────────
    global_buffer, global_name = resultado.resumen_global
    global_buffer.seek(0)
    st.download_button(
        "⬇️ Descargar Resumen Global",
        data=global_buffer, file_name=global_name, mime=XLSX_MIME, type="primary",
    )

    # ── Descarga por vendedor (buscador) ──────────────────────────────────────
    st.subheader("Reporte individual por vendedor")
    vendedor = st.selectbox("Buscar KAM", resultado.vendedores, index=None,
                            placeholder="Escribe o selecciona un vendedor…")
    if vendedor:
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
