# 🦉 Comisiones — Buho Media

Aplicación web (Streamlit) que calcula las comisiones de vendedores a partir de 4 Excel
alojados en una carpeta privada de **Google Drive**, leídos con una **Service Account**.
Genera los reportes en memoria y los ofrece para descarga directa. Sin terminal y sin
subir archivos a mano.

## Arquitectura

```
app.py                  # UI Streamlit (login, controles, métricas, descargas, admin)
src/comisiones.py       # capa de datos Drive + validación + pipeline de comisiones
.streamlit/
  secrets.toml          # credenciales (LOCAL, gitignored)
  secrets.toml.example  # plantilla
requirements.txt
README_OPERACIONES.md   # guía para Finanzas/Operaciones
```

- **Fuentes**: se resuelven por prefijo de nombre en la carpeta de Drive (el .xlsx más
  reciente). Config en `FILE_PATTERNS` / `CANONICAL_NAMES` (`src/comisiones.py`).
- **Validación**: `REQUIRED_COLUMNS` + `DataValidationError` con mensajes amigables.
- **Salidas**: `io.BytesIO` → `st.download_button` (no se escribe a disco).

## Ejecutar en local

1. `pip install -r requirements.txt`
2. Crea `.streamlit/secrets.toml` a partir de `.streamlit/secrets.toml.example` con:
   - `[gcp_service_account]` — el JSON de la Service Account.
   - `[drive] folder_id` — carpeta de Drive compartida con el `client_email` de la SA.
   - `[auth] username` / `password` — credenciales del login.
3. `python -m streamlit run app.py`

## Desplegar en Streamlit Community Cloud

1. Repo en GitHub (privado). **No** subir `.streamlit/secrets.toml` ni el JSON de la SA
   (ya están en `.gitignore`).
2. share.streamlit.io → New app → repo, branch `main`, archivo `app.py`.
3. *Advanced settings* → Python **3.12** o **3.13**.
4. Pega el contenido de `secrets.toml` en *Settings → Secrets*.
5. Deploy.

## Seguridad

- Credenciales solo vía `st.secrets`; nada hardcodeado.
- La Service Account debe tener acceso (editor) a la carpeta de Drive.
- Rotar la llave de la SA si se expuso.
