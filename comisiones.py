"""
DEPRECADO — la lógica se movió a `src/comisiones.py`.
=====================================================
Este archivo se conserva solo como puente de compatibilidad. La aplicación Streamlit
importa todo desde `src.comisiones`. Puedes borrar este archivo con seguridad.

    from src.comisiones import run_pipeline, load_data_sources_from_drive, ...
"""

from src.comisiones import *  # noqa: F401,F403
