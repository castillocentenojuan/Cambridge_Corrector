"""
app.py — Interfaz Streamlit para Cambridge Writing Corrector
=============================================================
Ejecutar con:  streamlit run app.py
Requiere:      corrector_cambridge_v5.py en la misma carpeta
               redacciones/ con los PDFs de referencia
"""

import streamlit as st
from pathlib import Path

import corrector_cambridge_v5 as engine
from corrector_cambridge_v5 import (
    CambridgeCorrector,
    process_excel,
    get_essay_types,
    init_client,
)

# ─────────────────────────────────────────────────────────────
# CONFIGURACIÓN DE PÁGINA
# ─────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Cambridge Writing Corrector",
    page_icon="🎓",
    layout="centered",
)

st.markdown("""
<style>
/* Botón principal */
div[data-testid="stButton"] > button[kind="primary"] {
    background-color: #2e75b6;
    color: white;
    font-weight: bold;
    border-radius: 6px;
    height: 3em;
    width: 100%;
}
/* Botón de descarga */
div[data-testid="stDownloadButton"] > button {
    background-color: #1a7a4a;
    color: white;
    font-weight: bold;
    border-radius: 6px;
    height: 3em;
    width: 100%;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# ESTADO DE SESIÓN
# ─────────────────────────────────────────────────────────────

if "api_key"       not in st.session_state: st.session_state.api_key       = ""
if "running"       not in st.session_state: st.session_state.running       = False
if "result_bytes"  not in st.session_state: st.session_state.result_bytes  = None
if "result_name"   not in st.session_state: st.session_state.result_name   = ""
if "stats"         not in st.session_state: st.session_state.stats         = None

# ─────────────────────────────────────────────────────────────
# CABECERA
# ─────────────────────────────────────────────────────────────

st.title("🎓 Cambridge Writing Corrector")
st.caption("Corrección automática de redacciones C1/C2 · Powered by Llama 3.3 70B (Groq)")
st.markdown("---")

# ─────────────────────────────────────────────────────────────
# INICIALIZACIÓN DE MOTOR (SILENCIOSA)
# ─────────────────────────────────────────────────────────────
try:
    # Lee la clave oculta en la configuración de Streamlit Cloud (Advanced Settings -> Secrets)
    secret_key = st.secrets["LLAMA_API_KEY"]
    init_client(secret_key)
except KeyError:
    st.error("Error de servidor: Credenciales de IA no configuradas. Contacta al administrador.")
    st.stop()

# ─────────────────────────────────────────────────────────────
# FORMULARIO PRINCIPAL
# ─────────────────────────────────────────────────────────────

essay_types = get_essay_types()

col1, col2 = st.columns(2)
with col1:
    selected_type = st.selectbox(
        "📝 Tipo de redacción",
        options=essay_types,
        disabled=st.session_state.running,
        help="Debe existir el PDF de ejemplo en redacciones/<Tipo>/<Tipo>.pdf",
    )
with col2:
    uploaded_file = st.file_uploader(
        "📂 Sube el Excel de estudiantes",
        type=["xlsx"],
        disabled=st.session_state.running,
        help="El Excel debe tener las redacciones en la columna E.",
    )

# ─────────────────────────────────────────────────────────────
# VALIDACIONES PREVIAS (feedback inmediato sin esperar al botón)
# ─────────────────────────────────────────────────────────────

ready = True

if not st.session_state.api_key:
    st.info("👈 Introduce tu API Key en la barra lateral para continuar.")
    ready = False

if not essay_types:
    st.error("❌ No se encontraron tipos de redacción. Comprueba la carpeta `redacciones/`.")
    ready = False

if not uploaded_file and ready:
    st.info("⬆️ Sube el archivo Excel para continuar.")
    ready = False

# ─────────────────────────────────────────────────────────────
# BOTÓN DE ACCIÓN
# ─────────────────────────────────────────────────────────────

st.markdown("---")
boton_label = "⏳ Corrección en curso…" if st.session_state.running else "🚀 Corregir Redacciones"

if st.button(boton_label, type="primary", disabled=not ready or st.session_state.running):

    st.session_state.running      = True
    st.session_state.result_bytes = None
    st.session_state.stats        = None

    # ── Inicializar corrector (lee PDFs del disco) ────────────────────────────
    with st.spinner("Cargando rúbrica y ejemplo de corrección…"):
        try:
            corrector = CambridgeCorrector(selected_type)
        except Exception as e:
            st.error(f"❌ Error al cargar los archivos de referencia: {e}")
            st.session_state.running = False
            st.stop()

    excel_bytes = uploaded_file.read()
    output_name = uploaded_file.name.replace(".xlsx", "_CORREGIDO.xlsx")

    # ── Zona de progreso en tiempo real ──────────────────────────────────────
    st.markdown("### Progreso")
    progress_bar = st.progress(0.0)
    status_text  = st.empty()
    log_box      = st.empty()
    log_lines    = []

    def on_progress(current, total, name, status):
        pct = current / total
        progress_bar.progress(pct)

        if status == "skipped":
            log_lines.append(f"⏭️ `{name}` — Omitido (ya corregido)")
        elif status == "correcting":
            status_text.text(f"Corrigiendo {current}/{total}: {name}…")
        elif status.startswith("done:"):
            score = status.split(":", 1)[1]
            nota  = f"{score}/20" if score else "—"
            log_lines.append(f"✅ `{name}` — Nota: **{nota}**")
            status_text.text(f"Procesado {current}/{total}")
        elif status.startswith("error:"):
            msg = status.split(":", 1)[1]
            log_lines.append(f"❌ `{name}` — Error: {msg}")

        # Mostrar últimas 25 líneas para no saturar la pantalla
        log_box.markdown("\n".join(log_lines[-25:]))

    # ── Ejecutar proceso ──────────────────────────────────────────────────────
    try:
        output_buffer, stats = process_excel(
            excel_bytes=excel_bytes,
            corrector=corrector,
            progress_callback=on_progress,
        )
        st.session_state.result_bytes = output_buffer.getvalue()
        st.session_state.result_name  = output_name
        st.session_state.stats        = stats

    except Exception as e:
        st.error(f"❌ Error inesperado durante el proceso: {e}")

    finally:
        st.session_state.running = False
        progress_bar.progress(1.0)
        status_text.empty()

# ─────────────────────────────────────────────────────────────
# RESULTADO Y DESCARGA
# ─────────────────────────────────────────────────────────────

if st.session_state.result_bytes and st.session_state.stats:
    stats = st.session_state.stats
    st.markdown("---")
    st.markdown("### ✅ Proceso completado")

    # Métricas en columnas
    m1, m2, m3 = st.columns(3)
    m1.metric("Corregidas",  stats["success"])
    m2.metric("Omitidas",    stats["skipped"])
    m3.metric("Con error",   len(stats["failed"]))

    if stats["failed"]:
        st.warning("Alumnos con error: " + ", ".join(stats["failed"]))

    st.download_button(
        label="⬇️ Descargar Excel corregido",
        data=st.session_state.result_bytes,
        file_name=st.session_state.result_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.balloons()

# ─────────────────────────────────────────────────────────────
# PIE
# ─────────────────────────────────────────────────────────────

st.markdown("---")
st.caption("Cambridge Writing Corrector · Los datos no se almacenan en ningún servidor.")
