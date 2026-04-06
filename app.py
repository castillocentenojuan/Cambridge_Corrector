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
)

import uuid

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
if "session_id" not in st.session_state:
    st.session_state.session_id = uuid.uuid4().hex

# ─────────────────────────────────────────────────────────────
# CABECERA
# ─────────────────────────────────────────────────────────────

st.title("🎓 Cambridge Writing Corrector")
st.caption("Corrección automática de redacciones C1/C2 · Powered by Llama 3.3 70B (Groq)")
st.markdown("---")

# ─────────────────────────────────────────────────────────────
# FORMULARIO PRINCIPAL
# ─────────────────────────────────────────────────────────────

essay_types = get_essay_types()

selected_model = st.selectbox(
    "🤖 Modelo de IA",
    options=["Llama 3.3 70B (Groq)", "Gemini 2.0 Flash (Google)"],
    disabled=st.session_state.running
)

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

if not essay_types:
    st.error("❌ No se encontraron tipos de redacción. Comprueba la carpeta `redacciones/`.")
    ready = False

if not uploaded_file and ready:
    st.info("⬆️ Sube el archivo Excel para continuar.")
    ready = False

# ─────────────────────────────────────────────────────────────
# ─────────────────────────────────────────────────────────────
# BOTONES DE ACCIÓN (Controladores)
# ─────────────────────────────────────────────────────────────
st.markdown("---")

col1, col2 = st.columns(2)

with col1:
    # 1. El interruptor de encendido
    if st.button("🚀 Corregir Redacciones", type="primary", disabled=not ready or st.session_state.running):
        st.session_state.running = True
        st.session_state.result_bytes = None
        st.session_state.stats = None
        st.rerun()  # 🔄 Recarga la página para mostrar la interfaz de "En curso"

with col2:
    # 2. El interruptor de apagado de emergencia
    if st.session_state.running:
        if st.button("⏹️ Detener corrección", type="secondary"):
            st.session_state.running = False
            st.rerun()

# ─────────────────────────────────────────────────────────────
# EJECUCIÓN DEL MOTOR
# ─────────────────────────────────────────────────────────────

# Si el interruptor está encendido, el motor trabaja
if st.session_state.running:
    
    # ── Inicializar corrector ────────────────────────────
    with st.spinner("Cargando rúbrica y conectando con la IA…"):
        try:
            if "Gemini" in selected_model:
                p_provider = "gemini"
                p_key = st.secrets["GEMINI_API_KEY"]
            else:
                p_provider = "groq"
                p_key = st.secrets["LLAMA_API_KEY"]

            corrector = CambridgeCorrector(
                essay_type=selected_type, 
                provider=p_provider, 
                api_key=p_key
            )
        except Exception as e:
            st.error(f"❌ Error al inicializar: {e}")
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
            log_lines.append(f"⏭️ `{name}` — Omitido")
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

        log_box.markdown("\n".join(log_lines[-25:]))

    # ── Llamada al proceso ───────────────────────────────────────────────────
    try:
        def on_save(current_bytes: bytes):
            st.session_state.result_bytes = current_bytes

        output_buffer, stats = process_excel(
            excel_bytes=excel_bytes,
            corrector=corrector,
            progress_callback=on_progress,
            save_callback=on_save,
            stop_flag=lambda: False # Ignoramos la lógica de la otra IA, Streamlit hará de verdugo.
        )
        st.session_state.result_bytes = output_buffer.getvalue()
        st.session_state.result_name  = output_name
        st.session_state.stats        = stats

    except Exception as e:
        st.error(f"❌ Error inesperado: {e}")

    finally:
        st.session_state.running = False
        progress_bar.progress(1.0)
        status_text.empty()
        st.rerun() # 🔄 Al terminar, recarga la página para mostrar el botón de descarga verde.
        
# ─────────────────────────────────────────────────────────────
# RESULTADO Y DESCARGA
# ─────────────────────────────────────────────────────────────

if st.session_state.result_bytes and st.session_state.stats:
    stats = st.session_state.stats
    st.markdown("---")
    st.markdown("### ✅ Proceso completado")

    # Solo mostrar métricas si el proceso terminó limpio y devolvió stats
    if st.session_state.stats:
        stats = st.session_state.stats
        m1, m2, m3 = st.columns(3)
        m1.metric("Corregidas",  stats["success"])
        m2.metric("Omitidas",    stats["skipped"])
        m3.metric("Con error",   len(stats["failed"]))

        if stats["failed"]:
            st.warning("Alumnos con error: " + ", ".join(stats["failed"]))
    else:
        st.warning("⚠️ El proceso se interrumpió antes de terminar, pero puedes descargar el progreso guardado hasta el momento.")

    # Asegurar que output_name exista en la sesión (por si falló muy rápido)
    if not st.session_state.result_name:
        st.session_state.result_name = "CORRECCION_PARCIAL.xlsx"

    st.download_button(
        label="⬇️ Descargar Excel corregido",
        data=st.session_state.result_bytes,
        file_name=st.session_state.result_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    if st.session_state.stats:
        st.balloons()

# ─────────────────────────────────────────────────────────────
# PIE
# ─────────────────────────────────────────────────────────────

st.markdown("---")
st.caption("Cambridge Writing Corrector · Los datos no se almacenan en ningún servidor.")
