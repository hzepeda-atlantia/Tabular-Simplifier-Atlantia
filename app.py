import streamlit as st
import backend
import os

# --- Page Config ---
st.set_page_config(
    page_title="Tabular Simplifier",
    page_icon="📊",
    layout="wide"
)

# --- Load Custom CSS ---
def load_css(file_name):
    with open(file_name) as f:
        st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)

load_css("styles.css")

# --- Header ---
st.markdown('<h1 style="text-align: center; margin-bottom: 2rem;">📊 <span class="gradient-text">Tabular Simplifier</span></h1>', unsafe_allow_html=True)

# --- Explanation ---
st.info(
    "**Criterio para Marcar Diferencias Significativas:**  \n"
    "Las diferencias se calculan y resaltan cuando un dato cumple con las siguientes condiciones:  \n"
    "1.  **Base Mínima:** El grupo cuenta con el número suficiente de respuestas para ser considerado válido.  \n"
    "2.  **Diferencia Estadística:** El dato es significativamente mayor que al menos **la mitad** de los demás grupos válidos."
)

# --- Sidebar ---
with st.sidebar:
    st.image("assets/logo.png", use_container_width=True)
    st.markdown("### ⚙️ Configuración")
    
    mode = st.radio(
        "Modo de Procesamiento",
        ["Full Processing", "DS Only"],
        help="Elige el nivel de detalle del reporte."
    )

    sheet_name = st.text_input(
        "Nombre de la Hoja", 
        value="T2", 
        help="La hoja del Excel que contiene los datos principales."
    )

    base_min = st.number_input(
        "Mínimo de Base",
        min_value=0,
        max_value=500,
        value=50,
        step=1,
        help="Bases menores a este número no serán consideradas para significancia."
    )

    segment_pdp = False
        
    st.markdown("---")
    st.info("ℹ️ **Tip:** Asegúrate de que tu archivo Excel tenga la estructura correcta.")

# --- Main Content ---
# Use a more centered layout for the upload card
col1, col2, col3 = st.columns([1, 6, 1])

with col2:
    # Card Container with glassmorphism via CSS
    with st.container(border=True):
        st.markdown("### 📤 Cargar Archivo")
        st.markdown("Sube tu archivo `.xlsx` para comenzar la simplificación.")
        
        uploaded_file = st.file_uploader("", type=["xlsx"])
        
        if uploaded_file:
            st.success(f"✅ Archivo cargado: **{uploaded_file.name}**")
            
            st.markdown("<br>", unsafe_allow_html=True)
            
            if st.button("🚀 Procesar Archivo", type="primary", use_container_width=True):
                with st.spinner("✨ Simplificando tabulares..."):
                    try:
                        output_path = None
                        
                        if mode == "Full Processing":
                            output_path = backend.process_workbook_by_pdp(
                                uploaded_file,
                                segment_with_pdp=segment_pdp,
                                general_sheet=sheet_name,
                                base_min=base_min
                            )
                        else: # DS Only
                            output_path = backend.process_workbook_ds_only(
                                uploaded_file,
                                sheet_name=sheet_name,
                                base_min=base_min
                            )
                        
                        if output_path and os.path.exists(output_path):
                            st.balloons()
                            st.markdown("### 🎉 ¡Simplificación Completada!")
                            st.markdown("Tu archivo está listo para descargar.")
                            
                            # Read file for download
                            with open(output_path, "rb") as f:
                                file_data = f.read()
                            
                            out_name = f"PROCESSED_{uploaded_file.name}"
                            
                            st.download_button(
                                label="📥 Descargar Resultado",
                                data=file_data,
                                file_name=out_name,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                            
                            # Cleanup temp file
                            os.remove(output_path)
                            
                    except Exception as e:
                        st.error(f"❌ Ocurrió un error: {e}")
                        st.exception(e)

        else:
            st.info("Esperando archivo...")