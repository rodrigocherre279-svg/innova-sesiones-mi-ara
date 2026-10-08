"""
Generador de Presentaciones de Clase - Innova Schools
Interfaz Web para la profesora Araceli (Streamlit Community Cloud)
"""

import os
import io
import time
import streamlit as st
from PIL import Image
import session_engine

# Configuracion de pagina
st.set_page_config(
    page_title="Innova Schools - Generador de Sesiones",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS personalizados para identidad Innova Schools
st.markdown("""
<style>
    .main-title {
        color: #00669E;
        font-family: 'Roboto', sans-serif;
        font-size: 2.2rem;
        font-weight: 700;
        margin-bottom: 0.2rem;
    }
    .sub-title {
        color: #555555;
        font-size: 1.1rem;
        margin-bottom: 1.5rem;
    }
    .info-card {
        background-color: #EDF5FA;
        border-left: 5px solid #00669E;
        padding: 1rem 1.2rem;
        border-radius: 8px;
        margin-bottom: 1rem;
    }
    .badge-obj {
        background-color: #741753;
        color: white;
        padding: 0.3rem 0.8rem;
        border-radius: 20px;
        font-size: 0.85rem;
        font-weight: 600;
    }
    .stDownloadButton button {
        background-color: #00669E !important;
        color: white !important;
        font-weight: bold !important;
        font-size: 1.1rem !important;
        padding: 0.6rem 2rem !important;
        border-radius: 8px !important;
        border: none !important;
        box-shadow: 0 4px 6px rgba(0, 102, 158, 0.2);
    }
    .stDownloadButton button:hover {
        background-color: #004d77 !important;
    }
</style>
""", unsafe_allow_html=True)

# Rutas de recursos
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(BASE_DIR, "assets")
LOGO_PATH = os.path.join(ASSETS_DIR, "innova_schools_logo.png")
SAMPLES_DIR = os.path.join(ASSETS_DIR, "samples")

# Encabezado
col_logo, col_header = st.columns([1, 4])
with col_logo:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=180)
    else:
        st.markdown("## 🏫 Innova")

with col_header:
    st.markdown("<div class='main-title'>Generador de Sesiones de Clase</div>", unsafe_allow_html=True)
    st.markdown("<div class='sub-title'>Convierte tus planes de sesión (.docx) en presentaciones interactivas de 16 diapositivas con el sistema didáctico oficial de secundaria.</div>", unsafe_allow_html=True)

# Barra lateral
with st.sidebar:
    st.markdown("### 👩‍🏫 Panel de Docente")
    st.markdown("**Profesora:** Araceli")
    st.markdown("**Nivel:** Secundaria / Inglés (A1)")
    st.markdown("---")
    
    st.markdown("#### 📌 ¿Cómo usar esta herramienta?")
    st.markdown("""
    1. **Sube tu archivo** de Diseño Instruccional (`.docx`).
    2. Revisa el **objetivo pedagógico** y los datos detectados.
    3. Haz clic en **'Generar Presentación PPTX'**.
    4. **Descarga tu archivo** listo para proyectar en el aula.
    """)
    st.markdown("---")
    st.caption("Sistema automatizado desarrollado para Innova Schools con metodología CLT, andamiaje y rúbricas oficiales.")

# Pestañas principales
tab_crear, tab_galeria, tab_estructura = st.tabs(["🚀 Crear Presentación", "🖼️ Recursos Visuales", "📋 Estructura de 16 Slides"])

with tab_crear:
    st.markdown("### 1. Sube tu Plan de Sesión (.docx)")
    uploaded_file = st.file_uploader(
        "Arrastra y suelta tu archivo Word de diseño instruccional aquí:",
        type=["docx"],
        help="Sube un archivo como A1_Unit5_Session2_Instructional_Design.docx"
    )

    if uploaded_file is not None:
        try:
            # Parsear el archivo subido
            plan_data = session_engine.parse_docx(uploaded_file)
            
            # Mostrar tarjeta con datos detectados
            st.markdown(f"""
            <div class='info-card'>
                <h4>📋 Plan Didáctico Detectado</h4>
                <p><b>Nivel y Unidad:</b> {plan_data['level']} • {plan_data['unit']} <i>("{plan_data['unit_title']}")</i></p>
                <p><b>Sesión:</b> Sesión N° {plan_data['session_num']} • <b>Área:</b> {plan_data['skill_area']}</p>
                <p><b>Gramática Principal:</b> <span style='color: #00669E; font-weight: bold;'>{plan_data['content_grammar']}</span></p>
                <p><b>Objetivo de Aprendizaje:</b> <span class='badge-obj'>{plan_data['objective']}</span></p>
                <p><b>Vocabulario Clave:</b> {plan_data['content_vocab'][:120] if plan_data['content_vocab'] else 'Vocabulario temático del nivel'}</p>
            </div>
            """, unsafe_allow_html=True)

            # Boton de generacion
            st.markdown("### 2. Generar Diapositivas")
            if st.button("✨ Generar Presentación PPTX Completa", type="primary"):
                progress_bar = st.progress(0)
                status_text = st.empty()

                status_text.text("Analizando alineamiento curricular y desempeños...")
                progress_bar.progress(20)
                time.sleep(0.4)

                status_text.text("Preparando recursos gráficos, organizadores y temporizadores...")
                progress_bar.progress(50)
                time.sleep(0.4)

                status_text.text(f"Construyendo las 16 diapositivas para {plan_data['unit']} ({plan_data['unit_title']})...")
                progress_bar.progress(80)
                
                # Construir PPTX
                uploaded_file.seek(0)
                pptx_buffer = session_engine.build_presentation(plan_data)
                
                progress_bar.progress(100)
                status_text.text("¡Presentación generada con éxito!")
                time.sleep(0.2)
                st.success("🎉 ¡Tu presentación está lista para descargar!")

                # Boton de descarga
                filename_out = f"{plan_data['level'].replace(' ', '_')}_Unit{plan_data['unit'].replace('Unit ', '')}_Session{plan_data['session_num']}_PPT.pptx"
                
                st.download_button(
                    label=f"📥 Descargar {filename_out}",
                    data=pptx_buffer,
                    file_name=filename_out,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )

                # Vista previa resumida dinamica
                st.markdown("---")
                st.markdown("#### ✨ Resumen de la Presentación Generada:")
                col_c1, col_c2 = st.columns(2)
                
                is_historical = "was" in plan_data['content_grammar'].lower() or plan_data['unit_num'] == "6" and plan_data['session_num'] == "6"
                
                with col_c1:
                    if is_historical:
                        st.markdown("""
                        * ✅ **Slide 1**: Portada Unit 6: "Important events" (`#00669E`).
                        * ✅ **Slide 2**: Reglas de clase (*Our golden rules* 01-04).
                        * ✅ **Slide 3**: Warm-up con foto de Machu Picchu y preguntas.
                        * ✅ **Slide 4**: Objetivo histórico en fondo Magenta (`#741753`).
                        * ✅ **Slide 5**: Texto didáctico: El descubrimiento de Machu Picchu (1911).
                        * ✅ **Slide 6**: Trabajo colaborativo: Roles Interviewer & Speaker.
                        * ✅ **Slide 7**: Discusión oral: Preguntas + reto de usar *was/were* + timer.
                        * ✅ **Slide 8**: Check-in con tipografía *Chewy* y sticker *"Good Job!"*.
                        """)
                    else:
                        st.markdown(f"""
                        * ✅ **Slide 1**: Portada {plan_data['unit']}: "{plan_data['unit_title']}".
                        * ✅ **Slide 2**: Reglas de clase (*Our golden rules* 01-04).
                        * ✅ **Slide 3**: Warm-up situacional con imagen y andamiaje oral.
                        * ✅ **Slide 4**: Diapositiva de Objetivo en fondo Magenta (`#741753`).
                        * ✅ **Slide 5**: Entrada de vocabulario temático con cuadrícula.
                        * ✅ **Slide 6**: Diálogo modelo guiado con precios o estructuras.
                        * ✅ **Slide 7**: Tabla de expresiones útiles y frases clave.
                        * ✅ **Slide 8**: Trabajo colaborativo con roles definidos.
                        """)
                with col_c2:
                    if is_historical:
                        st.markdown("""
                        * ✅ **Slide 9**: Organizador gráfico histórico (*Brainstorming organizer*).
                        * ✅ **Slide 10**: Tarea oral del discurso histórico + 4 Criterios de Éxito.
                        * ✅ **Slide 11**: Foco gramatical inductivo: Past Simple con **WAS** y **WERE**.
                        * ✅ **Slide 12**: Rúbrica oficial de evaluación Innova (C, B, A, AD).
                        * ✅ **Slide 13**: Redacción del borrador en el cuaderno con temporizador.
                        * ✅ **Slide 14**: Presentación oral (*Showtime*) y coevaluación con 3 estrellas.
                        * ✅ **Slide 15**: Conciencia lingüística: Errores comunes con *was* vs *were*.
                        * ✅ **Slide 16**: Autoevaluación final (*Can-do statements*).
                        """)
                    else:
                        st.markdown("""
                        * ✅ **Slide 9**: Check-in con tipografía *Chewy* y sticker *"Good Job!"*.
                        * ✅ **Slide 10**: Tarea de producción con 4 Criterios de Éxito.
                        * ✅ **Slide 11**: Hoja de notas de lluvia de ideas con temporizador.
                        * ✅ **Slide 12**: Rúbrica oficial de evaluación Innova (C, B, A, AD).
                        * ✅ **Slide 13**: Práctica de ensayo con temporizador.
                        * ✅ **Slide 14**: Actuación frente a la clase y coevaluación por estrellas.
                        * ✅ **Slide 15**: Corrección de errores comunes (*Common mistakes*).
                        * ✅ **Slide 16**: Autoevaluación final (*Can-do statements*).
                        """)

        except Exception as e:
            st.error(f"Error al procesar el archivo: {e}")
            st.info("Asegúrate de subir un archivo de diseño instruccional (.docx) válido.")
    else:
        st.info("👆 Por favor sube un archivo `.docx` para comenzar. Si quieres probar con un ejemplo, usa los planes ubicados en la carpeta `PLAN_DE_SESION`.")

with tab_galeria:
    st.markdown("### 🖼️ Galería de Recursos Visuales Integrados")
    st.markdown("Estas imágenes fotorrealistas e institucionales se insertan automáticamente en las diapositivas para asegurar un alto estándar pedagógico:")
    
    col_img1, col_img2, col_img3 = st.columns(3)
    
    img1_path = os.path.join(SAMPLES_DIR, "cafe_ordering_scene_1791484978565.jpg")
    img2_path = os.path.join(SAMPLES_DIR, "cafe_food_grid_1791485095937.jpg")
    img3_path = os.path.join(SAMPLES_DIR, "cafe_menu_card_1791485139541.jpg")

    with col_img1:
        if os.path.exists(img1_path):
            st.image(img1_path, caption="Slide 3: Escena de Diálogo (Mesero y Estudiantes)", use_container_width=True)
    with col_img2:
        if os.path.exists(img2_path):
            st.image(img2_path, caption="Slide 5: Collage de 6 Comidas y Bebidas", use_container_width=True)
    with col_img3:
        if os.path.exists(img3_path):
            st.image(img3_path, caption="Slide 6: Menú en Pizarra (The Corner Cafe)", use_container_width=True)

    st.markdown("---")
    st.markdown("#### Recursos Institucionales y Pedagógicos:")
    col_ic1, col_ic2, col_ic3, col_ic4 = st.columns(4)
    with col_ic1:
        ic_target = os.path.join(ASSETS_DIR, "objective_target_icon.png")
        if os.path.exists(ic_target):
            st.image(ic_target, caption="Slide 4: Icono de Objetivo", width=120)
    with col_ic2:
        ic_goodjob = os.path.join(ASSETS_DIR, "good_job_sticker.png")
        if os.path.exists(ic_goodjob):
            st.image(ic_goodjob, caption="Slide 9: Sticker 'Good Job!'", width=140)
    with col_ic3:
        ic_timer = os.path.join(ASSETS_DIR, "timer_03_00.jpg")
        if os.path.exists(ic_timer):
            st.image(ic_timer, caption="Slide 11: Temporizador Digital", width=150)
    with col_ic4:
        ic_rubric = os.path.join(ASSETS_DIR, "innova_rubric_table.png")
        if os.path.exists(ic_rubric):
            st.image(ic_rubric, caption="Slide 12: Rúbrica Oficial (C, B, A, AD)", width=170)

with tab_estructura:
    st.markdown("### 📋 Estructura Oficial de la Sesión (16 Diapositivas)")
    st.markdown("Cada presentación sigue el flujo de 3 fases (*Pre-Stage, Core Stage, Post-Stage*):")

    with st.expander("Fase 1: Pre-Stage (Activación, Normas y Vocabulario - 25 min)"):
        st.markdown("""
        * **Slide 1 (Portada)**: Fondo azul Innova (`#00669E`), título de unidad, nivel y sesión.
        * **Slide 2 (Reglas de Oro)**: Normas de convivencia del aula (01 Levantar la mano, 02 Escuchar instrucciones, 03 Participar activamente, 04 Ser amable).
        * **Slide 3 (Warm-up)**: Imagen detonante en un café, elicitación de vocabulario (*waiter, waitress, customer, menu*) y burbujas de andamiaje oral.
        * **Slide 4 (Objetivo)**: Fondo magenta (`#741753`), fecha formal, etiqueta de sesión e icono de diana con flecha.
        * **Slide 5 (Vocabulario)**: Cuadrícula fotográfica de 6 comidas con numeración y modelo de oración.
        * **Slide 6 (Diálogo Modelo)**: Menú en pizarra y conversación completa para rellenar vacíos.
        * **Slide 7 (Expresiones Útiles)**: Tarjetas de frases para el mesero vs frases para el cliente, con consejos de entonación cortés.
        """)

    with st.expander("Fase 2: Core Stage (Trabajo Colaborativo y Ensayo - 25 min)"):
        st.markdown("""
        * **Slide 8 (Roles)**: Grupos de 3 con roles definidos: Rol 1 (Mesero/a), Rol 2 (Cliente 1), Rol 3 (Cliente 2).
        * **Slide 9 (Check-in)**: Pregunta de comprobación *"Did we complete the activity successfully?"* con fuente Chewy y sticker *"Good Job!"*.
        * **Slide 10 (Criterios de Éxito)**: Tarea de conversación con 4 criterios claros y evaluables.
        * **Slide 11 (Lluvia de Ideas)**: Tarjetas de notas por rol y temporizador de 3 minutos.
        * **Slide 12 (Rúbricas)**: Rúbrica institucional de evaluación con niveles C, B, A y AD.
        * **Slide 13 (Ensayo Fase 1)**: Práctica con notas en grupos de 3 y temporizador de 2 minutos.
        """)

    with st.expander("Fase 3: Post-Stage (Actuación, Feedback y Reflexión - 10 min)"):
        st.markdown("""
        * **Slide 14 (Showtime)**: Presentación sin notas frente a la clase y coevaluación con sistema de 3 estrellas.
        * **Slide 15 (Conciencia Lingüística)**: Análisis de errores frecuentes en la pizarra vs inglés cortés y correcto.
        * **Slide 16 (Reflexión)**: Declaraciones de logro (*Can-do statements*) para autoevaluación de los alumnos.
        """)
