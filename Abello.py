import streamlit as st
import sqlite3
import os
import io
try:
    import docx
except ImportError:
    st.error("⚠️ La librería 'python-docx' no está instalada en este entorno. Ejecuta 'pip install python-docx' en tu terminal.")
    st.stop()

# --- 1. CONFIGURACIÓN DE LA BASE DE DATOS SQLITE ---
def init_db():
    conn = sqlite3.connect('banco_preguntas.db')
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS preguntas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            enunciado TEXT NOT NULL,
            opcion_a TEXT NOT NULL,
            opcion_b TEXT NOT NULL,
            opcion_c TEXT NOT NULL,
            opcion_d TEXT NOT NULL,
            respuesta_correcta TEXT NOT NULL,
            grado TEXT NOT NULL,
            categoria TEXT NOT NULL,
            docente TEXT,
            imagen BLOB
        )
    ''')
    conn.commit()
    conn.close()

init_db()

# --- 2. FUNCIONES DE APOYO Y GESTIÓN DE DATOS ---
def obtener_preguntas(grado_filtro="Todos", categoria_filtro="Todas"):
    conn = sqlite3.connect('banco_preguntas.db')
    cursor = conn.cursor()
    
    query = "SELECT id, enunciado, opcion_a, opcion_b, opcion_c, opcion_d, respuesta_correcta, grado, categoria, docente, imagen FROM preguntas WHERE 1=1"
    params = []
    
    if grado_filtro != "Todos":
        query += " AND grado = ?"
        params.append(grado_filtro)
        
    if categoria_filtro != "Todas":
        query += " AND categoria = ?"
        params.append(categoria_filtro)
        
    cursor.execute(query, params)
    data = cursor.fetchall()
    conn.close()
    return data

def insertar_pregunta_db(enunciado, a, b, c, d, correcta, grado, categoria, docente, imagen_bytes=None):
    conn = sqlite3.connect('banco_preguntas.db')
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO preguntas (enunciado, opcion_a, opcion_b, opcion_c, opcion_d, respuesta_correcta, grado, categoria, docente, imagen)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (enunciado, a, b, c, d, correcta, grado, categoria, docente, imagen_bytes))
    conn.commit()
    conn.close()

def procesar_word_con_imagenes(archivo_subido, grado, categoria, docente):
    """Extrae párrafos e imágenes incrustadas de un archivo de Word"""
    try:
        doc = docx.Document(archivo_subido)
        contador = 0

        for para in doc.paragraphs:
            texto = para.text.strip()
            if texto:
                # Buscar si el párrafo tiene una imagen asociada en su estructura interna
                img_blob = None
                for run in para.runs:
                    elementos_dibujo = run._r.xpath('.//a:blip')
                    for blip in elementos_dibujo:
                        embed_id = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                        if embed_id in doc.part.rels:
                            target_part = doc.part.rels[embed_id].target_part
                            img_blob = target_part.blob
                            break

                insertar_pregunta_db(
                    enunciado=texto,
                    a="Opción A por defecto",
                    b="Opción B por defecto",
                    c="Opción C por defecto",
                    d="Opción D por defecto",
                    correcta="A",
                    grado=grado,
                    categoria=categoria,
                    docente=docente,
                    imagen_bytes=img_blob
                )
                contador += 1
        return contador
    except Exception as e:
        st.error(f"Error al procesar el archivo Word: {e}")
        return 0

def generar_word_evaluacion(preguntas_seleccionadas, grado_eval, incluir_respuestas):
    doc = docx.Document()
    doc.add_heading(f'Evaluación Académica - Grado {grado_eval}', 0)
    doc.add_paragraph('Selecciona la respuesta correcta para cada uno de los siguientes ítems:\n')
    
    for i, p in enumerate(preguntas_seleccionadas, 1):
        p_id, enunciado, a, b, c, d, correcta, grado, cat, docente, imagen = p
        
        doc.add_paragraph(f"{i}. {enunciado}")
        
        # Si la pregunta tiene una imagen almacenada, la insertamos en el documento Word
        if imagen:
            image_stream = io.BytesIO(imagen)
            try:
                doc.add_picture(image_stream, width=docx.shared.Inches(3.5))
            except Exception:
                pass
                
        doc.add_paragraph(f"   A) {a}")
        doc.add_paragraph(f"   B) {b}")
        doc.add_paragraph(f"   C) {c}")
        doc.add_paragraph(f"   D) {d}")
        doc.add_paragraph("") # Espacio
        
    if incluir_respuestas:
        doc.add_page_break()
        doc.add_heading('Clave de Respuestas', level=1)
        for i, p in enumerate(preguntas_seleccionadas, 1):
            doc.add_paragraph(f"{i}. Opción correcta: {p[6]}")
            
    archivo_salida = f"evaluacion_grado_{grado_eval}.docx"
    doc.save(archivo_salida)
    return archivo_salida

# --- 3. INTERFAZ DE USUARIO CON STREAMLIT ---
st.title("📚 Sistema Integral de Banco de Preguntas por Grados")

grados_disponibles = ["3°", "4°", "5°", "6°", "7°", "8°", "9°", "10°", "11°"]
categorias_disponibles = ["Cálculo", "Geometría", "Aritmética", "Estadística", "Álgebra"]

menu = st.sidebar.selectbox("Navegación", ["Generar Examen", "Registrar Pregunta Manual", "Subir Documento Word"])

if menu == "Generar Examen":
    st.header("Generador de Evaluaciones por Grado")
    
    col_g1, col_c1 = st.columns(2)
    with col_g1:
        grado_elegido = st.selectbox("Selecciona el Grado a evaluar:", ["Todos"] + grados_disponibles)
    with col_c1:
        cat_elegida = st.selectbox("Filtrar por categoría:", ["Todas"] + categorias_disponibles)
    
    preguntas_disponibles = obtener_preguntas(grado_elegido, cat_elegida)
    
    if not preguntas_disponibles:
        st.info("No hay preguntas registradas para los filtros seleccionados.")
    else:
        st.write(f"Mostrando preguntas disponibles. Selecciona las que deseas incluir en la evaluación:")
        preguntas_seleccionadas = []
        
        for p in preguntas_disponibles:
            p_id, enunciado, a, b, c, d, correcta, grado, cat, docente, imagen = p
            
            col_check, col_img = st.columns([4, 1])
            with col_check:
                seleccionado = st.checkbox(f"**[Grado: {grado}] [{cat}]** {enunciado} *(Autor: {docente})*", key=f"chk_{p_id}")
            with col_img:
                if imagen:
                    st.image(imagen, width=100, caption="Incluye imagen")
            
            if seleccionado:
                preguntas_seleccionadas.append(p)
                
        st.markdown("---")
        incluir_clave = st.checkbox("Incluir clave de respuestas al final del documento Word")
        
        if st.button("📥 Descargar Examen en Word (.docx)"):
            if not preguntas_seleccionadas:
                st.warning("Por favor, marca al menos una pregunta.")
            else:
                path_archivo = generar_word_evaluacion(preguntas_seleccionadas, grado_elegido, incluir_clave)
                with open(path_archivo, "rb") as f:
                    st.download_button(
                        label="Hacer clic aquí para descargar el archivo",
                        data=f,
                        file_name=f"evaluacion_grado_{grado_elegido}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                st.success("¡Documento generado con éxito!")

elif menu == "Registrar Pregunta Manual":
    st.header("Módulo Docente: Crear Pregunta con Opciones")
    
    with st.form("form_manual"):
        docente = st.text_input("Nombre del Docente:")
        
        col_g2, col_c2 = st.columns(2)
        with col_g2:
            grado_reg = st.selectbox("Selecciona el Grado al que pertenece:", grados_disponibles)
        with col_c2:
            categoria = st.selectbox("Categoría:", categorias_disponibles)
            
        enunciado = st.text_area("Enunciado de la pregunta:")
        
        col1, col2 = st.columns(2)
        with col1:
            opcion_a = st.text_input("Opción A:")
            opcion_b = st.text_input("Opción B:")
        with col2:
            opcion_c = st.text_input("Opción C:")
            opcion_d = st.text_input("Opción D:")
            
        respuesta_correcta = st.selectbox("Selecciona la Opción Correcta:", ["A", "B", "C", "D"])
        
        enviar = st.form_submit_button("Guardar Pregunta en el Banco")
        
        if enviar:
            if docente and enunciado and opcion_a and opcion_b and opcion_c and opcion_d:
                insertar_pregunta_db(enunciado, opcion_a, opcion_b, opcion_c, opcion_d, respuesta_correcta, grado_reg, categoria, docente, None)
                st.success("¡Pregunta guardada correctamente en el sistema para el grado seleccionado!")
            else:
                st.error("Por favor completa todos los campos obligatorios.")

elif menu == "Subir Documento Word":
    st.header("Módulo Docente: Subir Examen o Documento en Word")
    st.write("Sube un archivo `.docx` asignándole un grado específico. El sistema extraerá los textos y las imágenes de forma automática[span_1](start_span)[span_1](end_span).")
    
    with st.form("form_subir"):
        docente = st.text_input("Nombre del Docente:")
        
        col_g3, col_c3 = st.columns(2)
        with col_g3:
            grado_subida = st.selectbox("Grado al que corresponden estas preguntas:", grados_disponibles)
        with col_c3:
            categoria = st.selectbox("Categoría asociada:", categorias_disponibles)
            
        archivo_word = st.file_uploader("Selecciona el archivo Word (.docx)", type=["docx"])
        
        procesar = st.form_submit_button("Procesar y Extraer Contenido")
        
        if procesar:
            if docente and archivo_word:
                total = procesar_word_con_imagenes(archivo_word, grado_subida, categoria, docente)
                if total > 0:
                    st.success(f"¡Se han importado exitosamente {total} elementos para el Grado {grado_subida}!")
                else:
                    st.warning("No se detectaron textos válidos en el documento.")
            else:
                st.error("Por favor ingresa tu nombre y adjunta un archivo válido.")
