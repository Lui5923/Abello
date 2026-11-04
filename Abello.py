import streamlit as st
import sqlite3
import io

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
            categoria TEXT NOT NULL,
            docente TEXT
 
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
            categoria TEXT NOT NULL,
            docente TEXT
        )
    ''')


# --- 2. FUNCIONES DE APOYO ---
def obtener_preguntas(categoria_filtro="Todas"):
    conn = sqlite3.connect('banco_preguntas.db')
    cursor = conn.cursor()
    if categoria_filtro == "Todas":
        cursor.execute("SELECT * FROM preguntas")
    else:
        cursor.execute("SELECT * FROM preguntas WHERE categoria = ?", (categoria_filtro,))
    data = cursor.fetchall()
    conn.close()
    return data

def insertar_pregunta(enunciado, a, b, c, d, correcta, categoria, docente):
    conn = sqlite3.connect('banco_preguntas.db')
    cursor = conn.cursor()
    cursor.execute(
        INSERT INTO preguntas (enunciado, opcion_a, opcion_b, opcion_c, opcion_d, respuesta_correcta, categoria, docente)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)(enunciado, a, b, c, d, correcta, categoria, docente))
    conn.commit()
    conn.close()

def generar_word(preguntas_seleccionadas, incluir_respuestas):
    doc = docx.Document()
    doc.add_heading('Evaluación Académica', 0)
    doc.add_paragraph('Seleccione las respuestas correctas o desarrolle los siguientes ítems:\n')
    
    for i, p in enumerate(preguntas_seleccionadas, 1):
        # p = (id, enunciado, a, b, c, d, correcta, categoria, docente)
        doc.add_paragraph(f"{i}. {p[1]}")
        doc.add_paragraph(f"   A) {p[2]}")
        doc.add_paragraph(f"   B) {p[3]}")
        doc.add_paragraph(f"   C) {p[4]}")
        doc.add_paragraph(f"   D) {p[5]}")
        doc.add_paragraph("") # Espacio
        
    if incluir_respuestas:
        doc.add_page_break()
        doc.add_heading('Clave de Respuestas', level=1)
        for i, p in enumerate(preguntas_seleccionadas, 1):
            doc.add_paragraph(f"{i}. {p[6]}")
            
    archivo_salida = "evaluacion_generada.docx"
    doc.save(archivo_salida)
    return archivo_salida

init_db()

# --- 2. FUNCIONES DE APOYO ---
def obtener_preguntas(categoria_filtro="Todas"):
    conn = sqlite3.connect('banco_preguntas.db')
    cursor = conn.cursor()
    if categoria_filtro == "Todas":
        cursor.execute("SELECT * FROM preguntas")
    else:
        cursor.execute("SELECT * FROM preguntas WHERE categoria = ?", (categoria_filtro,))
    data = cursor.fetchall()
    conn.close()
    return data

def procesar_word_subido(archivo_subido, categoria, docente):
    """Lee un archivo Word subido y extrae párrafos para agregarlos como preguntas"""
    try:
        doc = docx.Document(archivo_subido)
        contador = 0
        for para in doc.paragraphs:
            texto = para.text.strip()
            if texto: # Si el párrafo no está vacío
                # Guardamos cada párrafo como una pregunta general de apoyo
                insertar_pregunta(
                    enunciado=texto,
                    a="Opción A por defecto",
                    b="Opción B por defecto",
                    c="Opción C por defecto",
                    d="Opción D por defecto",
                    correcta="A",
                    categoria=categoria,
                    docente=docente
                )
                contador += 1
        return contador
    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")
        return 0

def generar_word_evaluacion(preguntas_seleccionadas, incluir_respuestas):
    doc = docx.Document()
    doc.add_heading('Evaluación Académica Generada', 0)
    doc.add_paragraph('Seleccione o desarrolle las respuestas correspondientes:\n')
    
    for i, p in enumerate(preguntas_seleccionadas, 1):
        # p = (id, enunciado, a, b, c, d, correcta, categoria, docente)
        doc.add_paragraph(f"{i}. {p[1]}")
        doc.add_paragraph(f"   A) {p[2]}")
        doc.add_paragraph(f"   B) {p[3]}")
        doc.add_paragraph(f"   C) {p[4]}")
        doc.add_paragraph(f"   D) {p[5]}")
        doc.add_paragraph("") # Espacio
        
    if incluir_respuestas:
        doc.add_page_break()
        doc.add_heading('Clave de Respuestas', level=1)
        for i, p in enumerate(preguntas_seleccionadas, 1):
            doc.add_paragraph(f"{i}. Clave: {p[6]} *(Pregunta de: {p[8]})*")
            
    archivo_salida = "evaluacion_final.docx"
    doc.save(archivo_salida)
    return archivo_salida

# --- 3. INTERFAZ DE USUARIO (STREAMLIT) ---
st.title("Sistema de Gestión y Banco de Preguntas")
st.title("Institución Educativa Distrital Técnica Bilingüe Jorge Nicolás Abello")

menu = st.sidebar.selectbox("Navegación", ["Generar Evaluación", "Subir Documento / Preguntas"])

if menu == "Generar Evaluación":
    st.header("Seleccionar Preguntas y Descargar Word")
    
    categorias = ["Todas", "Cálculo", "Geometría", "Aritmética", "Estadística", "Álgebra"]
    cat_elegida = st.selectbox("Filtrar por categoría:", categorias)
    
    preguntas_disponibles = obtener_preguntas(cat_elegida)
    
    if not preguntas_disponibles:
        st.info("No hay preguntas registradas todavía. ¡Sube un documento o agrega preguntas en la otra sección!")
    else:
        st.write("Selecciona las casillas de las preguntas que deseas exportar:")
        preguntas_seleccionadas = []
        
        for p in preguntas_disponibles:
            p_id, enunciado, a, b, c, d, correcta, cat, docente = p
            seleccionado = st.checkbox(f"**[{cat}]** {enunciado} *(Autor: {docente})*", key=f"chk_{p_id}")
            if seleccionado:
                preguntas_seleccionadas.append(p)
                
        st.markdown("---")
        incluir_clave = st.checkbox("Incluir hoja de respuestas al final del documento Word")
        
        if st.button("📥 Descargar Examen en Word (.docx)"):
            if not preguntas_seleccionadas:
                st.warning("Por favor, marca al menos una pregunta.")
            else:
                path_archivo = generar_word_evaluacion(preguntas_seleccionadas, incluir_clave)
                with open(path_archivo, "rb") as f:
                    st.download_button(
                        label="Hacer clic aquí para descargar el archivo",
                        data=f,
                        file_name="evaluacion.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                st.success("¡Archivo Word generado con éxito!")

elif menu == "Subir Documento / Preguntas":
    st.header("Módulo Docente: Subir Documento Word")
    st.write("Sube un archivo `.docx` con tus preguntas para que el sistema las extraiga y las ponga a disposición de todos los docentes.")
    
    with st.form("form_subir_doc"):
        docente = st.text_input("Nombre del Docente:")
        categoria = st.selectbox("Categoría de las preguntas:", ["Cálculo", "Geometría", "Aritmética", "Estadística", "Álgebra"])
        archivo_word = st.file_uploader("Sube tu archivo de Word (.docx)", type=["docx"])
        
        enviar = st.form_submit_button("Procesar y Guardar en el Banco")
        
        if enviar:
            if docente and archivo_word:
                total_agregadas = procesar_word_subido(archivo_word, categoria, docente)
                if total_agregadas > 0:
                    st.success(f"¡Se han extraído e integrado {total_agregadas} elementos al banco de datos exitosamente!")
                else:
                    st.warning("El archivo parecía vacío o no se pudieron leer párrafos.")
            else:
                st.error("Por favor ingresa tu nombre y adjunta un archivo válido.")
                
    st.markdown("---")
    st.subheader("Banco Actual de Preguntas en el Sistema")
    todas = obtener_preguntas("Todas")
    if not todas:
        st.write("Aún no hay registros.")
    else:
        for p in todas:
            st.write(f"- **[{p[7]}]** {p[1]} *(Subido por: {p[8]})*")
