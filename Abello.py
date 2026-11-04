import streamlit as st
import sqlite3
import os
import python docx

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
        )
    ''')
    # Insertar algunas preguntas de prueba si la tabla está vacía
    cursor.execute('SELECT COUNT(*) FROM preguntas')
    if cursor.fetchone()[0] == 0:
        preguntas_iniciales = [
            ("¿Cuál es la derivada de f(x) = x^2?", "2x", "x", "2", "x^2", "A", "Cálculo", "Prof. Luis"),
            ("¿Cuánto es la suma de los ángulos internos de un triángulo?", "180°", "360°", "90°", "270°", "A", "Geometría", "Prof. Luis"),
            ("¿Cuál es el valor de Pi (aproximado)?", "3.1416", "2.718", "1.414", "1.732", "A", "Aritmética", "Prof. María")
        ]
        cursor.executemany('''
            INSERT INTO preguntas (enunciado, opcion_a, opcion_b, opcion_c, opcion_d, respuesta_correcta, categoria, docente)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ''', preguntas_iniciales)
        conn.commit()
    conn.close()

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

def insertar_pregunta(enunciado, a, b, c, d, correcta, categoria, docente):
    conn = sqlite3.connect('banco_preguntas.db')
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO preguntas (enunciado, opcion_a, opcion_b, opcion_c, opcion_d, respuesta_correcta, categoria, docente)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    ''', (enunciado, a, b, c, d, correcta, categoria, docente))
    conn.commit()
    conn.close()

def generar_word(preguntas_seleccionadas, incluir_respuestas):
    doc = python docx.Document()
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

# --- 3. INTERFAZ DE USUARIO (STREAMLIT) ---
st.title("📚 Sistema de Gestión y Generación de Evaluaciones")

menu = st.sidebar.selectbox("Navegación", ["Generar Evaluación", "Banco de Preguntas / Subir Preguntas"])

if menu == "Generar Evaluación":
    st.header("Generador de Evaluaciones a Medida")
    
    # Filtro por categoría
    categorias = ["Todas", "Cálculo", "Geometría", "Aritmética"]
    cat_elegida = st.selectbox("Filtrar por categoría:", categorias)
    
    preguntas_disponibles = obtener_preguntas(cat_elegida)
    
    if not preguntas_disponibles:
        st.info("No hay preguntas registradas en esta categoría.")
    else:
        st.write("Selecciona las preguntas que deseas incluir en el examen:")
        
        preguntas_seleccionadas_ids = []
        
        for p in preguntas_disponibles:
            p_id, enunciado, a, b, c, d, correcta, cat, docente = p
            # Creamos un checkbox por cada pregunta
            seleccionado = st.checkbox(f"**[{cat}]** {enunciado} *(Autor: {docente})*", key=f"preg_{p_id}")
            if seleccionado:
                preguntas_seleccionadas_ids.append(p)
                
        st.markdown("---")
        incluir_clave = st.checkbox("Incluir hoja de respuestas al final del documento")
        
        if st.button("📥 Descargar Examen en Word (.docx)"):
            if not preguntas_seleccionadas_ids:
                st.warning("Por favor, selecciona al menos una pregunta.")
            else:
                path_archivo = generar_word(preguntas_seleccionadas_ids, incluir_clave)
                with open(path_archivo, "rb") as f:
                    st.download_button(
                        label="Haz clic aquí para descargar tu archivo",
                        data=f,
                        file_name="evaluacion.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                st.success("¡Documento generado exitosamente!")

elif menu == "Banco de Preguntas / Subir Preguntas":
    st.header("Módulo Docente: Subir Pregunta al Banco")
    
    with st.form("form_subir_pregunta"):
        docente = st.text_input("Nombre del Docente:")
        categoria = st.selectbox("Categoría / Tema:", ["Cálculo", "Geometría", "Aritmética", "Estadística", "Álgebra"])
        enunciado = st.text_area("Enunciado de la pregunta:")
        
        col1, col2 = st.columns(2)
        with col1:
            opcion_a = st.text_input("Opción A:")
            opcion_b = st.text_input("Opción B:")
        with col2:
            opcion_c = st.text_input("Opción C:")
            opcion_d = st.text_input("Opción D:")
            
        respuesta_correcta = st.selectbox("Respuesta Correcta:", ["A", "B", "C", "D"])
        
        enviar = st.form_submit_button("Guardar Pregunta en el Sistema")
        
        if enviar:
            if docente and enunciado and opcion_a and opcion_b and opcion_c and opcion_d:
                insertar_pregunta(enunciado, opcion_a, opcion_b, opcion_c, opcion_d, respuesta_correcta, categoria, docente)
                st.success("¡Pregunta agregada correctamente al banco de datos!")
            else:
                st.error("Por favor completa todos los campos obligatorios.")
                
    st.markdown("---")
    st.subheader("Preguntas Registradas Actualmente en el Sistema")
    todas = obtener_preguntas("Todas")
    for p in todas:
        st.write(f"- **[{p[7]}]** {p[1]} *(Creada por: {p[8]})*")
