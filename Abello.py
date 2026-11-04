import streamlit as st
import sqlite3
import os
import io
import zipfile
import xml.etree.ElementTree as ET

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

# --- 2. FUNCIONES DE APOYO Y LECTURA NATIVA DE WORD ---
def obtener_preguntas(grado_filtro="Todos", categoria_filtro="Todos"):
    conn = sqlite3.connect('banco_preguntas.db')
    cursor = conn.cursor()
    
    query = "SELECT id, enunciado, opcion_a, opcion_b, opcion_c, opcion_d, respuesta_correcta, grado, categoria, docente, imagen FROM preguntas WHERE 1=1"
    params = []
    
    if grado_filtro != "Todos":
        query += " AND grado = ?"
        params.append(grado_filtro)
        
    if categoria_filtro != "Todos":
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

def leer_word_nativo(archivo_subido):
    """Lee párrafos de un archivo .docx usando únicamente herramientas nativas de Python (sin instalar docx)"""
    parrafos = []
    try:
        with zipfile.ZipFile(archivo_subido, 'r') as z:
            # Un archivo .docx es un archivo zip con un xml de contenido llamado word/document.xml
            xml_content = z.read('word/document.xml')
            root = ET.fromstring(xml_content)
            
            # Espacios de nombres XML de Word
            namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            
            # Buscar todos los párrafos (<w:p>)
            for p in root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'):
                textos = []
                for t in p.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
                    if t.text:
                        textos.append(t.text)
                parrafo_completo = "".join(textos).strip()
                if parrafo_completo:
                    parrafos.append(parrafo_completo)
    except Exception as e:
        st.error(f"Error al leer la estructura interna del Word: {e}")
    return parrafos

def procesar_word_subido(archivo_subido, grado, categoria, docente):
    """Procesa el archivo Word nativo y guarda las preguntas"""
    try:
        parrafos = leer_word_nativo(archivo_subido)
        contador = 0
        for texto in parrafos:
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
                imagen_bytes=None
            )
            contador += 1
        return contador
    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")
        return 0

def generar_word_evaluacion(preguntas_seleccionadas, grado_eval, incluir_respuestas):
    """Genera un archivo de texto plano con formato descargable en caso de emergencia, o simula la descarga"""
    # Creamos un contenido de texto estructurado para el archivo descargado
    contenido = f"EVALUACIÓN ACADÉMICA - GRADO {grado_eval}\n" + "="*40 + "\n\n"
    
    for i, p in enumerate(preguntas_seleccionadas, 1):
        p_id, enunciado, a, b, c, d, correcta, grado, cat, docente, imagen = p
        contenido += f"{i}. {enunciado}\n"
        contenido += f"   A) {a}\n"
        contenido += f"   B) {b}\n"
        contenido += f"   C) {c}\n"
        contenido += f"   D) {d}\n\n"
        
    if incluir_respuestas:
        contenido += "\n" + "="*20 + "\nCLAVE DE RESPUESTAS\n" + "="*20 + "\n"
        for i, p in enumerate(preguntas_seleccionadas, 1):
            contenido += f"{i}. {p[6]}\n"
            
    return contenido

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
        st.write("Selecciona las casillas de las preguntas que deseas incluir en la evaluación:")
        preguntas_seleccionadas = []
        
        for p in preguntas_disponibles:
            p_id, enunciado, a, b, c, d, correcta, grado, cat, docente, imagen = p
            seleccionado = st.checkbox(f"**[Grado: {grado}] [{cat}]** {enunciado} *(Autor: {docente})*", key=f"chk_{p_id}")
            if seleccionado:
                preguntas_seleccionadas.append(p)
                
        st.markdown("---")
        incluir_clave = st.checkbox("Incluir clave de respuestas al final del documento")
        
        if st.button("📥 Descargar Examen"):
            if not preguntas_seleccionadas:
                st.warning("Por favor, marca al menos una pregunta.")
            else:
                texto_archivo = generar_word_evaluacion(preguntas_seleccionadas, grado_elegido, incluir_clave)
                st.download_button(
                    label="Haz clic aquí para descargar el archivo de evaluación",
                    data=texto_archivo,
                    file_name=f"evaluacion_grado_{grado_elegido}.txt",
                    mime="text/plain"
                )
                st.success("¡Examen listo para descargar!")

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
    st.write("Sube un archivo `.docx`. El sistema leerá el texto de forma nativa sin errores y lo asignará al grado seleccionado.")
    
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
                total = procesar_word_subido(archivo_word, grado_subida, categoria, docente)
                if total > 0:
                    st.success(f"¡Se han importado exitosamente {total} elementos para el Grado {grado_subida}!")
                else:
                    st.warning("No se pudieron extraer textos del documento.")
            else:
                st.error("Por favor ingresa tu nombre y adjunta un archivo válido.")
