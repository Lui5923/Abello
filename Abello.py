import streamlit as st
import sqlite3
import os
import io

# Importación segura de docx
try:
    import docx
except ImportError:
    st.error("⚠️ La librería 'python-docx' no se encuentra instalada en el servidor. Asegúrate de tener un archivo llamado exactamente 'requirements.txt' con el texto 'streamlit' y 'python-docx'.")
    st.stop()

# --- 1. CONFIGURACIÓN DE LA BASE DE DATOS SQLITE ---
def init_db():
    conn = sqlite3.connect('banco_preguntas.db')
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS preguntas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            enunciado TEXT NOT NULL,
            opcion_a TEXT,
            opcion_b TEXT,
            opcion_c TEXT,
            opcion_d TEXT,
            respuesta_correcta TEXT,
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
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (enunciado, a, b, c, d, correcta, grado, categoria, docente, imagen_bytes))
    conn.commit()
    conn.close()

def generar_word_evaluacion(preguntas_seleccionadas, grado_eval, incluir_respuestas):
    doc = docx.Document()
    
    # --- PLANTILLA DE ENCABEZADO PARA EL EXAMEN ---
    doc.add_heading(f'EVALUACIÓN ACADÉMICA - GRADO {grado_eval}', level=0)
    doc.add_paragraph("Nombre del Estudiante: __________________________________________________")
    doc.add_paragraph(f"Grado: {grado_eval}       Fecha: _________________       Calificación: ____________")
    doc.add_paragraph("Instrucciones: Lea atentamente cada enunciado, observe las figuras si las hay y responda lo solicitado.\n")
    doc.add_paragraph("--------------------------------------------------------------------------------------------------------------------")
    
    # --- CUERPO DE LAS PREGUNTAS Y OPCIONES ---
    for i, p in enumerate(preguntas_seleccionadas, 1):
        p_id, enunciado, a, b, c, d, correcta, grado, cat, docente, imagen = p
        
        # Enunciado
        doc.add_paragraph(f"{i}. {enunciado}")
        
        # Insertar imagen en el Word si la pregunta tiene una asociada
        if imagen:
            image_stream = io.BytesIO(imagen)
            try:
                doc.add_picture(image_stream, width=docx.shared.Inches(3.5))
            except Exception:
                pass
                
        # Mostrar opciones SOLO si fueron escritas de verdad y no contienen texto genérico de relleno
        if a and "por definir" not in a.lower():
            doc.add_paragraph(f"   A) {a}")
        if b and "por definir" not in b.lower():
            doc.add_paragraph(f"   B) {b}")
        if c and "por definir" not in c.lower():
            doc.add_paragraph(f"   C) {c}")
        if d and "por definir" not in d.lower():
            doc.add_paragraph(f"   D) {d}")
            
        doc.add_paragraph("") # Espacio entre preguntas
        
    # --- HOJA DE RESPUESTAS (OPCIONAL) ---
    if incluir_respuestas:
        doc.add_page_break()
        doc.add_heading('CLAVE DE RESPUESTAS (USO EXCLUSIVO DEL DOCENTE)', level=1)
        for i, p in enumerate(preguntas_seleccionadas, 1):
            if p[6] and p[6].strip():
                doc.add_paragraph(f"{i}. Opción correcta: {p[6]} *(Grado: {p[7]})*")
            else:
                doc.add_paragraph(f"{i}. Pregunta abierta *(Grado: {p[7]})*")
            
    archivo_salida = f"evaluacion_grado_{grado_eval}.docx"
    doc.save(archivo_salida)
    return archivo_salida

# --- 3. INTERFAZ DE USUARIO CON STREAMLIT ---
st.sidebar.image("https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcSdknrErzN7GfW0RgOyaT5WN7PiV9Hw0clyslV0V9yBjNRYzKj1h9T8gnZ4&s=10", width=150)
st.title("Sistema Integral de Banco de Preguntas por Grados")

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
            
            col_check, col_img = st.columns([4, 1])
            with col_check:
                seleccionado = st.checkbox(f"**[Grado: {grado}] [{cat}]** {enunciado} *(Autor: {docente})*", key=f"chk_{p_id}")
            with col_img:
                if imagen:
                    st.image(imagen, width=90, caption="Con imagen")
            
            if seleccionado:
                preguntas_seleccionadas.append(p)
                
        st.markdown("---")
        incluir_clave = st.checkbox("Incluir clave de respuestas al final del documento Word")
        
        if st.button("📥 Descargar Examen en Word (.docx)"):
            if not preguntas_seleccionadas:
                st.warning("Por favor, marca al menos una pregunta.")
            else:
                path_archivo = generar_word_evaluacion(preguntas_seleccionadas, grado_elegido if grado_elegido != "Todos" else "Varios", incluir_clave)
                with open(path_archivo, "rb") as f:
                    st.download_button(
                        label="Hacer clic aquí para descargar tu archivo Word",
                        data=f,
                        file_name=f"evaluacion_grado_{grado_elegido}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                st.success("¡Examen en Word generado limpio y sin opciones de relleno!")

elif menu == "Registrar Pregunta Manual":
    st.header("Módulo Docente: Crear Pregunta con Opciones e Imagen")
    
    with st.form("form_manual"):
        docente = st.text_input("Nombre del Docente:")
        
        col_g2, col_c2 = st.columns(2)
        with col_g2:
            grado_reg = st.selectbox("Selecciona el Grado al que pertenece:", grados_disponibles)
        with col_c2:
            categoria = st.selectbox("Categoría:", categorias_disponibles)
            
        enunciado = st.text_area("Enunciado de la pregunta:")
        imagen_subida = st.file_uploader("Adjuntar imagen para la pregunta (Opcional)", type=["png", "jpg", "jpeg"])
        
        col1, col2 = st.columns(2)
        with col1:
            opcion_a = st.text_input("Opción A (Opcional):")
            opcion_b = st.text_input("Opción B (Opcional):")
        with col2:
            opcion_c = st.text_input("Opción C (Opcional):")
            opcion_d = st.text_input("Opción D (Opcional):")
            
        respuesta_correcta = st.selectbox("Selecciona la Opción Correcta:", ["Ninguna / Abierta", "A", "B", "C", "D"])
        
        enviar = st.form_submit_button("Guardar Pregunta en el Banco")
        
        if enviar:
            if docente and enunciado:
                img_bytes = imagen_subida.read() if imagen_subida else None
                rc = respuesta_correcta if respuesta_correcta != "Ninguna / Abierta" else ""
                insertar_pregunta_db(enunciado, opcion_a, opcion_b, opcion_c, opcion_d, rc, grado_reg, categoria, docente, img_bytes)
                st.success("¡Pregunta guardada correctamente!")
            else:
                st.error("Por favor completa el nombre del docente y el enunciado.")

elif menu == "Subir Documento Word":
    st.header("Módulo Docente: Subir Examen o Documento en Word")
    st.write("Sube un archivo `.docx`. Las preguntas se registrarán limpias, sin opciones falsas ni respuestas automáticas.")
    
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
                try:
                    doc_subido = docx.Document(archivo_word)
                    contador = 0
                    for para in doc_subido.paragraphs:
                        texto = para.text.strip()
                        # Evitar párrafos vacíos o líneas cortas de encabezado que no sean preguntas
                        if len(texto) > 3:
                            img_blob = None
                            for run in para.runs:
                                elementos_dibujo = run._r.xpath('.//a:blip')
                                for blip in elementos_dibujo:
                                    embed_id = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                                    if embed_id in doc_subido.part.rels:
                                        img_blob = doc_subido.part.rels[embed_id].target_part.blob
                                        break

                            insertar_pregunta_db(
                                enunciado=texto,
                                a="",
                                b="",
                                c="",
                                d="",
                                correcta="",
                                grado=grado_subida,
                                categoria=categoria,
                                docente=docente,
                                imagen_bytes=img_blob
                            )
                            contador += 1
                    st.success(f"¡Se han importado exitosamente {contador} preguntas limpias para el Grado {grado_subida}!")
                except Exception as e:
                    st.error(f"Error al procesar el archivo: {e}")
            else:
                st.error("Por favor ingresa tu nombre y adjunta un archivo válido.")
