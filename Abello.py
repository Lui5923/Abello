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
            tema TEXT,
            docente TEXT,
            imagen BLOB
        )
    ''')
    # Actualización automática si la base de datos ya existía sin la columna 'tema'
    cursor.execute("PRAGMA table_info(preguntas)")
    columnas = [col[1] for col in cursor.fetchall()]
    if 'tema' not in columnas:
        cursor.execute("ALTER TABLE preguntas ADD COLUMN tema TEXT")
    conn.commit()
    conn.close()

init_db()

# --- 2. FUNCIONES DE APOYO Y GESTIÓN DE DATOS ---
def obtener_preguntas(grado_filtro="Todos", categoria_filtro="Todos", tema_filtro="Todos"):
    conn = sqlite3.connect('banco_preguntas.db')
    cursor = conn.cursor()
    
    query = "SELECT id, enunciado, opcion_a, opcion_b, opcion_c, opcion_d, respuesta_correcta, grado, categoria, tema, docente, imagen FROM preguntas WHERE 1=1"
    params = []
    
    if grado_filtro != "Todos":
        query += " AND grado = ?"
        params.append(grado_filtro)
        
    if categoria_filtro != "Todos":
        query += " AND categoria = ?"
        params.append(categoria_filtro)
        
    if tema_filtro != "Todos":
        query += " AND tema = ?"
        params.append(tema_filtro)
        
    cursor.execute(query, params)
    data = cursor.fetchall()
    conn.close()
    return data

def insertar_pregunta_db(enunciado, a, b, c, d, correcta, grado, categoria, tema, docente, imagen_bytes=None):
    conn = sqlite3.connect('banco_preguntas.db')
    cursor = conn.cursor()
    # Corrección exacta: 12 columnas en la tabla frente a los 12 valores enviados
    cursor.execute('''
        INSERT INTO preguntas (enunciado, opcion_a, opcion_b, opcion_c, opcion_d, respuesta_correcta, grado, categoria, tema, docente, imagen)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (enunciado, a, b, c, d, correcta, grado, categoria, tema, docente, imagen_bytes))
    conn.commit()
    conn.close()

def eliminar_pregunta_db(pregunta_id):
    conn = sqlite3.connect('banco_preguntas.db')
    cursor = conn.cursor()
    cursor.execute("DELETE FROM preguntas WHERE id = ?", (pregunta_id,))
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
        p_id, enunciado, a, b, c, d, correcta, grado, cat, tema, docente, imagen = p
        
        doc.add_paragraph(f"{i}. {enunciado}")
        
        if imagen:
            image_stream = io.BytesIO(imagen)
            try:
                doc.add_picture(image_stream, width=docx.shared.Inches(3.5))
            except Exception:
                pass
                
        if a and "por definir" not in a.lower():
            doc.add_paragraph(f"   A) {a}")
        if b and "por definir" not in b.lower():
            doc.add_paragraph(f"   B) {b}")
        if c and "por definir" not in c.lower():
            doc.add_paragraph(f"   C) {c}")
        if d and "por definir" not in d.lower():
            doc.add_paragraph(f"   D) {d}")
            
        doc.add_paragraph("") 
        
    # --- HOJA DE RESPUESTAS (OPCIONAL) ---
    if incluir_respuestas:
        doc.add_page_break()
        doc.add_heading('CLAVE DE RESPUESTAS (USO EXCLUSIVO DEL DOCENTE)', level=1)
        for i, p in enumerate(preguntas_seleccionadas, 1):
            if p[6] and p[6].strip():
                doc.add_paragraph(f"{i}. Opción correcta: {p[6]} *(Grado: {p[7]} | Tema: {p[9]})*")
            else:
                doc.add_paragraph(f"{i}. Pregunta abierta *(Grado: {p[7]} | Tema: {p[9]})*")
            
    archivo_salida = f"evaluacion_grado_{grado_eval}.docx"
    doc.save(archivo_salida)
    return archivo_salida

# --- 3. INTERFAZ DE USUARIO CON STREAMLIT ---
st.sidebar.image("https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcSdknrErzN7GfW0RgOyaT5WN7PiV9Hw0clyslV0V9yBjNRYzKj1h9T8gnZ4&s=10", width=150)
st.title("Institución Educativa Distrital Técnica Bilingüe Jorge Nicolás Abello")
st.title("Sistema Integral de Banco de Preguntas por Grados")
st.sidebar.markdown("### Menú de Navegación")

grados_disponibles = ["3°", "4°", "5°", "6°", "7°", "8°", "9°", "10°", "11°"]
categorias_disponibles = ["Cálculo", "Geometría", "Aritmética", "Estadística", "Álgebra"]

menu = st.sidebar.selectbox("Selecciona una opción", ["Generar Examen", "Registrar Pregunta Manual", "Subir Documento Word", "Administrar / Eliminar Preguntas"])

if menu == "Generar Examen":
    st.header("Generador de Evaluaciones por Grado")
    
    col_g1, col_c1, col_t1 = st.columns(3)
    with col_g1:
        grado_elegido = st.selectbox("Selecciona el Grado:", ["Todos"] + grados_disponibles)
    with col_c1:
        cat_elegida = st.selectbox("Filtrar por categoría:", ["Todas"] + categorias_disponibles)
    with col_t1:
        preguntas_filtro_prev = obtener_preguntas(grado_elegido, cat_elegida, "Todos")
        temas_existentes = sorted(list(set([p[9] for p in preguntas_filtro_prev if p[9]])))
        tema_elegido = st.selectbox("Filtrar por Tema:", ["Todos"] + temas_existentes)
    
    preguntas_disponibles = obtener_preguntas(grado_elegido, cat_elegida, tema_elegido)
    
    if not preguntas_disponibles:
        st.info("No hay preguntas registradas para los filtros seleccionados.")
    else:
        st.write("📖 **Haz clic en cada pregunta para leer su enunciado completo y seleccionarla:**")
        preguntas_seleccionadas = []
        
        for p in preguntas_disponibles:
            p_id, enunciado, a, b, c, d, correcta, grado, cat, tema, docente, imagen = p
            
            with st.expander(f"[Grado: {grado}] [{cat}] Tema: {tema or 'General'} - {enunciado[:50]}... (Autor: {docente})"):
                st.markdown(f"**Tema:** `{tema or 'General'}`")
                st.markdown(f"**Enunciado completo:**\n\n{enunciado}")
                
                if a and "por definir" not in a.lower():
                    st.write(f"**A)** {a}")
                if b and "por definir" not in b.lower():
                    st.write(f"**B)** {b}")
                if c and "por definir" not in c.lower():
                    st.write(f"**C)** {c}")
                if d and "por definir" not in d.lower():
                    st.write(f"**D)** {d}")
                
                if imagen:
                    st.image(imagen, width=220, caption="Imagen asociada a la pregunta")
                
                seleccionado = st.checkbox("✔ Incluir esta pregunta en el examen", key=f"chk_{p_id}")
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
                st.success("¡Examen en Word generado con éxito!")

elif menu == "Registrar Pregunta Manual":
    st.header("Módulo Docente: Crear Pregunta con Tema, Opciones e Imagen")
    
    with st.form("form_manual"):
        docente = st.text_input("Nombre del Docente:")
        
        col_g2, col_c2 = st.columns(2)
        with col_g2:
            grado_reg = st.selectbox("Selecciona el Grado:", grados_disponibles)
        with col_c2:
            categoria = st.selectbox("Categoría (Asignatura):", categorias_disponibles)
            
        tema_manual = st.text_input("Tema de la pregunta (ej. Teorema de Pitágoras, Funciones Cuadráticas):")
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
            if docente and enunciado and tema_manual:
                img_bytes = imagen_subida.read() if imagen_subida else None
                rc = respuesta_correcta if respuesta_correcta != "Ninguna / Abierta" else ""
                insertar_pregunta_db(enunciado, opcion_a, opcion_b, opcion_c, opcion_d, rc, grado_reg, categoria, tema_manual, docente, img_bytes)
                st.success("¡Pregunta guardada correctamente con su respectivo tema!")
            else:
                st.error("Por favor completa el nombre del docente, el tema y el enunciado.")

elif menu == "Subir Documento Word":
    st.header("Módulo Docente: Subir Examen o Documento en Word")
    st.write("Sube un archivo `.docx`. El sistema detectará un tema automático que podrás modificar si es incorrecto.")
    
    docente = st.text_input("Nombre del Docente:")
    
    col_g3, col_c3 = st.columns(2)
    with col_g3:
        grado_subida = st.selectbox("Grado correspondiente:", grados_disponibles)
    with col_c3:
        categoria = st.selectbox("Categoría asociada:", categorias_disponibles)
        
    archivo_word = st.file_uploader("Selecciona el archivo Word (.docx)", type=["docx"])
    
    tema_sugerido = "Evaluación general"
    if archivo_word:
        try:
            doc_temp = docx.Document(archivo_word)
            for p_temp in doc_temp.paragraphs:
                t_temp = p_temp.text.strip()
                if len(t_temp) > 5:
                    tema_sugerido = t_temp[:60]
                    break
        except Exception:
            pass
            
    tema_word = st.text_input("Tema detectado automáticamente (Puedes modificarlo si está mal):", value=tema_sugerido)
    
    if st.button("Procesar y Extraer Contenido"):
        if docente and archivo_word and tema_word:
            try:
                doc_subido = docx.Document(archivo_word)
                contador = 0
                for para in doc_subido.paragraphs:
                    texto = para.text.strip()
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
                            tema=tema_word,
                            docente=docente,
                            imagen_bytes=img_blob
                        )
                        contador += 1
                st.success(f"¡Se han importado exitosamente {contador} preguntas bajo el tema '{tema_word}' para el Grado {grado_subida}!")
            except Exception as e:
                st.error(f"Error al procesar el archivo: {e}")
        else:
            st.error("Por favor completa tu nombre, el tema y adjunta un archivo válido.")

elif menu == "Administrar / Eliminar Preguntas":
    st.header("🗑️ Administrar y Eliminar Preguntas")
    st.write("Busca preguntas por grado, categoría o tema y elimínalas si encuentras algún error.")
    
    col_del1, col_del2 = st.columns(2)
    with col_del1:
        grado_del = st.selectbox("Filtrar por Grado:", ["Todos"] + grados_disponibles, key="filtro_del_grado")
    with col_del2:
        cat_del = st.selectbox("Filtrar por Categoría:", ["Todos"] + categorias_disponibles, key="filtro_del_cat")
        
    preguntas_para_borrar = obtener_preguntas(grado_del, cat_del, "Todos")
    
    if not preguntas_para_borrar:
        st.info("No hay preguntas registradas con los filtros seleccionados.")
    else:
        st.write(f"Se encontraron **{len(preguntas_para_borrar)}** preguntas:")
        
        for p in preguntas_para_borrar:
            p_id, enunciado, a, b, c, d, correcta, grado, cat, tema, docente, imagen = p
            
            with st.expander(f"📌 [Grado: {grado}] [{cat}] Tema: {tema} - {enunciado[:50]}... (Autor: {docente})"):
                st.write(f"**Tema:** {tema}")
                st.write(f"**Enunciado completo:** {enunciado}")
                st.write(f"**Grado:** {grado} | **Categoría:** {cat} | **Autor:** {docente}")
                if imagen:
                    st.image(imagen, width=150, caption="Imagen asociada")
                
                if st.button(f"🗑️ Eliminar esta pregunta (ID: {p_id})", key=f"btn_del_{p_id}"):
                    eliminar_pregunta_db(p_id)
                    st.success(f"¡Pregunta ID {p_id} eliminada correctamente!")
                    st.rerun()
