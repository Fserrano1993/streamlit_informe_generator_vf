from docx import Document
from docx.shared import Inches
import streamlit as st
import tempfile
import os

def generate_doc(placeholders, fotos, is_juridico):
    template_path = "PLANTILLA BASE JURIDICO - V3 generador.docx" if is_juridico else "PLANTILLA BASE - V3 generador.docx"
    doc = Document(template_path)

    # Reemplazo de placeholders en texto
    for p in doc.paragraphs:
        for key, value in placeholders.items():
            if key in p.text:
                p.text = p.text.replace(key, value)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, value in placeholders.items():
                    if key in cell.text:
                        cell.text = cell.text.replace(key, value)

    # Añadir imágenes al final
    if fotos:
        doc.add_page_break()
        doc.add_paragraph("Reportaje fotográfico")
        table = doc.add_table(rows=3, cols=2)
        table.autofit = True
        idx = 0
        for i in range(3):
            for j in range(2):
                if idx < len(fotos):
                    cell = table.cell(i, j)
                    cell.paragraphs[0].add_run().add_picture(fotos[idx], width=Inches(2.2))
                    idx += 1

    # Guardar
    tmpdir = tempfile.mkdtemp()
    filepath = os.path.join(tmpdir, f"{placeholders.get('{{EXPEDIENTE}}', 'informe')}.docx")
    doc.save(filepath)
    return filepath

# Interfaz Streamlit
st.set_page_config(page_title="Generador de Informes", layout="centered")
st.title("Generador de Informes Periciales")

st.markdown("### Paso 1: Datos del Encargo")
texto_encargo = st.text_area("Pega aquí el texto del encargo")
is_juridico = st.checkbox("¿Es defensa jurídica?", value=False)
fotos = st.file_uploader("Sube las fotos (máx. 6)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

if st.button("Generar informe"):
    if not texto_encargo:
        st.error("Debe pegar el texto del encargo.")
    else:
        # Sustituye esto por un parser real en producción
        placeholders = {
            "{{EXPEDIENTE}}": "SIN_EXP",
            "{{ASEGURADO}}": "Nombre Apellido",
            "{{DIR_CATASTRO}}": "Dirección de ejemplo",
            "{{PROVINCIA_CATASTRO}}": "Barcelona",
            "{{SUPERFICIE_CATASTRO}}": "100",
            "{{AÑO_CATASTRO}}": "1990",
            "{{USO_CATASTRO}}": "Residencial"
        }
        paths_fotos = []
        if fotos:
            for f in fotos:
                temp_path = os.path.join(tempfile.gettempdir(), f.name)
                with open(temp_path, "wb") as tmpf:
                    tmpf.write(f.read())
                paths_fotos.append(temp_path)

        path = generate_doc(placeholders, paths_fotos, is_juridico)
        with open(path, "rb") as file:
            st.download_button("Descargar informe Word", data=file, file_name=os.path.basename(path), mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
