import streamlit as st
from docx import Document
from io import BytesIO
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from datetime import datetime

# Diccionario para asociar idioma con plantilla
PLANTILLAS = {
    "Español": "Carta Tipo - Incorporaciones.docx",
    "Portugués": "Carta Tipo - IncorporacionesPOR.docx",
    "Inglés": "Carta Tipo - IncorporacionesENG.docx",
}

# Función para reemplazar texto en el documento
def reemplazar_campos(template_path, reemplazos):
    doc = Document(template_path)

    for para in doc.paragraphs:
        for key, value in reemplazos.items():
            if key in para.text:
                inline_text = "".join(run.text for run in para.runs)
                inline_text = inline_text.replace(key, value)
                for run in para.runs:
                    run.text = ""
                para.runs[0].text = inline_text

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, value in reemplazos.items():
                    if key in cell.text:
                        cell.text = cell.text.replace(key, value)

    return doc

st.title("Generador de Carta de Incorporaciones")

idioma = st.selectbox("Seleccione el idioma", list(PLANTILLAS.keys()))

# Campos del formulario
nombre = st.text_input("Inserte Nombre")
localizador = st.text_input("Inserte Localizador")
fecha_input = st.text_input("Inserte Fecha (DD/MM/YYYY)")
ciudad = st.text_input("Inserte Ciudad")
trayecto = st.text_input("Inserte Trayecto")
hora_presentacion = st.text_input("Inserte Hora de Presentación")
hora_salida = st.text_input("Inserte Hora de Salida")
punto_encuentro = st.text_input("Inserte Punto de Encuentro")
direccion = st.text_input("Inserte Dirección")

# Validación de fecha y obtención del día
try:
    fecha_obj = datetime.strptime(fecha_input, "%d/%m/%Y")
    dia_semana = fecha_obj.strftime("%A")  # Día en inglés
    dias_traducidos = {
        "Español": {
            "Monday": "Lunes", "Tuesday": "Martes", "Wednesday": "Miércoles",
            "Thursday": "Jueves", "Friday": "Viernes", "Saturday": "Sábado", "Sunday": "Domingo"
        },
        "Portugués": {
            "Monday": "Segunda-feira", "Tuesday": "Terça-feira", "Wednesday": "Quarta-feira",
            "Thursday": "Quinta-feira", "Friday": "Sexta-feira", "Saturday": "Sábado", "Sunday": "Domingo"
        },
        "Inglés": {
            "Monday": "Monday", "Tuesday": "Tuesday", "Wednesday": "Wednesday",
            "Thursday": "Thursday", "Friday": "Friday", "Saturday": "Saturday", "Sunday": "Sunday"
        }
    }
    dia_traducido = dias_traducidos[idioma][dia_semana]
    fecha_valida = True
except ValueError:
    st.error("Formato de fecha inválido. Use el formato DD/MM/YYYY.")
    fecha_valida = False

# Reemplazos si la fecha es válida
if fecha_valida and st.button("Generar Documento"):
    reemplazos = {
        "(INSERTENOMBRE)": nombre,
        "(LOCALIZADOR)": localizador,
        "(INSERTEFECHA)": fecha_input,
        "(DIA)": dia_traducido,
        "(CIUDAD)": ciudad,
        "(INSERTETRAYECTO)": trayecto,
        "(HORAPRESENTACION)": hora_presentacion,
        "(HORASALIDA)": hora_salida,
        "(PUNTOENCUENTRO)": punto_encuentro,
        "(INSERTEDIRECCION)": direccion
    }

    plantilla = PLANTILLAS[idioma]
    doc = reemplazar_campos(plantilla, reemplazos)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    st.download_button(
        label="Descargar Documento",
        data=buffer,
        file_name=f"{localizador}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
