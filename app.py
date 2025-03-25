import streamlit as st
import docx
import datetime
import os

def load_template(language):
    templates = {
        "Español": "Carta Tipo - Incorporaciones - copia.docx",
        "Portugués": "Carta Tipo - IncorporacionesPOR - copia.docx",
        "Inglés": "Carta Tipo - IncorporacionesENG - copia.docx"
    }
    return templates.get(language, None)

def replace_text(doc, replacements):
    for para in doc.paragraphs:
        for key, value in replacements.items():
            if key in para.text:
                para.text = para.text.replace(key, value)

def format_date(fecha, idioma):
    meses = {
        "Español": ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"],
        "Portugués": ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"],
        "Inglés": ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
    }
    mes_texto = meses[idioma][fecha.month - 1]
    return f"{fecha.day} - {mes_texto} - {fecha.year}"

st.title("Generador de Cartas")

idioma = st.selectbox("Seleccione el idioma de la carta", ["Español", "Portugués", "Inglés"])
nombre = st.text_input("Nombre")
localizador = st.text_input("Localizador")
fecha_input = st.date_input("Fecha")
ciudad = st.text_input("Ciudad")
trayecto = st.text_input("Trayecto")
hora_presentacion = st.text_input("Hora de Presentación")
hora_salida = st.text_input("Hora de Salida")
punto_encuentro = st.text_input("Punto de Encuentro")
direccion = st.text_input("Dirección")

if st.button("Generar Carta"):
    template_path = load_template(idioma)
    if not template_path:
        st.error("No se encontró la plantilla para el idioma seleccionado.")
    else:
        doc = docx.Document(template_path)
        fecha_formateada = format_date(fecha_input, idioma)
        dia_texto = fecha_input.strftime("%A")
        replacements = {
            "(NOM)": nombre,
            "(LOC)": localizador,
            "(FECHA)": fecha_formateada,
            "(DIA)": dia_texto,
            "(CIU)": ciudad,
            "(TRAY)": trayecto,
            "(PRES)": hora_presentacion,
            "(SALI)": hora_salida,
            "(ENCU)": punto_encuentro,
            "(DIRE)": direccion
        }
        replace_text(doc, replacements)
        output_filename = f"{localizador}.docx"
        doc.save(output_filename)
        st.success(f"Carta generada exitosamente: {output_filename}")
        with open(output_filename, "rb") as file:
            st.download_button("Descargar Carta", file, file_name=output_filename)
