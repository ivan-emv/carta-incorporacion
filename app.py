import streamlit as st
import docx
from docx.shared import Pt
from datetime import datetime
import shutil

# Diccionario de traducción de meses
date_translation = {
    "es": {"January": "Enero", "February": "Febrero", "March": "Marzo", "April": "Abril", "May": "Mayo", "June": "Junio", "July": "Julio", "August": "Agosto", "September": "Septiembre", "October": "Octubre", "November": "Noviembre", "December": "Diciembre"},
    "pt": {"January": "Janeiro", "February": "Fevereiro", "March": "Março", "April": "Abril", "May": "Maio", "June": "Junho", "July": "Julho", "August": "Agosto", "September": "Setembro", "October": "Outubro", "November": "Novembro", "December": "Dezembro"},
    "en": {"January": "January", "February": "February", "March": "March", "April": "April", "May": "May", "June": "June", "July": "July", "August": "August", "September": "September", "October": "October", "November": "November", "December": "December"}
}

# Selección del idioma
st.title("Generador de Cartas de Incorporación")
idioma = st.selectbox("Seleccione el idioma de la carta", ["Español", "Portugués", "Inglés"])
idioma_cod = {"Español": "es", "Portugués": "pt", "Inglés": "en"}[idioma]

doc_files = {
    "es": "Carta Tipo - Incorporaciones.docx",
    "pt": "Carta Tipo - IncorporacionesPOR.docx",
    "en": "Carta Tipo - IncorporacionesENG.docx"
}

doc_file = doc_files[idioma_cod]

# Campos del formulario
nombre = st.text_input("Nombre")
localizador = st.text_input("Localizador")
fecha_input = st.text_input("Fecha (DD/MM/YYYY)")
ciudad = st.text_input("Ciudad")
trayecto = st.text_input("Trayecto")
hora_presentacion = st.text_input("Hora de Presentación")
hora_salida = st.text_input("Hora de Salida")
punto_encuentro = st.text_input("Punto de Encuentro")
direccion = st.text_input("Dirección")

if st.button("Generar Carta"):
    try:
        # Validar formato de fecha
        fecha_obj = datetime.strptime(fecha_input, "%d/%m/%Y")
        mes = fecha_obj.strftime("%B")
        mes_traducido = date_translation[idioma_cod][mes]
        fecha_formateada = f"{fecha_obj.day} - {mes_traducido} - {fecha_obj.year}"
        dia_semana = fecha_obj.strftime("%A")
        
        # Copiar documento base para no modificar el original
        output_filename = f"{localizador}.docx"
        shutil.copy(doc_file, output_filename)
        doc = docx.Document(output_filename)
        
        # Reemplazar variables sin afectar formato ni imágenes
        for para in doc.paragraphs:
            for run in para.runs:
                run.text = run.text.replace("(INSERTENOMBRE)", nombre)
                run.text = run.text.replace("(LOCALIZADOR)", localizador)
                run.text = run.text.replace("(INSERTEFECHA)", fecha_formateada)
                run.text = run.text.replace("(DIA)", dia_semana)
                run.text = run.text.replace("(CIUDAD)", ciudad)
                run.text = run.text.replace("(INSERTETRAYECTO)", trayecto)
                run.text = run.text.replace("(HORAPRESENTACION)", hora_presentacion)
                run.text = run.text.replace("(HORASALIDA)", hora_salida)
                run.text = run.text.replace("(PUNTOENCUENTRO)", punto_encuentro)
                run.text = run.text.replace("(INSERTEDIRECCION)", direccion)
        
        # Guardar documento con modificaciones
        doc.save(output_filename)
        
        # Permitir descarga del archivo
        with open(output_filename, "rb") as f:
            st.download_button("Descargar Carta", f, file_name=output_filename)
    except ValueError:
        st.error("Formato de fecha incorrecto. Use DD/MM/YYYY.")
