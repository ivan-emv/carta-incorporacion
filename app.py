import streamlit as st
from docx import Document
from io import BytesIO
import zipfile
import shutil
import os
from datetime import datetime

# Diccionario para asociar idioma con plantilla
PLANTILLAS = {
    "Español": "Carta Tipo - Incorporaciones.docx",
    "Portugués": "Carta Tipo - IncorporacionesPOR.docx",
    "Inglés": "Carta Tipo - IncorporacionesENG.docx",
}

# Función para modificar XML directamente sin alterar el formato
def reemplazar_texto_xml(docx_path, reemplazos):
    temp_dir = "temp_docx"
    temp_docx = "temp_modified.docx"
    
    # Extraer el contenido del .docx (que es un archivo .zip)
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)
    
    # Buscar y modificar TODOS los archivos XML en el documento (para asegurar el reemplazo en todas las secciones)
    for root, _, files in os.walk(temp_dir):
        for file in files:
            if file.endswith(".xml"):
                xml_path = os.path.join(root, file)
                with open(xml_path, "r", encoding="utf-8") as f:
                    xml_content = f.read()
                
                # Reemplazar los textos asegurando que no se corrompa el XML
                for key, value in reemplazos.items():
                    xml_content = xml_content.replace(key, value)
                
                # Escribir el archivo modificado
                with open(xml_path, "w", encoding="utf-8") as f:
                    f.write(xml_content)
    
    # Volver a comprimir el .docx modificado
    with zipfile.ZipFile(temp_docx, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for foldername, subfolders, filenames in os.walk(temp_dir):
            for filename in filenames:
                file_path = os.path.join(foldername, filename)
                zipf.write(file_path, os.path.relpath(file_path, temp_dir))
    
    # Limpiar archivos temporales
    shutil.rmtree(temp_dir)
    
    return temp_docx

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
punto_encuentro = st.text_area("Inserte Punto de Encuentro")
direccion = st.text_input("Inserte Dirección")

# Validación de fecha y obtención del día y mes en texto
try:
    fecha_obj = datetime.strptime(fecha_input, "%d/%m/%Y")
    dia_semana = fecha_obj.strftime("%A")  # Día en inglés
    dia_num = fecha_obj.strftime("%d")
    mes_num = fecha_obj.strftime("%m")
    anio = fecha_obj.strftime("%Y")
    
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
    
    mes_traducido = {
        "01": "Enero", "02": "Febrero", "03": "Marzo", "04": "Abril", "05": "Mayo", "06": "Junio",
        "07": "Julio", "08": "Agosto", "09": "Septiembre", "10": "Octubre", "11": "Noviembre", "12": "Diciembre"
    }
    
    fecha_formateada = f"{dia_num} - {mes_traducido[mes_num]} - {anio}"
    fecha_valida = True
except ValueError:
    st.error("Formato de fecha inválido. Use el formato DD/MM/YYYY.")
    fecha_valida = False

# Reemplazos si la fecha es válida
if fecha_valida and st.button("Generar Documento"):
    reemplazos = {
        "(INSERTENOMBRE)": nombre,
        "(LOCALIZADOR)": localizador,
        "(INSERTEFECHA)": fecha_formateada,
        "(DIA)": dias_traducidos[idioma][dia_semana],
        "(CIUDAD)": ciudad,
        "(INSERTETRAYECTO)": trayecto,
        "(HORAPRESENTACION)": hora_presentacion,
        "(HORASALIDA)": hora_salida,
        "(PUNTODEENCUENTRO)": punto_encuentro,
        "(INSERTEDIRECCION)": direccion
    }
    
    plantilla = PLANTILLAS[idioma]
    docx_modificado = reemplazar_texto_xml(plantilla, reemplazos)
    
    with open(docx_modificado, "rb") as file:
        st.download_button(
            label="Descargar Documento",
            data=file,
            file_name=f"{localizador}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    
    os.remove(docx_modificado)
