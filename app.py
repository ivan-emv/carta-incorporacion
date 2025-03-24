import streamlit as st
from datetime import datetime
from docx import Document
import os

# Función para cargar el archivo y modificar los campos
def modificar_carta(archivo, nombre, localizador, fecha, ciudad, trayecto, hora_presentacion, hora_salida, punto_encuentro, direccion):
    doc = Document(archivo)
    
    # Reemplazar campos
    for parrafo in doc.paragraphs:
        parrafo.text = parrafo.text.replace('(INSERTENOMBRE)', nombre)
        parrafo.text = parrafo.text.replace('(LOCALIZADOR)', localizador)
        parrafo.text = parrafo.text.replace('(INSERTEFECHA)', fecha)
        parrafo.text = parrafo.text.replace('(CIUDAD)', ciudad)
        parrafo.text = parrafo.text.replace('(INSERTETRAYECTO)', trayecto)
        parrafo.text = parrafo.text.replace('(HORAPRESENTACION)', hora_presentacion)
        parrafo.text = parrafo.text.replace('(HORASALIDA)', hora_salida)
        parrafo.text = parrafo.text.replace('(PUNTOENCUENTRO)', punto_encuentro)
        parrafo.text = parrafo.text.replace('(INSERTEDIRECCION)', direccion)
    
    # Guardar documento modificado
    nombre_archivo = f"{localizador}.docx"
    doc.save(nombre_archivo)
    return nombre_archivo

# Interfaz de usuario en Streamlit
def main():
    st.title("Generador de Carta Tipo")

    # Selección de idioma
    idioma = st.selectbox("Selecciona el idioma", ("Español", "Portugués", "Inglés"))
    
    # Archivos correspondientes a cada idioma
    if idioma == "Español":
        archivo = "Carta Tipo - Incorporaciones.docx"
    elif idioma == "Portugués":
        archivo = "Carta Tipo - IncorporacionesPOR.docx"
    else:
        archivo = "Carta Tipo - IncorporacionesENG.docx"

    # Cargar el archivo
    archivo_path = os.path.join('path/to/your/files', archivo)

    # Campos de entrada
    nombre = st.text_input("Nombre")
    localizador = st.text_input("Localizador")
    fecha = st.text_input("Fecha (DD/MM/AAAA)", "")
    ciudad = st.text_input("Ciudad")
    trayecto = st.text_input("Trayecto")
    hora_presentacion = st.text_input("Hora de Presentación")
    hora_salida = st.text_input("Hora de Salida")
    punto_encuentro = st.text_input("Punto de Encuentro")
    direccion = st.text_input("Dirección")

    # Validación de la fecha
    try:
        fecha_obj = datetime.strptime(fecha, "%d/%m/%Y")
        dia = fecha_obj.strftime("%A")
    except ValueError:
        dia = None

    if dia is None:
        st.error("La fecha ingresada no es válida. Debe estar en formato DD/MM/AAAA.")
    else:
        st.write(f"El día correspondiente es: {dia}")

    # Botón para generar la carta
    if st.button("Generar Carta"):
        if all([nombre, localizador, fecha, ciudad, trayecto, hora_presentacion, hora_salida, punto_encuentro, direccion]):
            archivo_generado = modificar_carta(archivo_path, nombre, localizador, fecha, ciudad, trayecto, hora_presentacion, hora_salida, punto_encuentro, direccion)
            st.success(f"La carta ha sido generada con éxito. Puedes descargarla [aquí](/{archivo_generado}).")
        else:
            st.error("Por favor, completa todos los campos.")

if __name__ == "__main__":
    main()
