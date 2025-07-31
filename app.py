import streamlit as st
import os
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# === CONFIGURACIÓN ===
NOMBRE_JSON = "lmcselfies.json"
CARPETA_ID_DRIVE = "1BjHmxl7eIaR1c1WtQVC25tc9TtKe1vrh"  # ID real de tu carpeta

# === AUTENTICACIÓN ===
def crear_servicio_drive():
    with open(NOMBRE_JSON) as fuente:
        info = json.load(fuente)
    creds = service_account.Credentials.from_service_account_info(
        info,
        scopes=["https://www.googleapis.com/auth/drive"]
    )
    servicio = build('drive', 'v3', credentials=creds)
    return servicio

# === SUBIR ARCHIVO A DRIVE ===
def subir_archivo_a_drive(servicio, archivo_local, nombre_destino, folder_id):
    archivo_metadata = {
        'name': nombre_destino,
        'parents': [folder_id]
    }
    media = MediaFileUpload(archivo_local, resumable=True)
    archivo = servicio.files().create(
        body=archivo_metadata,
        media_body=media,
        fields='id'
    ).execute()
    return archivo.get('id')

# === INTERFAZ DE USUARIO ===
st.title("📤 Subir Reporte de Selfies a Google Drive")

archivo = st.file_uploader("Selecciona un archivo Excel", type=["xlsx", "xls"])

if archivo is not None:
    with open(archivo.name, "wb") as f:
        f.write(archivo.read())

    if st.button("⬆️ Subir a Drive"):
        with st.spinner("Autenticando y subiendo..."):
            servicio = crear_servicio_drive()
            archivo_id = subir_archivo_a_drive(servicio, archivo.name, archivo.name, CARPETA_ID_DRIVE)

        st.success("✅ Archivo subido correctamente.")
        enlace = f"https://drive.google.com/drive/folders/{CARPETA_ID_DRIVE}"
        st.markdown(
            f'<a href="{enlace}" target="_blank">'
            f'<button style="background-color:#007BFF; color:white; padding:10px; border:none; border-radius:5px;">'
            f'📁 Abrir carpeta en Google Drive'
            f'</button></a>', unsafe_allow_html=True
        )

        os.remove(archivo.name)
