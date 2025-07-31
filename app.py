import streamlit as st
import requests
import re
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import os

# === CONFIGURACION ===
CARPETA_ID_DRIVE = "1BjHmxl7eIaR1c1WtQVC25tc9TtKe1vrh"

# === FUNCIONES ===
def convertir_fecha_hora(fecha_hora_str):
    meses = {
        "January": "01", "February": "02", "March": "03", "April": "04",
        "May": "05", "June": "06", "July": "07", "August": "08",
        "September": "09", "October": "10", "November": "11", "December": "12"
    }
    match = re.match(r"(\d{1,2}) de ([a-zA-Z]+) de (\d{4}) en horas: (\d{2}:\d{2}:\d{2})", fecha_hora_str)
    if match:
        dia, mes, anio, hora = match.groups()
        mes_num = meses.get(mes, "00")
        return f"{dia.zfill(2)}/{mes_num}/{anio} {hora}"
    return fecha_hora_str

def crear_servicio_drive():
    info = dict(st.secrets["gdrive"])
    info["private_key"] = info["private_key"].replace("\\n", "\n")
    creds = service_account.Credentials.from_service_account_info(
        info,
        scopes=["https://www.googleapis.com/auth/drive"]
    )
    return build('drive', 'v3', credentials=creds)

def subir_a_drive(servicio, archivo_local, nombre_en_drive, folder_id):
    metadata = {'name': nombre_en_drive, 'parents': [folder_id]}
    media = MediaFileUpload(archivo_local, resumable=True)
    archivo = servicio.files().create(body=metadata, media_body=media, fields='id').execute()
    return archivo['id']

# === INTERFAZ ===
st.title("📸 Generar y Subir Reporte de Selfies desde SIGOF")

usuario = st.text_input("👤 Usuario")
clave = st.text_input("🔐 Clave", type="password")

if st.button("🔓 Iniciar sesión y generar reporte"):
    with st.spinner("Procesando..."):
        login_url = "http://sigof.distriluz.com.pe/plus/usuario/login"
        data_url = "http://sigof.distriluz.com.pe/plus/ComlecOrdenlecturas/ajax_mostar_mapa_selfie"

        credentials = {
            "data[Usuario][usuario]": usuario,
            "data[Usuario][pass]": clave
        }

        headers = {
            "User-Agent": "Mozilla/5.0",
            "Referer": login_url,
        }

        with requests.Session() as session:
            login_response = session.post(login_url, data=credentials, headers=headers)
            if "Usuario o contraseña incorrecto" in login_response.text:
                st.error("❌ Usuario o clave incorrectos")
                st.stop()

            data_response = session.get(data_url, headers=headers)

        data = re.sub(r"<\/?\w+.*?>", "", data_response.text.replace("\\/", "/"))
        blocks = re.split(r"Ver detalle", data)

        results = {}
        for block in blocks:
            fecha = re.search(r"Fecha Selfie:\s*(\d{1,2} de [a-zA-Z]+ de \d{4} en horas: \d{2}:\d{2}:\d{2})", block)
            lecturista = re.search(r"Lecturista:\s*([\w\sÁÉÍÓÚáéíóúÑñ]+)", block)
            url = re.search(r"url\":\"(https[^\"]+)", block)

            if fecha and lecturista and url:
                fecha_formateada = convertir_fecha_hora(fecha.group(1).strip())
                fecha_solo, _ = fecha_formateada.split(" ")
                nombre = lecturista.group(1).strip()
                url_img = url.group(1).strip()
                key = (nombre, fecha_solo)
                results.setdefault(key, {"URLs Imagen": []})["URLs Imagen"].append(url_img)

        if not results:
            st.warning("⚠️ No se encontraron selfies para estas credenciales.")
            st.stop()

        max_urls = max(len(i["URLs Imagen"]) for i in results.values())
        url_cols = [f"Url_foto {i+1}" for i in range(max_urls)]
        vista_cols = [f"Vista Url_foto {i+1}" for i in range(max_urls)]
        all_cols = ["Fecha Selfie", "Lecturista"] + url_cols + vista_cols

        wb = Workbook()
        ws = wb.active
        ws.title = "LmcSelfiesLectura"
        ws.append(all_cols)

        for i, ((lecturista, fecha), info) in enumerate(results.items(), start=2):
            row = [fecha, lecturista] + info["URLs Imagen"] + [""] * (max_urls - len(info["URLs Imagen"]))
            ws.append(row + [""] * max_urls)
            for j in range(max_urls):
                url_cell = f"{get_column_letter(3 + j)}{i}"
                vista_cell = f"{get_column_letter(3 + max_urls + j)}{i}"
                ws[vista_cell] = f'=IMAGE({url_cell},4,200,140)'
                ws.row_dimensions[i].height = 151

        for cell in ws[1]:
            cell.fill = PatternFill(start_color="007BFF", fill_type="solid")
            cell.font = Font(color="FFFFFF", bold=True)
            cell.alignment = Alignment(horizontal='center')

        filename = "Lmc_ReporteSelfie.xlsx"
        wb.save(filename)

        servicio = crear_servicio_drive()
        subir_a_drive(servicio, filename, filename, CARPETA_ID_DRIVE)
        enlace = f"https://drive.google.com/drive/folders/{CARPETA_ID_DRIVE}"

        st.success("✅ Reporte generado y subido a Drive.")
        st.markdown(
            f'<a href="{enlace}" target="_blank">'
            f'<button style="background:#007BFF;color:white;padding:10px;border:none;border-radius:5px">📁 Abrir carpeta en Drive</button>'
            f'</a>', unsafe_allow_html=True)

        os.remove(filename)
