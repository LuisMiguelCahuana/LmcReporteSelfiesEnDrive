# app.py
import streamlit as st
import requests
import re
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import os

# Título
st.markdown(
    "<h3 style='text-align: center; color: white; background-color: #007BFF; padding: 10px;'>"
    "🔐 HUMANO, INGRESA TUS CREDENCIALES DE SIGOF WEB"
    "</h3>", unsafe_allow_html=True)

# Formulario de credenciales
usuario = st.text_input("👤 Usuario", placeholder="Ingrese su usuario")
clave = st.text_input("🔑 Clave", placeholder="Ingrese su contraseña", type="password")

if st.button("🔓 Iniciar sesión"):
    if not usuario or not clave:
        st.warning("⚠️ Debes ingresar usuario y clave.")
    else:
        with st.spinner("Iniciando sesión..."):
            login_url = "http://sigof.distriluz.com.pe/plus/usuario/login"
            data_url = "http://sigof.distriluz.com.pe/plus/ComlecOrdenlecturas/ajax_mostar_mapa_selfie"

            session = requests.Session()
            credentials = {
                "data[Usuario][usuario]": usuario,
                "data[Usuario][pass]": clave
            }
            headers = {
                "User-Agent": "Mozilla/5.0",
                "Referer": login_url,
            }

            login_response = session.post(login_url, data=credentials, headers=headers)

            if "Usuario o contraseña incorrecto" in login_response.text:
                st.error("🧠 Credenciales incorrectas. Intente nuevamente.")
            else:
                data_response = session.get(data_url, headers=headers)
                data = data_response.text.replace("\\/", "/")
                data = re.sub(r"<\/?\w+.*?>", "", data)
                data = re.sub(r"\s+", " ", data).strip()
                blocks = re.split(r"Ver detalle", data)

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

                results = {}
                for block in blocks:
                    fecha = re.search(r"Fecha Selfie:\s*(\d{1,2} de [a-zA-Z]+ de \d{4} en horas: \d{2}:\d{2}:\d{2})", block)
                    lecturista = re.search(r"Lecturista:\s*([\w\sÁÉÍÓÚáéíóúÑñ]+)", block)
                    url = re.search(r"url\":\"(https[^\"]+)", block)

                    if fecha and lecturista and url:
                        fecha_hora = convertir_fecha_hora(fecha.group(1).strip())
                        fecha_solo, _ = fecha_hora.split(" ")
                        nombre = lecturista.group(1).strip()
                        imagen_url = url.group(1).strip()
                        key = (nombre, fecha_solo)
                        if key not in results:
                            results[key] = {"URLs Imagen": []}
                        results[key]["URLs Imagen"].append(imagen_url)

                if results:
                    max_urls = max(len(i["URLs Imagen"]) for i in results.values())
                    url_columns = [f"Url_foto {i+1}" for i in range(max_urls)]
                    vista_columns = [f"Vista Url_foto {i+1}" for i in range(max_urls)]
                    all_columns = ["Fecha Selfie", "Lecturista"] + url_columns + vista_columns

                    wb = Workbook()
                    ws = wb.active
                    ws.title = "LmcSelfiesLectura"
                    ws.freeze_panes = "A2"
                    ws.append(all_columns)

                    ws.column_dimensions['A'].width = 12
                    ws.column_dimensions['B'].width = 18

                    header_fill = PatternFill(start_color="007BFF", end_color="007BFF", fill_type="solid")
                    header_font = Font(color="FFFFFF", bold=True)

                    for cell in ws[1]:
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = Alignment(horizontal='center')

                    for i, ((lecturista, fecha_selfie), info) in enumerate(results.items(), start=2):
                        row = [fecha_selfie, lecturista] + info["URLs Imagen"] + [""] * (max_urls - len(info["URLs Imagen"]))
                        ws.append(row + [""] * max_urls)
                        ws[f"A{i}"].alignment = Alignment(horizontal='left')
                        ws[f"B{i}"].alignment = Alignment(horizontal='justify')
                        for j in range(max_urls):
                            url_cell = f"{get_column_letter(3 + j)}{i}"
                            formula_cell = f"{get_column_letter(3 + max_urls + j)}{i}"
                            ws[formula_cell] = f'=IMAGE({url_cell};;3;200;140)'
                            ws.column_dimensions[get_column_letter(3 + max_urls + j)].width = round(140 / 7, 1)
                        ws.row_dimensions[i].height = 151

                    filename = "Lmc_ReporteSelfie.xlsx"
                    wb.save(filename)

                    with open(filename, "rb") as f:
                        st.success("✅ Reporte generado con éxito. Descárgalo aquí:")
                        st.download_button(
                            label="📥 Descargar Lmc_ReporteSelfie.xlsx",
                            data=f,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    os.remove(filename)
                else:
                    st.warning("⚠️ No se encontraron datos o las credenciales no tienen selfies.")

