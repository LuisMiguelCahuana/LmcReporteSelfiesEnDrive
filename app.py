import streamlit as st
import requests
import re
import pandas as pd
import os
import xlsxwriter

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

                    filename = "Lmc_ReporteSelfie.xlsx"
                    workbook = xlsxwriter.Workbook(filename)
                    worksheet = workbook.add_worksheet("LmcSelfiesLectura")

                    header_format = workbook.add_format({
                        'bold': True,
                        'font_color': 'white',
                        'bg_color': '#007BFF',
                        'align': 'center'
                    })

                    # Escribir cabeceras
                    for col_num, col_name in enumerate(all_columns):
                        worksheet.write(0, col_num, col_name, header_format)
                        worksheet.set_column(col_num, col_num, 22)

                    # Escribir filas
                    for row_num, ((lecturista, fecha_selfie), info) in enumerate(results.items(), start=1):
                        worksheet.write(row_num, 0, fecha_selfie)
                        worksheet.write(row_num, 1, lecturista)

                        for j, url in enumerate(info["URLs Imagen"]):
                            url_col = 2 + j
                            img_col = 2 + max_urls + j
                            worksheet.write(row_num, url_col, url)
                            formula = f'=IMAGEN({xlsxwriter.utility.xl_col_to_name(url_col)}{row_num + 1};;3;200;140)'
                            worksheet.write_formula(row_num, img_col, formula)
                            worksheet.set_column(img_col, img_col, 20)
                        worksheet.set_row(row_num, 140)

                    workbook.close()

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
