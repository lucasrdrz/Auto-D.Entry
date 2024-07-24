import streamlit as st
import pandas as pd
import qrcode
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import cv2
from pyzbar.pyzbar import decode
import numpy as np
from googleapiclient.discovery import build
from google.oauth2 import service_account

pip.main(["install", "openpyxl"])

# Función para generar el código QR
def generate_qr(data):
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=2,  # Reducir el tamaño de la caja
        border=1,    # Reducir el tamaño del borde
    )
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    return img

# Función para leer el archivo de Excel, generar QR y guardarlo en la celda F36
def process_file(file, sheet_name='backup', usecols="A:D", nrows=28):
    df = pd.read_excel(file, sheet_name=sheet_name, usecols=usecols, nrows=nrows)
    df = df.fillna('')
    data = df.to_string(index=False)
    img = generate_qr(data)

    # Cargar el archivo de Excel usando openpyxl
    wb = load_workbook(file)
    ws = wb[sheet_name]

    # Guardar el QR en un archivo temporal para insertarlo en el Excel
    qr_image_stream = BytesIO()
    img.save(qr_image_stream, format='PNG')
    qr_image_stream.seek(0)

    # Guardar una copia de los datos de la imagen QR antes de cerrar el flujo
    qr_image_data = qr_image_stream.getvalue()
    
    img = Image(BytesIO(qr_image_data))

    # Insertar el QR en la celda F36
    ws.add_image(img, 'F36')

    # Guardar el archivo modificado en un BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return output, qr_image_data, df

# Función para decodificar el QR de una imagen
def decode_qr(image_data):
    file_bytes = np.asarray(bytearray(image_data), dtype=np.uint8)
    img = cv2.imdecode(file_bytes, cv2.IMREAD_COLOR)
    decoded_objects = decode(img)
    results = []
    for obj in decoded_objects:
        results.append(obj.data.decode('utf-8'))
    return results

# Configurar las credenciales y el servicio de la API de Google Sheets
SERVICE_ACCOUNT_FILE = './key.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service = build('sheets', 'v4', credentials=credentials)
SPREADSHEET_ID = '1uC3qyYAmThXMfJ9Pwkompbf9Zs6MWhuTqT8jTVLYdr0'

def get_last_row(sheet_id, sheet_name, column='B'):
    try:
        # Leer los valores de la columna especificada
        result = service.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range=f"'{sheet_name}'!{column}:{column}"
        ).execute()
        values = result.get('values', [])
        return len(values) + 1
    except Exception as e:
        st.write(f"Error al obtener la última fila: {e}")
        return None

# Configurar la aplicación de Streamlit
st.title('Generador y Lector de Código QR para Excel')
st.write('Arrastra y suelta tu archivo Excel a continuación:')

uploaded_file = st.file_uploader('Sube tu archivo Excel', type=['xlsx'])

if uploaded_file is not None:
    sheet_name = st.text_input('Nombre de la hoja', value='backup')
    usecols = st.text_input('Columnas a leer (ej. A:D)', value='A:D')
    nrows = st.number_input('Número de filas a leer', min_value=1, value=28)

    if st.button('Generar QR y Mostrar Información'):
        processed_file, qr_image_data, df = process_file(uploaded_file, sheet_name, usecols, nrows)

        # Guardar el DataFrame en session_state
        st.session_state.df = df

        # Decodificar y mostrar la información del QR
        results = decode_qr(qr_image_data)
        if results:
            st.write('Datos del QR decodificados:')
            for result in results:
                st.write(result)
        else:
            st.write('No se encontró ningún QR en la imagen.')

        st.download_button(
            label="Descargar archivo modificado",
            data=processed_file,
            file_name="archivo_modificado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # Agregar un botón para cargar datos en Google Sheet
    if 'df' in st.session_state and st.button('Cargar datos en Google Sheet'):
        df = st.session_state.df

        # Extraer los valores específicos del DataFrame
        ticket_2 = df['TICKET'][2]  # Valor en la posición 2 PAQ
        tecnico_5 = df['DESCRIPCION'][5]  # Valor en la posición 5 DESCRIPCION

        for i in range(9, 23):
            numero_de_parte = df['NUMERO DE PARTE'][i] if i < len(df['NUMERO DE PARTE']) else ""
            ticket_9 = df['TICKET'][i] if i < len(df['TICKET']) else ""
            cantidad = df['CANT.'][i] if i < len(df['CANT.']) else ""

            # Verificar si todos los valores están completos
            if numero_de_parte and ticket_9 and cantidad:
                # Obtener la última fila con datos y comenzar desde ahí
                last_row = get_last_row(SPREADSHEET_ID, "Sheet1")  # Cambiar a "Sheet1"

                if last_row:
                    ranges = {
                        'NUMERO DE PARTE': f'B{last_row}',  # Ajusta la celda según sea necesario
                        'TICKET_9': f'H{last_row}',         # Ajusta la celda según sea necesario
                        'CANTIDAD': f'I{last_row}',         # Ajusta la celda según sea necesario
                        'TICKET_2': f'G{last_row}',         # Ajusta la celda según sea necesario
                        'TECNICO': f'E{last_row}'           # Ajusta la celda según sea necesario
                    }

                    # Preparar los valores en el formato adecuado
                    values = [
                        [numero_de_parte],
                        [ticket_9],
                        [cantidad],
                        [ticket_2],
                        [tecnico_5]
                    ]

                    # Llamada a la API para insertar los valores
                    for value, range_name in zip(values, ranges.values()):
                        result = service.spreadsheets().values().update(
                            spreadsheetId=SPREADSHEET_ID,
                            range=range_name,
                            valueInputOption='USER_ENTERED',
                            body={'values': [value]}
                        ).execute()
                        st.write(f"Datos insertados correctamente en {range_name}. {result.get('updatedCells')} celdas actualizadas.")
            else:
                st.write(f"Fila {i + 1}: datos incompletos, omitiendo.")

