import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from googleapiclient.discovery import build
from google.oauth2 import service_account
import os
import json

# Función para leer el archivo de Excel y cargarlo en un DataFrame
def process_file(file, sheet_name='backup', usecols="A:D", nrows=28):
    df = pd.read_excel(file, sheet_name=sheet_name, usecols=usecols, nrows=nrows)
    df = df.fillna('')

    # Cargar el archivo de Excel usando openpyxl
    wb = load_workbook(file)
    ws = wb[sheet_name]

    # Guardar el archivo modificado en un BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return output, df

# Configurar las credenciales y el servicio de la API de Google Sheets
def load_credentials():
    try:
        SERVICE_ACCOUNT_INFO = os.getenv('GCP_KEY_JSON')
        info = json.loads(SERVICE_ACCOUNT_INFO)
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        credentials = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
        return build('sheets', 'v4', credentials=credentials)
    except Exception as e:
        st.error(f"Error al configurar las credenciales: {e}")
        st.stop()

service = load_credentials()

SPREADSHEET_ID = '1uC3qyYAmThXMfJ9Pwkompbf9Zs6MWhuTqT8jTVLYdr0'

def get_last_row(sheet_id, sheet_name, column='B'):
    try:
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
st.title('Cargar y Mostrar Información de Excel')
st.write('Arrastra y suelta tu archivo Excel a continuación:')

uploaded_file = st.file_uploader('Sube tu archivo Excel', type=['xlsx'])

if uploaded_file is not None:
    sheet_name = st.text_input('Nombre de la hoja', value='backup')
    usecols = st.text_input('Columnas a leer (ej. A:D)', value='A:D')
    nrows = st.number_input('Número de filas a leer', min_value=1, value=28)

    if st.button('Mostrar Información'):
        processed_file, df = process_file(uploaded_file, sheet_name, usecols, nrows)

        st.session_state.df = df

        st.write('Datos del archivo Excel:')
        st.dataframe(df)

        st.download_button(
            label="Descargar archivo modificado",
            data=processed_file,
            file_name="archivo_modificado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    if 'df' in st.session_state and st.button('Cargar datos en Google Sheet'):
        df = st.session_state.df

        ticket_2 = df['TICKET'][2]
        tecnico_5 = df['DESCRIPCION'][5]

        for i in range(9, 23):
            numero_de_parte = df['NUMERO DE PARTE'][i] if i < len(df['NUMERO DE PARTE']) else ""
            ticket_9 = df['TICKET'][i] if i < len(df['TICKET']) else ""
            cantidad = df['CANT.'][i] if i < len(df['CANT.']) else ""

            if numero_de_parte and ticket_9 and cantidad:
                last_row = get_last_row(SPREADSHEET_ID, "Sheet1")

                if last_row:
                    ranges = {
                        'NUMERO DE PARTE': f'B{last_row}',
                        'TICKET_9': f'H{last_row}',
                        'CANTIDAD': f'I{last_row}',
                        'TICKET_2': f'G{last_row}',
                        'TECNICO': f'E{last_row}'
                    }

                    values = [
                        [numero_de_parte],
                        [ticket_9],
                        [cantidad],
                        [ticket_2],
                        [tecnico_5]
                    ]

                    for value, range_name in zip(values, ranges.values()):
                        try:
                            result = service.spreadsheets().values().update(
                                spreadsheetId=SPREADSHEET_ID,
                                range=range_name,
                                valueInputOption='USER_ENTERED',
                                body={'values': [value]}
                            ).execute()
                            st.write(f"Datos insertados correctamente en {range_name}. {result.get('updatedCells')} celdas actualizadas.")
                        except Exception as e:
                            st.write(f"Error al insertar los datos en {range_name}: {e}")
            else:
                st.write(f"Fila {i + 1}: datos incompletos, omitiendo.")




