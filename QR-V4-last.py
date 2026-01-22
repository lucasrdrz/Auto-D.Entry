import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from googleapiclient.discovery import build
from google.oauth2 import service_account
import json

# =========================
# Procesar archivo Excel
# =========================
def process_file(file, sheet_name='backup', usecols="A:D", nrows=28):
    file.seek(0)  # 🔴 volver al inicio
    file.seek(0)
    df = pd.read_excel(file, sheet_name=sheet_name, nrows=nrows)
    df = df.iloc[:, 0:4]  # en vez de usecols="A:D"
    df = df.fillna('')

    # Crear un Excel limpio desde el DataFrame (NO desde el original)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)

    output.seek(0)
    return output, df

# =========================
# Google Sheets credentials
# =========================
def load_credentials():
    try:
        SERVICE_ACCOUNT_INFO = st.secrets["GCP_KEY_JSON"]
        info = json.loads(SERVICE_ACCOUNT_INFO)
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        credentials = service_account.Credentials.from_service_account_info(
            info, scopes=SCOPES
        )
        return build('sheets', 'v4', credentials=credentials)
    except Exception as e:
        st.error(f"Error al configurar las credenciales: {e}")
        st.stop()

service = load_credentials()

SPREADSHEET_ID = '1uC3qyYAmThXMfJ9Pwkompbf9Zs6MWhuTqT8jTVLYdr0'

# =========================
# Obtener última fila
# =========================
def get_last_row(sheet_id, sheet_name, column='B'):
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range=f"'{sheet_name}'!{column}:{column}"
        ).execute()
        values = result.get('values', [])
        return len(values) + 1
    except Exception as e:
        st.error(f"Error al obtener la última fila: {e}")
        return None

# =========================
# Streamlit UI
# =========================
st.title('Cargar y Mostrar Información de Excel')

uploaded_file = st.file_uploader(
    'Arrastra y suelta tu archivo Excel', type=['xlsx']
)

if uploaded_file is not None:
    sheet_name = st.text_input('Nombre de la hoja', value='backup')
    usecols = st.text_input('Columnas a leer (ej. A:D)', value='A:D')
    nrows = st.number_input('Número de filas a leer', min_value=1, value=28)

    if st.button('Mostrar Información'):
        processed_file, df = process_file(
            uploaded_file, sheet_name, usecols, nrows
        )
        st.session_state.df = df

        st.subheader('Datos del archivo Excel')
        st.dataframe(df)

        st.download_button(
            label="Descargar archivo",
            data=processed_file,
            file_name="archivo_modificado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # =========================
    # Cargar datos en Google Sheets
    # =========================
    if 'df' in st.session_state and st.button('Cargar datos en Google Sheet'):
        df = st.session_state.df

        ticket_2 = df['TICKET'][2]
        tecnico_5 = df['DESCRIPCION'][5]
        nombre_carga = df['CANT.'][1]

        # 🔹 CONCATENACIÓN CORRECTA DESDE EXCEL
        banco_26 = str(df['DESCRIPCION'][26]).strip()
        descripcion_27 = str(df['DESCRIPCION'][27]).strip()

        descripcion_49 = " | ".join(
            x for x in [banco_26, descripcion_27] if x
        )

        # =========================
        # Carga fila por fila
        # =========================
        for i in range(9, 23):
            numero_de_parte = df['NUMERO DE PARTE'][i] if i < len(df) else ""
            ticket_9 = df['TICKET'][i] if i < len(df) else ""
            cantidad = df['CANT.'][i] if i < len(df) else ""

            if numero_de_parte and ticket_9 and cantidad:
                last_row = get_last_row(SPREADSHEET_ID, "Sheet1")

                ranges = {
                    'NUMERO DE PARTE': f'B{last_row}',
                    'TICKET_9': f'H{last_row}',
                    'CANTIDAD': f'I{last_row}',
                    'TICKET_2': f'G{last_row}',
                    'TECNICO': f'E{last_row}',
                    'DESCRIPCION_CONCAT': f'J{last_row}',
                    'Nombre_Carga': f'K{last_row}'
                }

                values = [
                    numero_de_parte,
                    ticket_9,
                    cantidad,
                    ticket_2,
                    tecnico_5,
                    descripcion_49,
                    nombre_carga
                ]

                for value, cell in zip(values, ranges.values()):
                    service.spreadsheets().values().update(
                        spreadsheetId=SPREADSHEET_ID,
                        range=cell,
                        valueInputOption='USER_ENTERED',
                        body={'values': [[value]]}
                    ).execute()
            else:
                st.write(f"Fila {i + 1}: datos incompletos, omitida.")








