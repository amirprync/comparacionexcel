import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# Subir archivo Excel
uploaded_file = st.file_uploader("Carga tu archivo de Excel con dos hojas", type=['xlsx'])

if uploaded_file:
    # Leer el archivo Excel y las hojas
    df_hoja1 = pd.read_excel(uploaded_file, sheet_name='Hoja1', engine='openpyxl')
    df_hoja2 = pd.read_excel(uploaded_file, sheet_name='Hoja2', engine='openpyxl')

    # Mostrar el contenido de las hojas
    st.write("Contenido de Hoja1:")
    st.dataframe(df_hoja1)
    st.write("Contenido de Hoja2:")
    st.dataframe(df_hoja2)

    # Convertir las fechas a un formato comparable si no lo están ya
    df_hoja1['FechaConcertacion'] = pd.to_datetime(df_hoja1['FechaConcertacion'], errors='coerce')
    df_hoja2['FechaConcertacion'] = pd.to_datetime(df_hoja2['FechaConcertacion'], errors='coerce')

    # Encontrar las coincidencias
    coincidencias = pd.merge(df_hoja1, df_hoja2, on=['Comitente', 'FechaConcertacion'])

    # Cargar el archivo con openpyxl para poder marcar los registros
    wb = load_workbook(uploaded_file)
    ws_hoja2 = wb['Hoja2']

    # Resaltar las filas coincidentes en la Hoja2
    fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    for index, row in coincidencias.iterrows():
        for i in range(2, ws_hoja2.max_row + 1):  # Asumiendo que hay un encabezado
            if (ws_hoja2.cell(row=i, column=1).value == row['Comitente'] and 
                ws_hoja2.cell(row=i, column=2).value == row['FechaConcertacion']):
                for j in range(1, ws_hoja2.max_column + 1):
                    ws_hoja2.cell(row=i, column=j).fill = fill

    # Guardar el archivo con las marcas
    output_filename = '/mnt/data/resultado.xlsx'
    wb.save(output_filename)

    # Mostrar el enlace de descarga
    with open(output_filename, 'rb') as f:
        st.download_button('Descargar archivo con registros marcados', f, file_name='resultado.xlsx')
