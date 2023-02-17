import pyodbc
import pandas as pd
import PySimpleGUI as sg
import calendar

conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=\\servidor\comun\FELJA.mdb;UID=administrador;PWD=F3lj42008!'
conn = pyodbc.connect(conn_str)

layout = [
    [sg.Text('Selecciona el rango de fechas')],
    [sg.CalendarButton('Desde', target='dateDesde', key='desde'), sg.Input(key='dateDesde', size=(10,1)),
     sg.CalendarButton('Hasta', target='dateHasta', key='hasta'), sg.Input(key='dateHasta', size=(10,1))],
    [sg.Button('Generar reporte')],
]

window = sg.Window('Reporte de asistencia', layout)

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED:
        break
    if event == 'Generar reporte':
        fecha_desde = pd.to_datetime(values['dateDesde']).strftime('%m/%d/%Y')
        fecha_hasta = pd.to_datetime(values['dateHasta']).strftime('%m/%d/%Y')
        
        query = f"""
            SELECT horaspersonal_id, terceros_autonum, FORMAT(horaspersonal_hora, 'Short Date') as Fecha, FORMAT(horaspersonal_hora, 'Short Time') as [Date and Time], TIPO, SENSOR, horaspersonal_check, RAZON
            FROM horas_personal_exportar
            WHERE horaspersonal_hora BETWEEN #{fecha_desde}# AND #{fecha_hasta}#
        """
        df = pd.read_sql_query(query, conn)

        # Renombrar la columna "terceros_autonum" a "ID Card"
        df = df.rename(columns={"terceros_autonum": "ID Card"})

        # Cambiar el formato de la fecha en la columna "Fecha"
        df["Fecha"] = pd.to_datetime(df["Fecha"], dayfirst=True, format="%d/%m/%Y").dt.strftime("%d/%m/%Y")
        df["Fecha"] = pd.to_datetime(df["Fecha"]).dt.strftime("%d/%m/%Y")
        df['Date and Time'] = pd.to_datetime(df['Date and Time'], format='%H:%M')

        df["Fecha"] = pd.to_datetime(df["Fecha"], dayfirst=True, format="%d/%m/%Y").dt.strftime("%d/%m/%Y")
        df["Fecha"] = pd.to_datetime(df["Fecha"]).dt.strftime("%d/%m/%Y")

        # Agregar el campo "Nombre" con los nombres de los días en inglés
        df["Day"] = pd.to_datetime(df["Fecha"], format="%d/%m/%Y").dt.day_name()

        # Extraer la hora de la columna "Date and Time" y separarla en dos nuevas columnas "Time In" y "Time Out"
        df['Fecha'] = pd.to_datetime(df['Fecha'], format='%d/%m/%Y')
        df_grouped = df.groupby(['ID Card', 'Fecha']).agg({'Date and Time': ['min', 'max']})
        df_grouped.columns = ['Time In', 'Time Out']
        df_grouped = df_grouped.reset_index()
        df = pd.merge(df, df_grouped, on=['ID Card', 'Fecha'])

        # Eliminar la columna "Date and Time" y renombrar la columna "Fecha" a "Date"
        df = df.drop(columns=["Date and Time"])
        df = df.rename(columns={"Fecha": "Date", "Time In": "Time In", "Time Out": "Time Out"})
        df["Time In"] = pd.to_datetime(df["Time In"], format='%I:%M %p').dt.strftime("%H:%M:%S")
        df["Time Out"] = pd.to_datetime(df["Time Out"], format='%I:%M %p').dt.strftime("%H:%M:%S")

        # Crear un diccionario para convertir los nombres de los días en inglés a español
        days_dict = {calendar.day_name[i]: dia for i, dia in enumerate(["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"])}

        # Cambiar el nombre de los días en la columna "Nombre" a español
        df["Day"] = df["Day"].map(days_dict)

        # Crear una lista de nombres de columna para exportar al archivo de Excel
        column_names = ["ID Card", "Day", "Date", "Time In", "Time Out"]

        # Solicitar al usuario un nombre y ubicación para el archivo de Excel generado
        filename = sg.popup_get_file("Guardar como...", save_as=True, file_types=(("Archivos de Excel", "*.xlsx"),))

        if filename:
            with pd.ExcelWriter(filename) as writer:
                # Escribir los datos procesados en el archivo de Excel
                df[column_names].to_excel(writer, index=False)

            sg.popup(f"Archivo guardado en {filename}")

conn.close()
window.close()