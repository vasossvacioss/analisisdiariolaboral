import tkinter as tk
from tkinter import filedialog
import pandas as pd
import xlsxwriter

def cargar_excel():
    file_path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx')])
    if file_path:
        df_prueba = pd.read_excel(file_path)
        df_prueba = df_prueba.rename(columns={
            'Anomalía': 'Tipo de Anomalía',
            'Fecha Medición': 'Fecha de Anomalía',
            'Asignado A': 'Asignado a'
        })

        
        df_prueba['Fecha de Anomalía'] = pd.to_datetime(df_prueba['Fecha de Anomalía'])
        conteo_anomalias_tipo_fecha = df_prueba.groupby(['Tipo de Anomalía', df_prueba['Fecha de Anomalía'].dt.year, df_prueba['Fecha de Anomalía'].dt.month]).size().unstack(level=[1, 2], fill_value=0)
        conteo_anomalias_asignado_tipo_fecha = df_prueba.groupby(['Asignado a', 'Tipo de Anomalía', df_prueba['Fecha de Anomalía'].dt.year, df_prueba['Fecha de Anomalía'].dt.month]).size().unstack(level=[2, 3], fill_value=0)
        conteo_usuarios = df_prueba['Asignado a'].value_counts()
        conteo_leidos = df_prueba[df_prueba['Estado'] == 'Leido']['Asignado a'].value_counts()
        efectividad = pd.DataFrame({
            'Asignado a': conteo_usuarios.index,
            'Cantidad Asignada': conteo_usuarios.values,
            'Leidos': conteo_leidos.values,
            'Efectividad': (conteo_leidos / conteo_usuarios * 100)
        })

        conteo_leidos_radio = df_prueba[df_prueba['Estado'] == 'Leido']['Radio'].value_counts()
        cantidad_asignada_radio = df_prueba.groupby('Radio').size()
        conteo_leidos_radio = conteo_leidos_radio.reindex(cantidad_asignada_radio.index, fill_value=0)
        efectividad_radio = pd.DataFrame({
            'Radio': conteo_leidos_radio.index,
            'Leidos': conteo_leidos_radio.values,
            'Cantidad Asignada': cantidad_asignada_radio.values,
            'Efectividad': (conteo_leidos_radio / cantidad_asignada_radio * 100)
        })

        
        fechas_minimas = df_prueba.groupby('Radio')['Fecha de Anomalía'].min()
        fechas_maximas = df_prueba.groupby('Radio')['Fecha de Anomalía'].max()
        tiempos_transcurridos = fechas_maximas - fechas_minimas
        tiempos_transcurridos_str = tiempos_transcurridos.apply(lambda x: f'{x.days} días {x.seconds // 3600} horas {x.seconds % 3600 // 60} minutos')

        
        resumen_tiempo = pd.DataFrame({
            'Radio': tiempos_transcurridos.index,
            'Fecha Minima': fechas_minimas.values,
            'Fecha Maxima': fechas_maximas.values,
            'Tiempo Transcurrido': tiempos_transcurridos_str.values
        })

        
        excel_output = 'resultado.xlsx'

        with pd.ExcelWriter(excel_output, engine='xlsxwriter') as writer:
            conteo_anomalias_tipo_fecha.to_excel(writer, sheet_name='Tipo y Fecha')
            conteo_anomalias_asignado_tipo_fecha.to_excel(writer, sheet_name='Asignado y Tipo')
            conteo_leidos_radio_mes = df_prueba[df_prueba['Estado'] == 'Leido'].groupby(['Radio', df_prueba['Fecha de Anomalía'].dt.month])['Estado'].count().unstack(level=1, fill_value=0)
            cantidad_asignada_radio_mes = df_prueba.groupby(['Radio', df_prueba['Fecha de Anomalía'].dt.month])['Estado'].count().unstack(level=1, fill_value=0)
            efectividad_radio_mes = conteo_leidos_radio_mes.div(cantidad_asignada_radio_mes) * 100
            efectividad_radio_mes['Cantidad Asignada'] = cantidad_asignada_radio_mes
            efectividad_radio_mes.to_excel(writer, sheet_name='Efectividad por Mes')
            efectividad.to_excel(writer, sheet_name='Efectividad', index=False)
            efectividad_radio.to_excel(writer, sheet_name='Leidos por Radio', index=False)
            resumen_tiempo.to_excel(writer, sheet_name='Resumen Tiempo', index=False)

            
            workbook = writer.book
            worksheet = writer.sheets['Resumen Tiempo']
            tiempo_format = workbook.add_format({'num_format': '0 "días" hh "horas" mm "minutos"'})
            worksheet.set_column('D:D', None, tiempo_format)

        print("El archivo Excel ha sido guardado con éxito.")

# ventana 
window = tk.Tk()

# Tamaño de la ventana
window.title("Cargar Excel")
window.geometry("300x100")

# botón de carga de archivo
btn_cargar = tk.Button(window, text="Cargar Excel", command=cargar_excel)
btn_cargar.pack(pady=20)

window.mainloop()
