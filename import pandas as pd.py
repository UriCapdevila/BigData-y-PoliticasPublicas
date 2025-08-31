import pandas as pd

# Cargar ambas hojas del Excel
archivo = r"C:\Users\Uri_C\OneDrive\Documentos\GitHub\BigData-y-PoliticasPublicas\CSV\estadisticas2025paralumnos.xlsx"  #reemplazá con la ruta real

estadisticas = pd.read_excel(archivo, sheet_name="estadistica_fin_de_cursado")  #nombre real de la hoja
cupos = pd.read_excel(archivo, sheet_name="resultados_desgranamiento")  #nombre real de la hoja

df = pd.merge(estadisticas, cupos, on=["cod_materia", "comision"], how="inner")