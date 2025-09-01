import pandas as pd
import plotly.express as px
import os

# 📁 Crear carpeta para guardar los gráficos interactivos
os.makedirs("html", exist_ok=True)

# 📥 Cargar ambas hojas del Excel
archivo = r"C:\Users\Uri_C\OneDrive\Documentos\GitHub\BigData-y-PoliticasPublicas\CSV\estadisticas2025paralumnos.xlsx"
estadisticas = pd.read_excel(archivo, sheet_name="estadistica_fin_de_cursado")
cupos = pd.read_excel(archivo, sheet_name="resultados_desgranamiento")

# 🔗 Unir las hojas por cod_materia y comision
df = pd.merge(estadisticas, cupos, on=["cod_materia", "comision"], how="inner")

# ============================================================
# 📊 GRÁFICO INTERACTIVO 1: Cupo vs Inscriptos por comisión
# ============================================================

# 🔄 Reorganizar datos para mostrar cupo e inscriptos como categorías
df_plot = df[["comision", "cupo_maximo", "cantidad_inscriptos"]].melt(
    id_vars="comision",                      # Mantener la columna 'comision' como identificador
    value_vars=["cupo_maximo", "cantidad_inscriptos"],  # Variables que se comparan
    var_name="Tipo",                         # Nueva columna que indica si es cupo o inscriptos
    value_name="Cantidad"                   # Valores numéricos
)

# 📈 Crear gráfico de barras agrupadas
fig_cupo = px.bar(
    df_plot,
    x="comision",                            # Eje X: comisiones
    y="Cantidad",                            # Eje Y: cantidad de alumnos
    color="Tipo",                            # Color por tipo (cupo o inscriptos)
    barmode="group",                         # Barras agrupadas
    title="Cupo vs. Inscriptos por comisión",# Título del gráfico
    labels={"Cantidad": "Cantidad de alumnos", "comision": "Comisión"}  # Etiquetas personalizadas
)

# 🎨 Ajustes visuales del gráfico
fig_cupo.update_layout(
    xaxis_tickangle=-45                     # Rotar etiquetas del eje X para mejor lectura
)

# 💾 Guardar gráfico como archivo HTML
fig_cupo.write_html("html/comision_cupo_vs_inscriptos.html")

# ============================================================
# 📊 GRÁFICO INTERACTIVO 2: Abandonos por materia (barras verticales)
# ============================================================

# 📊 Agrupar y ordenar datos por cantidad de abandonos
df_abandonos = df.groupby("actividad")["abandono"].sum().reset_index()
df_abandonos = df_abandonos.sort_values(by="abandono", ascending=False)

# 📈 Crear gráfico de barras verticales
fig_abandonos = px.bar(
    df_abandonos,
    x="actividad",                          # Eje X: materias
    y="abandono",                           # Eje Y: cantidad de abandonos
    title="Cantidad de abandonos por materia",
    labels={"actividad": "Materia", "abandono": "Cantidad de abandonos"},
    color="abandono",                       # Color según cantidad de abandonos
    color_continuous_scale="Reds"           # Escala de color: de claro a oscuro
)

# 🎨 Ajustes visuales del gráfico
fig_abandonos.update_layout(
    xaxis_tickangle=-45,                    # Rotar etiquetas del eje X
    xaxis_title="Materia",                  # Título del eje X
    yaxis_title="Cantidad de abandonos",    # Título del eje Y
    font=dict(size=10),                     # Tamaño general de fuente
    xaxis=dict(tickfont=dict(size=9)),      # Tamaño de etiquetas del eje X
    yaxis=dict(tickfont=dict(size=9)),      # Tamaño de etiquetas del eje Y
    bargap=0.02,                             # Espacio entre grupos de barras (más chico = barras más gruesas)
    bargroupgap=0.04,                        # Espacio entre barras dentro del mismo grupo
    margin=dict(t=60, b=120)                # Margen superior e inferior para evitar solapamiento
)

# 💾 Guardar gráfico como archivo HTML
fig_abandonos.write_html("html/abandonos_por_materia.html")