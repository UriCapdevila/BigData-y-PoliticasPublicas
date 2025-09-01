import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import plotly.express as px
import os

# Crear carpetas si no existen
os.makedirs("img", exist_ok=True)
os.makedirs("html", exist_ok=True)

# Cargar ambas hojas del Excel
archivo = r"C:\Users\Uri_C\OneDrive\Documentos\GitHub\BigData-y-PoliticasPublicas\CSV\estadisticas2025paralumnos.xlsx"
estadisticas = pd.read_excel(archivo, sheet_name="estadistica_fin_de_cursado")
cupos = pd.read_excel(archivo, sheet_name="resultados_desgranamiento")

# Unir las hojas por cod_materia y comision
df = pd.merge(estadisticas, cupos, on=["cod_materia", "comision"], how="inner")

# ================================
# 📉 Gráfico estático: Cupo vs Inscriptos
# ================================
plt.figure(figsize=(16,8))
df_plot = df[["comision", "cupo_maximo", "cantidad_inscriptos"]].melt(
    id_vars="comision",
    value_vars=["cupo_maximo", "cantidad_inscriptos"],
    var_name="Tipo",
    value_name="Cantidad"
)
orden_comisiones = df.groupby("comision")["cantidad_inscriptos"].sum().sort_values(ascending=False).index
sns.barplot(data=df_plot, x="comision", y="Cantidad", hue="Tipo", palette="Paired", order=orden_comisiones)
plt.xticks(rotation=45, ha="right")
plt.title("Comparación entre cupo máximo y cantidad de inscriptos por comisión", fontsize=14)
plt.ylabel("Cantidad de alumnos")
plt.xlabel("Comisión")
plt.legend(title="Tipo de dato")
for container in plt.gca().containers:
    plt.gca().bar_label(container, fmt="%.0f", label_type="edge", fontsize=8)
plt.tight_layout()
plt.savefig("img/grafico_cupo_inscriptos.png", dpi=300, bbox_inches="tight")
plt.close()

# ================================
# 🚪 Gráfico estático: Abandonos por materia
# ================================
plt.figure(figsize=(16,8))
orden_materias = df.groupby("actividad")["abandono"].sum().sort_values(ascending=False).index
sns.barplot(data=df, x="actividad", y="abandono", palette="mako", order=orden_materias)
plt.xticks(rotation=90)
plt.title("Cantidad total de abandonos por materia", fontsize=14)
plt.ylabel("Cantidad de abandonos")
plt.xlabel("Materia")
for container in plt.gca().containers:
    plt.gca().bar_label(container, fmt="%.0f", label_type="edge", fontsize=8)
plt.tight_layout()
plt.savefig("img/grafico_abandonos.png", dpi=300, bbox_inches="tight")
plt.close()

# ================================
# 📊 Gráfico interactivo: Cupo vs Inscriptos
# ================================
fig_cupo = px.bar(
    df_plot,
    x="comision",
    y="Cantidad",
    color="Tipo",
    barmode="group",
    title="Cupo vs. Inscriptos por comisión",
    labels={"Cantidad": "Cantidad de alumnos", "comision": "Comisión"}
)
fig_cupo.update_layout(xaxis_tickangle=-45)
fig_cupo.write_html("html/comision_cupo_vs_inscriptos.html")

# ================================
# 📊 Gráfico interactivo: Abandonos por materia
# ================================
df_abandonos = df.groupby("actividad")["abandono"].sum().reset_index()
df_abandonos = df_abandonos.sort_values(by="abandono", ascending=False)
fig_abandonos = px.bar(
    df_abandonos,
    x="actividad",
    y="abandono",
    title="Cantidad de abandonos por materia",
    labels={"actividad": "Materia", "abandono": "Cantidad de abandonos"}
)
fig_abandonos.update_layout(xaxis_tickangle=-90)
fig_abandonos.write_html("html/abandonos_por_materia.html")