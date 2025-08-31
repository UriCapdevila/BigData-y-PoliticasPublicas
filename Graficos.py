import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import os

# Crear carpeta para guardar imágenes
os.makedirs("img", exist_ok=True)

# Cargar ambas hojas del Excel
archivo = r"C:\Users\Uri_C\OneDrive\Documentos\GitHub\BigData-y-PoliticasPublicas\CSV\estadisticas2025paralumnos.xlsx"

estadisticas = pd.read_excel(archivo, sheet_name="estadistica_fin_de_cursado")
cupos = pd.read_excel(archivo, sheet_name="resultados_desgranamiento")

# Unir las hojas por cod_materia y comision
df = pd.merge(estadisticas, cupos, on=["cod_materia", "comision"], how="inner")

# Mostrar información básica
print(df.head())
print(df.columns)
print(df.info())
print(df.describe())

# Crear columna de tasa de aprobación
df["tasa_aprobacion"] = df["promociono"] / df["cantidad_inscriptos"]

# 📊 Gráfico de tasa de aprobación por comisión
plt.figure(figsize=(12,6))
sns.barplot(data=df, x="comision", y="tasa_aprobacion", palette="viridis")
plt.xticks(rotation=45)
plt.title("Tasa de aprobación por comisión")
plt.ylabel("Tasa de aprobación")
plt.xlabel("Comisión")
plt.tight_layout()
plt.savefig("img/grafico_aprobacion.png", dpi=300, bbox_inches="tight")
plt.show()

# 📉 Gráfico comparativo de cupo vs inscriptos
plt.figure(figsize=(12,6))
df_plot = df[["comision", "cupo_maximo", "cantidad_inscriptos"]].melt(
    id_vars="comision",
    value_vars=["cupo_maximo", "cantidad_inscriptos"],
    var_name="Tipo",
    value_name="Cantidad"
)
sns.barplot(data=df_plot, x="comision", y="Cantidad", hue="Tipo", palette="Set2")
plt.xticks(rotation=45)
plt.title("Cupo vs. Inscriptos por comisión")
plt.tight_layout()
plt.savefig("img/grafico_cupo_inscriptos.png", dpi=300, bbox_inches="tight")
plt.show()

# 🚪 Gráfico de abandonos por materia
plt.figure(figsize=(12,6))
sns.barplot(data=df, x="cod_materia", y="abandono", palette="rocket")
plt.xticks(rotation=45)
plt.title("Cantidad de abandonos por materia")
plt.ylabel("Abandonos")
plt.xlabel("Materia")
plt.tight_layout()
plt.savefig("img/grafico_abandonos.png", dpi=300, bbox_inches="tight")
plt.show()

# 🏆 Ranking de materias por tasa de promoción
df["tasa_promocion"] = df["promociono"] / df["cantidad_inscriptos"]
ranking = df.groupby("cod_materia")["tasa_promocion"].mean().sort_values(ascending=False)

plt.figure(figsize=(10,8))
sns.barplot(x=ranking.values, y=ranking.index, palette="coolwarm")
plt.title("Ranking de materias por tasa de promoción")
plt.xlabel("Tasa de promoción promedio")
plt.ylabel("Materia")
plt.tight_layout()
plt.savefig("img/grafico_ranking.png", dpi=300, bbox_inches="tight")
plt.show()