import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import re
import numpy as np
np.set_printoptions(threshold=np.inf)
# -------------------------------------
# Se cargan datos desde Excel
# -------------------------------------
archivo_salida = "Salida Escenarios.xlsx"
df_resumen = pd.read_excel(archivo_salida, sheet_name="Resumen Escenarios")
df_master = pd.read_excel(archivo_salida, sheet_name="Analisis energético")

# -------------------------------------
# Filtrar por elasticidad = -1.87
# -------------------------------------
df_resumen = df_resumen[df_resumen["Elasticidad"] == -1.87].copy()

def extraer_elasticidad(nombre_escenario):
    match = re.search(r"E=([-]?\d+\.?\d*)", str(nombre_escenario))
    return float(match.group(1)) if match else None

df_master["Elasticidad extraída"] = df_master["Escenario"].apply(extraer_elasticidad)

# Mantener los escenarios con E=-1.87 o el Escalonado Escenario 1
df_master = df_master[
    (df_master["Elasticidad extraída"] == -1.87) |
    (df_master["Escenario"] == "Escalonado Escenario 1")
].copy()

df_master.drop(columns=["Elasticidad extraída"], inplace=True)

# Ordenar escenarios numéricamente
def extraer_numero(nombre):
    match = re.search(r'Escenario (\d+)', nombre)
    return int(match.group(1)) if match else -1  # Escalonado queda primero

# Se adecuan los nombres de los escenarios para mejor entendimiento y visualización en el texto de la tesis.
# Comentar el replace si se desea mantener los nombres originales del archiv excel.
mapa_escenarios = {
    "Escalonado Escenario 1": "Escenario 1 - Escalonado",
    "Lineal Escenario 1 (E=-1.87)": "Escenario 1 - Lineal",
    "Lineal Escenario 2 (E=-1.87)": "Escenario 2",
    "Lineal Escenario 3 (E=-1.87)": "Escenario 3",
    "Lineal Escenario 4 (E=-1.87)": "Escenario 4",
    "Lineal Escenario 5 (E=-1.87)": "Escenario 5",
    "Lineal Escenario 6 (E=-1.87)": "Escenario 6",
    "Lineal Escenario 7 (E=-1.87)": "Escenario 7",
    "Lineal Escenario 8 (E=-1.87)": "Escenario 8",
    "Lineal Escenario 9 (E=-1.87)": "Escenario 9",
    "Lineal Escenario 10 (E=-1.87)": "Escenario 10"
}
df_master["Escenario"] = df_master["Escenario"].replace(mapa_escenarios)

orden_escenarios = sorted(df_master["Escenario"].unique(), key=extraer_numero)
df_master["Escenario"] = pd.Categorical(df_master["Escenario"], categories=orden_escenarios, ordered=True)
for esc in df_master["Escenario"].unique():
    print(esc)


# -------------------------------------
# GRÁFICOS
# -------------------------------------

# Gráfico 1: Variación de m³ por Tipo 2
df_litros = df_master.groupby(["Tipo 2", "Escenario"], as_index=False, observed=True)["Variación consumo anual (m3)"].sum()
plt.figure(figsize=(12, 7))
sns.barplot(data=df_litros, x='Tipo 2', y='Variación consumo anual (m3)', hue='Escenario', errorbar=None, palette="husl")
plt.title('Impacto de los escenarios sobre el consumo anual de combustible, diferenciado por categoría vehicular \n Escenarios filtrados por elasticidad = -1,87')
plt.xlabel('Tipo de vehículo')
plt.ylabel('Variación en el consumo de combustible (m³/año)')
plt.xticks(rotation=45)
plt.legend(title='Escenario', bbox_to_anchor=(1.05, 1), loc='upper left')
plt.tight_layout()
plt.savefig(r'C:\Users\emili\PycharmProjects\TesisUY\Gráficos\Variaciónm3porTipo2.png', dpi=300)
#plt.show()

# Gráfico 2: Variación de CO2 por Tipo 2
df_co2 = df_master.groupby(["Tipo 2", "Escenario"], as_index=False, observed=True)["Variación CO2 (ton)"].sum()
plt.figure(figsize=(12, 7))
sns.barplot(data=df_co2, x='Tipo 2', y='Variación CO2 (ton)', hue='Escenario', errorbar=None, palette="husl")
plt.title('Impacto de los escenarios sobre las emisiones de CO₂, diferenciado por categoría vehicular \n Escenarios filtrados por elasticidad = -1,87')
plt.xlabel('Tipo de vehículo')
plt.ylabel('Variación en las emisiones de CO₂ (ton/año)')
plt.xticks(rotation=45)
plt.legend(title='Escenario', bbox_to_anchor=(1.05, 1), loc='upper left')
plt.tight_layout()
plt.savefig(r'C:\Users\emili\PycharmProjects\TesisUY\Gráficos\VariaciónCO2porTipo2.png', dpi=300)
#plt.show()

# Gráfico 3: m³ por Tipo de motor
df_filtro = df_master[~df_master["Tipo de motor"].isin(["BEV"])].copy()

df_grouped = df_filtro.groupby(["Escenario", "Tipo de motor"], as_index=False, observed=True).agg({
    "Variación consumo anual (m3)": "sum",
    "Variación CO2 (ton)": "sum"
})

pivot_m3 = df_grouped.pivot(index='Escenario', columns='Tipo de motor', values='Variación consumo anual (m3)')
pivot_m3 = pivot_m3.reindex(index=orden_escenarios)

pivot_m3.plot(kind='bar', stacked=True, figsize=(12, 7))
plt.title('Impacto de los escenarios sobre el consumo anual de combustible, diferenciado por tipo de motor \n Escenarios filtrados por elasticidad = -1,87')
plt.ylabel('Variación en el consumo de combustible (m³/año)')
plt.xlabel("Escenario")
plt.xticks(rotation=45)
plt.legend(title="Tipo de motor", bbox_to_anchor=(1.05, 1), loc='upper left')
plt.tight_layout()
plt.savefig(r'C:\Users\emili\PycharmProjects\TesisUY\Gráficos\Variaciónm3porTipodemotor.png', dpi=300)
#plt.show()

# Gráfico 4: CO₂ por Tipo de motor (se mantiene BEV)
pivot_co2 = df_grouped.pivot(index='Escenario', columns='Tipo de motor', values='Variación CO2 (ton)')
pivot_co2 = pivot_co2.reindex(index=orden_escenarios)

pivot_co2.plot(kind='bar', stacked=True, figsize=(12, 7))
plt.title('Impacto de los escenarios sobre las emisiones de CO₂, diferenciado por tipo de motor \n Escenarios filtrados por elasticidad = -1,87')
plt.ylabel('Variación en las emisiones de CO₂ (ton/año)')
plt.xlabel("Escenario")
plt.xticks(rotation=45)
plt.legend(title="Tipo de motor", bbox_to_anchor=(1.05, 1), loc='upper left')
plt.tight_layout()
plt.savefig(r'C:\Users\emili\PycharmProjects\TesisUY\Gráficos\VariaciónCO2porTipodemotor.png', dpi=300)
#plt.show()
