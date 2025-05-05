import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import re

# -------------------------------------
# Cargar datos desde Excel
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

#Mantener los escenarios con E=-1.87 o el Escalonado Escenario 1
df_master = df_master[
    (df_master["Elasticidad extraída"] == -1.87) |
    (df_master["Escenario"] == "Escalonado Escenario 1")
].copy()

df_master.drop(columns=["Elasticidad extraída"], inplace=True)

# Ordenar escenarios numéricamente
def extraer_numero(nombre):
    match = re.search(r'Escenario (\d+)', nombre)
    return int(match.group(1)) if match else -1  # Escalonado queda primero

orden_escenarios = sorted(df_master["Escenario"].unique(), key=extraer_numero)
df_master["Escenario"] = pd.Categorical(df_master["Escenario"], categories=orden_escenarios, ordered=True)

# -------------------------------------
# GRÁFICOS
# -------------------------------------

# Gráfico 1: Variación de m³ por Tipo 2
df_litros = df_master.groupby(["Tipo 2", "Escenario"], as_index=False, observed=True)["Variación consumo anual (m3)"].sum()
plt.figure(figsize=(12, 7))
sns.barplot(data=df_litros, x='Tipo 2', y='Variación consumo anual (m3)', hue='Escenario', errorbar=None, palette="husl")
plt.title('Variación de m³ Consumidos por Tipo 2 y Escenario')
plt.xlabel('Tipo de Vehículo')
plt.ylabel('Variación Anual (m³)')
plt.xticks(rotation=45)
plt.legend(title='Escenario', bbox_to_anchor=(1.05, 1), loc='upper left')
plt.tight_layout()
plt.savefig(r'C:\Users\emili\PycharmProjects\TesisUY\Gráficos\Variaciónm3porTipo2.png', dpi=300)
#plt.show()

# Gráfico 2: Variación de CO2 por Tipo 2
df_co2 = df_master.groupby(["Tipo 2", "Escenario"], as_index=False, observed=True)["Variación CO2 (ton)"].sum()
plt.figure(figsize=(12, 7))
sns.barplot(data=df_co2, x='Tipo 2', y='Variación CO2 (ton)', hue='Escenario', errorbar=None, palette="husl")
plt.title('Variación de Emisiones de CO₂ por Tipo 2 y Escenario')
plt.xlabel('Tipo de Vehículo')
plt.ylabel('Variación CO₂ (toneladas)')
plt.xticks(rotation=45)
plt.legend(title='Escenario', bbox_to_anchor=(1.05, 1), loc='upper left')
plt.tight_layout()
plt.savefig(r'C:\Users\emili\PycharmProjects\TesisUY\Gráficos\VariaciónCO2porTipo2.png', dpi=300)
#plt.show()

# Gráfico 3: m³ por Tipo de motor
df_filtro = df_master.copy()

df_grouped = df_filtro.groupby(["Escenario", "Tipo de motor"], as_index=False, observed=True).agg({
    "Variación consumo anual (m3)": "sum",
    "Variación CO2 (ton)": "sum"
})

pivot_m3 = df_grouped.pivot(index='Escenario', columns='Tipo de motor', values='Variación consumo anual (m3)')
pivot_m3 = pivot_m3.reindex(index=orden_escenarios)

pivot_m3.plot(kind='bar', stacked=True, figsize=(12, 7))
plt.title('Variación de m³ Consumidos por Escenario y Tipo de Motor')
plt.ylabel('Variación anual (m³)')
plt.xlabel("Escenario")
plt.xticks(rotation=45)
plt.legend(title="Tipo de motor", bbox_to_anchor=(1.05, 1), loc='upper left')
plt.tight_layout()
plt.savefig(r'C:\Users\emili\PycharmProjects\TesisUY\Gráficos\Variaciónm3porTipodemotor.png', dpi=300)
#plt.show()

# Gráfico 4: CO₂ por Tipo de motor
pivot_co2 = df_grouped.pivot(index='Escenario', columns='Tipo de motor', values='Variación CO2 (ton)')
pivot_co2 = pivot_co2.reindex(index=orden_escenarios)

pivot_co2.plot(kind='bar', stacked=True, figsize=(12, 7))
plt.title('Variación de Emisiones de CO₂ por Escenario y Tipo de Motor')
plt.ylabel('Variación CO₂ (toneladas)')
plt.xlabel("Escenario")
plt.xticks(rotation=45)
plt.legend(title="Tipo de motor", bbox_to_anchor=(1.05, 1), loc='upper left')
plt.tight_layout()
plt.savefig(r'C:\Users\emili\PycharmProjects\TesisUY\Gráficos\VariaciónCO2porTipodemotor.png', dpi=300)
#plt.show()
