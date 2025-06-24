import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import re

# ------------------------
# Configuración
# ------------------------
archivo_resultados = "Salida Escenarios.xlsx"
resumen_df = pd.read_excel(archivo_resultados, sheet_name="Resumen Escenarios")

# ------------------------
# Extraer y filtrar por elasticidad = -1.87
# ------------------------
def extraer_elasticidad(nombre):
    match = re.search(r"E=([-]?\d+\.?\d*)", str(nombre))
    return float(match.group(1)) if match else None

resumen_df["Elasticidad"] = resumen_df["Escenario"].apply(extraer_elasticidad)
resumen_df.loc[resumen_df["Escenario"] == "Escalonado Escenario 1", "Elasticidad"] = -1.87
# Filtrar por elasticidad
resumen_df = resumen_df[resumen_df["Elasticidad"] == -1.87].copy()

# ------------------------
# Renombrar escenarios y definir orden. Se realiza a efectos de mejorar visualización en gráficos
# ------------------------
mapa_escenarios = {
    "Escalonado Escenario 1":       "Escenario 1 - Escalonado",
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
resumen_df["Escenario"] = resumen_df["Escenario"].replace(mapa_escenarios)
orden = list(mapa_escenarios.values())
resumen_df["Escenario"] = pd.Categorical(resumen_df["Escenario"], categories=orden, ordered=True)

# ------------------------
# Cálculo de ventas y diferencias
# ------------------------
resumen_df["Ventas Base Total"] = (resumen_df["Ventas Año Base (Sin BEV)"]
                                       + resumen_df["Ventas BEV Año Base"])
resumen_df["Ventas Escenario Total"] = (resumen_df["Ventas Escenario (Sin BEV)"]
                                       + resumen_df["Ventas BEV Escenario"])

resumen_df["Diff Ventas Sin BEV"] = (
    resumen_df["Ventas Escenario (Sin BEV)"] - resumen_df["Ventas Año Base (Sin BEV)"]
)
resumen_df["Diff Ventas BEV"] = (
    resumen_df["Ventas BEV Escenario"] - resumen_df["Ventas BEV Año Base"]
)
resumen_df["Diff Ventas Total"]  = (
    resumen_df["Ventas Escenario Total"] - resumen_df["Ventas Base Total"]
)

# ------------------------
# Gráfico 1: Diferencia de recaudación IMESI (en millones USD)
# ------------------------
fig, ax = plt.subplots(figsize=(12, 6))
bars = ax.bar(
    resumen_df["Escenario"],
    resumen_df["Diferencia Recaudación IMESI (USD)"],
    edgecolor="black"
)

# Formatear eje Y: mostrar en millones y miles separados por puntos
def millones_fmt(x, pos):
    m = x / 1e6
    s = f"{m:,.1f}"
    # intercambiar separador miles por '.', decimal por ','
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return s

ax.yaxis.set_major_formatter(mtick.FuncFormatter(millones_fmt))

ax.set_ylabel("Δ Recaudación IMESI (millones USD)")
ax.set_title("Variación de la recaudación de IMESI del escenario con respecto al año base \n Escenarios filtrados por elasticidad = -1,87")
plt.xticks(rotation=45, ha="right")
ax.grid(axis="y", linestyle="--", alpha=0.6)
plt.tight_layout()
plt.savefig("C:/Users/emili/PycharmProjects/TesisUY/Gráficos/Variación de ventas/Variación en racudación IMESI.png", dpi=300)
plt.show()
# ------------------------
# Gráfico 2: Diferencia de ventas por tipo
# ------------------------
x = np.arange(len(resumen_df))
w = 0.4

plt.figure(figsize=(12, 6))
plt.bar(x - w/2,
        resumen_df["Diff Ventas Sin BEV"],
        width=w,
        label="Δ Ventas No BEV",
        edgecolor="black")
plt.bar(x + w/2,
        resumen_df["Diff Ventas BEV"],
        width=w,
        label="Δ Ventas BEV",
        edgecolor="black")
plt.xticks(x, resumen_df["Escenario"], rotation=45, ha="right")
plt.ylabel("Δ Ventas (unidades)")
plt.title("Diferencia de ventas entre el escenario y el año base, diferenciado entre BEV y no BEV \n Escenarios filtrados por elasticidad = -1,87")
plt.legend()
plt.grid(axis="y", linestyle="--", alpha=0.6)
plt.tight_layout()
plt.savefig("C:/Users/emili/PycharmProjects/TesisUY/Gráficos/Variación de ventas/Variación ventas separado por BEV y no BEV.png", dpi=300)
plt.show()

# ------------------------
# Calculo de recaudación por unidad en escenario y base
# ------------------------
resumen_df["Reca_per_unit_esc"] = (
    resumen_df["Recaudación IMESI Escenario BEV (USD)"]
  + resumen_df["Recaudación IMESI Escenario (Sin BEV) (USD)"]
) / (
    resumen_df["Ventas BEV Escenario"]
  + resumen_df["Ventas Escenario (Sin BEV)"]
)

resumen_df["Reca_per_unit_base"] = (
    resumen_df["Recaudación IMESI Año Base BEV (USD)"]
  + resumen_df["Recaudación IMESI Año Base (Sin BEV) (USD)"]
) / (
    resumen_df["Ventas BEV Año Base"]
  + resumen_df["Ventas Año Base (Sin BEV)"]
)

# ------------------------
# Diferencia en recaudación por unidad
# ------------------------
resumen_df["Diff_Reca_per_unit"] = (
    resumen_df["Reca_per_unit_esc"] - resumen_df["Reca_per_unit_base"]
)

# ------------------------
# Gráfico 3: Variación en la recaudación de IMESI por unidad vendida, del escenario con respecto al año base
# ------------------------
fig, ax = plt.subplots(figsize=(12,6))
ax.bar(
    resumen_df["Escenario"],
    resumen_df["Diff_Reca_per_unit"],
    edgecolor="black"
)
ax.set_ylabel("Δ Recaudación IMESI por unidad (USD)")
ax.set_title("Variación en la recaudación de IMESI por unidad vendida, del escenario con respecto al año base \n Escenarios filtrados por elasticidad = -1,87")
ax.axhline(0, color='gray', linewidth=0.8)
ax.grid(axis='y', linestyle='--', alpha=0.6)
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.savefig("C:/Users/emili/PycharmProjects/TesisUY/Gráficos/Variación de ventas/Variación recaudación IMESI por unidad.png", dpi=300)
plt.show()






