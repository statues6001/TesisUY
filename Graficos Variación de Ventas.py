import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import re
import matplotlib.ticker as mtick


# ------------------------
# Configuración
# ------------------------
archivo_resultados = "Salida Escenarios.xlsx"
carpeta_salida = "C:/Users/emili/PycharmProjects/TesisUY/Gráficos/Variación de ventas/"

# ------------------------
# Leer datos
# ------------------------
resumen_df = pd.read_excel(archivo_resultados, sheet_name="Resumen Escenarios")
energetico_df = pd.read_excel(archivo_resultados, sheet_name="Analisis energético")

# ------------------------
# Filtrado por elasticidad = -1.87 o Escalonado Escenario 1
# ------------------------
def extraer_elasticidad(nombre_escenario):
    match = re.search(r"E=([-]?\d+\.?\d*)", str(nombre_escenario))
    return float(match.group(1)) if match else None

energetico_df["Elasticidad"] = energetico_df["Escenario"].apply(extraer_elasticidad)
energetico_df = energetico_df[
    (energetico_df["Elasticidad"] == -1.87) |
    (energetico_df["Escenario"] == "Escalonado Escenario 1")
].copy()

resumen_df = resumen_df[resumen_df["Elasticidad"] == -1.87].copy()

# ------------------------
# Ordenar escenarios correctamente
# ------------------------
def extraer_numero(nombre):
    match = re.search(r'Escenario (\d+)', nombre)
    return int(match.group(1)) if match else -1  # Escalonado primero
# Excluir diésel
energetico_df = energetico_df[energetico_df["Tipo de motor"] != "D"].copy()

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

# Reemplazar nombres en ambos DataFrames
energetico_df["Escenario"] = energetico_df["Escenario"].replace(mapa_escenarios)
resumen_df["Escenario"] = resumen_df["Escenario"].replace(mapa_escenarios)

# Redefinir orden luego del reemplazo
orden_escenarios = list(mapa_escenarios.values())
energetico_df["Escenario"] = pd.Categorical(energetico_df["Escenario"], categories=orden_escenarios, ordered=True)
resumen_df["Escenario"] = pd.Categorical(resumen_df["Escenario"], categories=orden_escenarios, ordered=True)


# ------------------------
# Gráfico 1: Comparación de ventas totales
# ------------------------
resumen_df['Ventas Año Base'] = resumen_df['Ventas Año Base (Sin BEV)'] + resumen_df['Ventas BEV Año Base']
resumen_df['Ventas Escenario'] = resumen_df['Ventas Escenario (Sin BEV)'] + resumen_df['Ventas BEV Escenario']

plt.figure(figsize=(12, 7))
ax = resumen_df.set_index('Escenario')[['Ventas Año Base', 'Ventas Escenario']].plot(kind='bar', ax=plt.gca())

plt.ylabel('Ventas de vehículos')
plt.title('Ventas totales por escenario en relación al año sin intervención \n Escenarios filtrados por elasticidad = -1,87')
plt.xticks(rotation=45, ha='right')

# Límite inferior del eje y
plt.ylim(45000, None)

# Formato del eje y con separador de miles
ax.yaxis.set_major_formatter(mtick.FuncFormatter(lambda x, _: f'{int(x):,}'.replace(',', '.')))

# Mover leyenda fuera del gráfico (abajo)
plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=2)

plt.grid(axis='y', linestyle='--')
plt.tight_layout()
plt.savefig(carpeta_salida + "Variación ventas totales.png", dpi=300)
plt.show()

# ------------------------
# Gráfico 2: Variación porcentual por tipo de motor
# ------------------------
ventas_agg = energetico_df.groupby(['Tipo de motor', 'Escenario'], observed=False).agg({
    'Ventas Escenario': 'sum',
    'Ventas Año Base': 'sum'
}).reset_index()

ventas_agg['Variacion pct'] = (ventas_agg['Ventas Escenario'] - ventas_agg['Ventas Año Base']) / ventas_agg['Ventas Año Base']

# Luego graficás:
plt.figure(figsize=(12, 6))
sns.barplot(
    data=ventas_agg,
    x="Tipo de motor",
    y="Variacion pct",
    hue="Escenario",
    errorbar=None
)

plt.ylabel('Variación porcentual de ventas')
plt.gca().yaxis.set_major_formatter(mtick.PercentFormatter(xmax=1, decimals=1))
plt.title('Variación porcentual en las unidades vendidas para cada escenario, diferenciado por tipo de motor \n Escenarios filtrados por elasticidad = -1,87')
plt.axhline(0, color='gray', linewidth=0.8)
plt.xticks(rotation=45)
plt.legend(title='Escenario', bbox_to_anchor=(1.05, 1), loc='upper left')
plt.tight_layout()
plt.savefig(carpeta_salida + "Variacion ventas por tipo motor.png", dpi=300)
plt.show()

# ------------------------
# Gráfico 3: Heatmap ventas por tipo de motor y escenario
# ------------------------
pivot_heatmap = energetico_df.pivot_table(index='Tipo de motor', columns='Escenario',
                                          values='Ventas Escenario', aggfunc='sum', observed=False)
pivot_base = energetico_df.pivot_table(index='Tipo de motor', columns='Escenario',
                                       values='Ventas Año Base', aggfunc='sum', observed=False)
variacion_pct = (pivot_heatmap - pivot_base) / pivot_base

plt.figure(figsize=(14, 8))
ax = sns.heatmap(variacion_pct, annot=True, fmt='.1%', cmap='RdYlGn', center=0)
colorbar = ax.collections[0].colorbar
colorbar.ax.yaxis.set_major_formatter(mtick.PercentFormatter(1.0))
plt.title('Mapa de calor: Variación porcentual en las unidades vendidas para cada escenario, diferenciado por tipo de motor \n Escenarios filtrados por elasticidad = -1,87')
plt.tight_layout()
plt.savefig(carpeta_salida + "Heatmap variacion ventas por tipo motor.png", dpi=300)
plt.show()

# ------------------------
# Gráfico 4: Heatmap ventas por Tipo 2 y escenario
# ------------------------
pivot_heatmap2 = energetico_df.pivot_table(index='Tipo 2', columns='Escenario',
                                           values='Ventas Escenario', aggfunc='sum', observed=False)
pivot_base2 = energetico_df.pivot_table(index='Tipo 2', columns='Escenario',
                                        values='Ventas Año Base', aggfunc='sum', observed=False)
variacion_pct2 = (pivot_heatmap2 - pivot_base2) / pivot_base2

plt.figure(figsize=(14, 8))
ax = sns.heatmap(variacion_pct2, annot=True, fmt='.1%', cmap='RdYlGn', center=0)
colorbar = ax.collections[0].colorbar
colorbar.ax.yaxis.set_major_formatter(mtick.PercentFormatter(1.0))
plt.title('Mapa de calor: Variación porcentual en las unidades vendidas para cada escenario, diferenciado por categoría vehicular \n Escenarios filtrados por elasticidad = -1,87')
plt.tight_layout()
plt.savefig(carpeta_salida + "Heatmap variacion ventas por tipo 2.png", dpi=300)
plt.show()

# -------------------------------------------------------------------------------------------
# Gráfico 5: Heatmap de consumo promedio (L/100km) por tipo de motor (ponderado por ventas)
# -------------------------------------------------------------------------------------------

# Filtrar: excluir eléctricos ("BEV")
df_consumo = energetico_df[energetico_df["Tipo de motor"] != "BEV"].copy()

# Paso 1: Calcular el producto del consumo por las ventas para cada fila.
# Este producto se usará como el numerador en el cálculo del promedio ponderado.
# La unidad resultante es (L/100 km) * unidades vendidas.
df_consumo['Producto Consumo x Ventas Año Base'] = df_consumo['Consumo Año Base (L/100 km)'] * df_consumo['Ventas Año Base']
df_consumo['Producto Consumo x Ventas Escenario'] = df_consumo['Consumo Escenario (L/100 km)'] * df_consumo['Ventas Escenario']


# Paso 2 y 3: Agrupar por tipo de motor y escenario, sumar los productos y las ventas,
# para luego calcular los promedios ponderados.
grouped_consumo = df_consumo.groupby(['Tipo de motor', 'Escenario'], observed=False).agg(
    Suma_Producto_Consumo_Ventas_Ano_Base=('Producto Consumo x Ventas Año Base', 'sum'),
    Suma_Ventas_Ano_Base=('Ventas Año Base', 'sum'),
    Suma_Producto_Consumo_Ventas_Escenario=('Producto Consumo x Ventas Escenario', 'sum'),
    Suma_Ventas_Escenario=('Ventas Escenario', 'sum')
).reset_index()

# Calcular el consumo promedio ponderado (L/100 km) para el Año Base y el Escenario.
# Esto se hace dividiendo la suma del producto (Consumo x Ventas) por la suma de las Ventas.
# Manejar divisiones por cero si 'Suma_Ventas_Ano_Base' o 'Suma_Ventas_Escenario' son 0,
# lo que resultaría en un consumo promedio ponderado de 0 para ese grupo.
grouped_consumo['Consumo Ponderado Año Base (L/100 km)'] = grouped_consumo.apply(
    lambda row: row['Suma_Producto_Consumo_Ventas_Ano_Base'] / row['Suma_Ventas_Ano_Base'] if row['Suma_Ventas_Ano_Base'] != 0 else 0,
    axis=1
)
grouped_consumo['Consumo Ponderado Escenario (L/100 km)'] = grouped_consumo.apply(
    lambda row: row['Suma_Producto_Consumo_Ventas_Escenario'] / row['Suma_Ventas_Escenario'] if row['Suma_Ventas_Escenario'] != 0 else 0,
    axis=1
)

# Ahora, construye las tablas pivote con los promedios ponderados calculados.
pivot_consumo_esc_ponderado = grouped_consumo.pivot_table(index='Tipo de motor', columns='Escenario',
                                                          values='Consumo Ponderado Escenario (L/100 km)', observed=False)
pivot_consumo_base_ponderado = grouped_consumo.pivot_table(index='Tipo de motor', columns='Escenario',
                                                           values='Consumo Ponderado Año Base (L/100 km)', observed=False)


# Calcular variación porcentual (ahora usando los valores ponderados).
# Es crucial manejar la división por cero si Consumo Ponderado Año Base es cero.
# Si el Consumo Ponderado Año Base es 0:
# - Si Consumo Ponderado Escenario también es 0, la variación es 0 (no hay cambio donde no había consumo).
# - Si Consumo Ponderado Escenario es > 0, la variación es infinita, lo que indica un nuevo consumo.
#   Para la visualización del heatmap, los infinitos pueden causar problemas y no se mostrarán bien.
#   Los reemplazamos por NaN y luego por 0, lo que hará que esas celdas no se coloreen o se muestren como 0%.
variacion_consumo_ponderado = (pivot_consumo_esc_ponderado - pivot_consumo_base_ponderado) / pivot_consumo_base_ponderado

variacion_consumo_ponderado = variacion_consumo_ponderado.replace([float('inf'), -float('inf')], pd.NA).fillna(0)


# Graficar Heatmap
plt.figure(figsize=(14, 8))
ax = sns.heatmap(variacion_consumo_ponderado, annot=True, fmt='.2%', cmap='RdYlGn', center=0)
colorbar = ax.collections[0].colorbar
colorbar.ax.yaxis.set_major_formatter(mtick.PercentFormatter(1.0, decimals=2))
plt.title('Mapa de calor: Variación del consumo promedio (L/100km), diferenciado por tipo de motor \n Escenarios filtrados por elasticidad = -1,87')
plt.tight_layout()
plt.savefig(carpeta_salida + "Heatmap consumo promedio por tipo de motor.png", dpi=300)
plt.show()
# ------------------------
# Gráfico 6: Heatmap de consumo promedio (L/100km) por Tipo 2 (ponderado por ventas)
# ------------------------

# No es necesario filtrar BEV para Tipo 2, ya que BEV es un tipo de motor.
# Si un Tipo 2 no tiene ventas de ningún motor, naturalmente su promedio ponderado será 0 si se maneja la división por cero.
# Si tu intención es excluir *completamente* los BEV de los cálculos de Tipo 2,
# entonces deberías mantener la línea de filtrado:
# df_consumo_tipo2 = energetico_df[energetico_df["Tipo de motor"] != "BEV"].copy()
# Si no, simplemente usa energetico_df.copy() o el DataFrame original directamente si no lo modificas.
# Para este ejemplo, asumiremos que quieres incluir todos los vehículos que tienen un Consumo L/100km,
# lo cual implica que BEV no deberían tener un valor en esa columna.
# Sin embargo, para ser coherentes con el gráfico anterior, mantendremos el filtro explícito.
df_consumo_tipo2 = energetico_df[energetico_df["Tipo de motor"] != "BEV"].copy()


# Paso 1: Calcular el producto del consumo por las ventas para cada fila.
# Este producto se usará como el numerador en el cálculo del promedio ponderado.
# La unidad resultante es (L/100 km) * unidades vendidas.
df_consumo_tipo2['Producto Consumo x Ventas Año Base'] = df_consumo_tipo2['Consumo Año Base (L/100 km)'] * df_consumo_tipo2['Ventas Año Base']
df_consumo_tipo2['Producto Consumo x Ventas Escenario'] = df_consumo_tipo2['Consumo Escenario (L/100 km)'] * df_consumo_tipo2['Ventas Escenario']


# Paso 2 y 3: Agrupar por Tipo 2 y Escenario, sumar los productos y las ventas,
# para luego calcular los promedios ponderados.
# ¡La clave aquí es cambiar 'Tipo de motor' a 'Tipo 2' en el groupby!
grouped_consumo_tipo2 = df_consumo_tipo2.groupby(['Tipo 2', 'Escenario'], observed=False).agg(
    Suma_Producto_Consumo_Ventas_Ano_Base=('Producto Consumo x Ventas Año Base', 'sum'),
    Suma_Ventas_Ano_Base=('Ventas Año Base', 'sum'),
    Suma_Producto_Consumo_Ventas_Escenario=('Producto Consumo x Ventas Escenario', 'sum'),
    Suma_Ventas_Escenario=('Ventas Escenario', 'sum')
).reset_index()

# Calcular el consumo promedio ponderado (L/100 km) para el Año Base y el Escenario.
# Esto se hace dividiendo la suma del producto (Consumo x Ventas) por la suma de las Ventas.
# Manejar divisiones por cero si 'Suma_Ventas_Ano_Base' o 'Suma_Ventas_Escenario' son 0,
# lo que resultaría en un consumo promedio ponderado de 0 para ese grupo.
grouped_consumo_tipo2['Consumo Ponderado Año Base (L/100 km)'] = grouped_consumo_tipo2.apply(
    lambda row: row['Suma_Producto_Consumo_Ventas_Ano_Base'] / row['Suma_Ventas_Ano_Base'] if row['Suma_Ventas_Ano_Base'] != 0 else 0,
    axis=1
)
grouped_consumo_tipo2['Consumo Ponderado Escenario (L/100 km)'] = grouped_consumo_tipo2.apply(
    lambda row: row['Suma_Producto_Consumo_Ventas_Escenario'] / row['Suma_Ventas_Escenario'] if row['Suma_Ventas_Escenario'] != 0 else 0,
    axis=1
)

# Ahora, construye las tablas pivote con los promedios ponderados calculados.
pivot_consumo_esc_tipo2_ponderado = grouped_consumo_tipo2.pivot_table(index='Tipo 2', columns='Escenario',
                                                                      values='Consumo Ponderado Escenario (L/100 km)', observed=False)
pivot_consumo_base_tipo2_ponderado = grouped_consumo_tipo2.pivot_table(index='Tipo 2', columns='Escenario',
                                                                       values='Consumo Ponderado Año Base (L/100 km)', observed=False)


# Calcular variación porcentual (ahora usando los valores ponderados).
# Es crucial manejar la división por cero si Consumo Ponderado Año Base es cero.
# Para este heatmap, NaNs no se mostrarán y Infinitos pueden causar problemas, así que los reemplazamos.
variacion_consumo_tipo2_ponderado = (pivot_consumo_esc_tipo2_ponderado - pivot_consumo_base_tipo2_ponderado) / pivot_consumo_base_tipo2_ponderado

variacion_consumo_tipo2_ponderado = variacion_consumo_tipo2_ponderado.replace([float('inf'), -float('inf')], pd.NA).fillna(0)


# Graficar Heatmap
plt.figure(figsize=(14, 8))
# Asegúrate de usar la variable correcta para el heatmap
ax = sns.heatmap(variacion_consumo_tipo2_ponderado, annot=True, fmt='.2%', cmap='RdYlGn', center=0)
colorbar = ax.collections[0].colorbar
colorbar.ax.yaxis.set_major_formatter(mtick.PercentFormatter(1.0, decimals=2))
plt.title('Mapa de calor: Variación del consumo promedio (L/100km), diferenciado por categoría vehicular \n Escenarios filtrados por elasticidad = -1,87')
plt.tight_layout()
plt.savefig(carpeta_salida + "Heatmap consumo promedio Tipo2.png", dpi=300) # Nombre de archivo actualizado
plt.show()