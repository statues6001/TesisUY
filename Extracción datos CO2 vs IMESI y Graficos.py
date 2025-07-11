import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import math

# Cargar base de datos del parque vehicular 2023
path = r'C:\Users\emili\PycharmProjects\TesisUY\Base de datos 2023 - Prueba_output.xlsx'
sheet = '2023'
df = pd.read_excel(path, sheet_name=sheet)

# Convertir a numérico las columnas de interés
df['CO2 NEDC (g/km)'] = pd.to_numeric(df['CO2 NEDC (g/km)'], errors='coerce')
df['Procesados'] = pd.to_numeric(df['Procesados'], errors='coerce')

# Filtrar datos. Se requieren las columnas para el boxplot y para ventas totales
df_clean = df.dropna(subset=['CO2 NEDC (g/km)', 'Consumo (L/100 km)', 'IMESI (%)', 'Procesados'])
df_clean = df_clean[df_clean['IMESI (%)'] != 0]

# Crear la columna IMESI_x100
df_clean['IMESI_x100'] = (df_clean['IMESI (%)'] * 100).round(2)

# Obtener los valores únicos de IMESI_x100 ordenados
unique_imesi = sorted(df_clean['IMESI_x100'].unique())
print("Valores únicos de IMESI_x100:", unique_imesi)

# Diccionario para mapear las categorías de vehículo
# Porcentajes y categorías corresponden a Decreto 390/021
map_imesi = {
    2.0: 'PHEV (0 - 2.500 c.c.)',
    3.45: 'HEV (0 - 2.500 c.c.)',
    6.0: 'Utilitarios gasolina (0 - 3.500 c.c.)',
    7.0: 'MHEV (0 - 1.500 c.c.)',
    14.0: 'MHEV (1.500 - 2.000 c.c.)',
    23.0: 'Automóviles y SUV gasolina (0 - 1.000 c.c.)',
    28.75: 'Automóviles y SUV gasolina (1.000 - 1.500 c.c.)',
    34.5: 'Automóviles y SUV gasolina (1.500 hasta 2.000 c.c.)',
    34.7: 'Utilitarios diesel (0 hasta 3.500 c.c.)',
    40.25: 'Automóviles y SUV gasolina (2.000 - 3.000 c.c.)',
    46.0: 'Automóviles y SUV gasolina (3.000 c.c. - ∞)',
    115.0: 'Automóviles y SUV diesel (0 - ∞)'
}

##############################################
# Paso 1: Calcular ventas totales por franja de IMESI
##############################################
ventas_totales_por_imesi = df_clean.groupby('IMESI_x100')['Procesados'].sum().reindex(unique_imesi)

df_resultados = ventas_totales_por_imesi.reset_index()
df_resultados.columns = ['IMESI_x100', 'Ventas_Totales']
df_resultados['IMESI (%)'] = df_resultados['IMESI_x100'] / 100.0
# Asignar categoría de vehículo según % IMESI
df_resultados['Categoría de Vehículo'] = df_resultados['IMESI_x100'].map(map_imesi)

##############################################
# Paso 2: Calcular estadísticas de CO₂ por franja de IMESI (boxplot)
##############################################
boxplot_data = {}
for imesi_value in unique_imesi:
    subset = df_clean[df_clean['IMESI_x100'] == imesi_value]['CO2 NEDC (g/km)']
    if len(subset) > 0:
        q1 = np.percentile(subset, 25)
        q2 = np.percentile(subset, 50)  # Mediana
        q3 = np.percentile(subset, 75)
        iqr = q3 - q1
        lower_whisker = max(subset[subset >= (q1 - 1.5 * iqr)].min(), subset.min())
        upper_whisker = min(subset[subset <= (q3 + 1.5 * iqr)].max(), subset.max())
        boxplot_data[imesi_value] = {
            'Q1 (25%) CO2 (g/km)': q1,
            'Mediana (50%) CO2 (g/km)': q2,
            'Q3 (75%) CO2 (g/km)': q3,
            'IQR': iqr,
            'Valor límite inferior CO2 (g/km)': lower_whisker,
            'Valor límite superior CO2 (g/km)': upper_whisker,
            'Valores atípicos CO2 (g/km)': list(subset[(subset < lower_whisker) | (subset > upper_whisker)].values)
        }

df_boxplot = pd.DataFrame.from_dict(boxplot_data, orient='index')
df_boxplot['Categoría de Vehículo'] = df_boxplot.index.map(map_imesi)
df_boxplot.index.name = '% IMESI'
df_boxplot['IMESI'] = df_boxplot.index.astype(float) / 100.0

# Agregar la Media ponderada de CO2 para cada IMESI como nueva columna
weighted_mean = df_clean.groupby('IMESI_x100').apply(
    lambda x: np.average(x['CO2 NEDC (g/km)'], weights=x['Procesados'])
).reindex(unique_imesi)
df_boxplot['CO₂ promedio ponderado por ventas (g/km)'] = weighted_mean.values

output_path_boxplot = r'C:\Users\emili\PycharmProjects\TesisUY\Datos CO2 vs IMESI 2023.xlsx'
df_boxplot.to_excel(output_path_boxplot, index=True)
print(f"Datos de boxplot guardados en: {output_path_boxplot}")

##############################################
# Paso 3: Generar gráfico combinado de boxplot (CO₂) y barras (ventas)
##############################################
colors = sns.color_palette("tab20", n_colors=len(unique_imesi))

fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 10), gridspec_kw={'height_ratios': [3, 1]})

# Boxplot de CO2 vs. % IMESI
sns.boxplot(
    x='IMESI_x100',
    y='CO2 NEDC (g/km)',
    data=df_clean,
    palette=colors,
    hue='IMESI_x100',
    order=unique_imesi,
    dodge=False,
    ax=ax1,
    legend=False
)
ax1.set_title("CO$_2$ NEDC (g/km) vs. % IMESI")
ax1.set_xlabel("")
ax1.set_ylabel("CO$_2$ NEDC (g/km)")

patches = []
for i, val in enumerate(unique_imesi):
    label_text = map_imesi[val] if val in map_imesi else str(val)
    patch = mpatches.Patch(color=colors[i], label=label_text)
    patches.append(patch)
ax1.legend(handles=patches, title="Tipo de vehículo", bbox_to_anchor=(1.05, 1), loc='upper left')

# Gráfico de barras de ventas totales
sns.barplot(
    x=ventas_totales_por_imesi.index,
    y=ventas_totales_por_imesi.values,
    hue=ventas_totales_por_imesi.index,
    palette=colors,
    ax=ax2,
    legend=False
)
ax2.set_xlabel("% IMESI")
ax2.set_ylabel("Ventas Totales")
ax2.set_title("Ventas totales por categoría IMESI")

plt.tight_layout()
fig.savefig(r'C:\Users\emili\PycharmProjects\TesisUY\Gráficos\Boxplot CO2vsIMESI y gráfico de ventas totales.png', dpi=300)
#plt.show()

##############################################
# Paso 4: Histogramas de CO₂ ponderados por ventas, por franja IMESI
##############################################
n_categories = len(unique_imesi)
cols = 3  # Número de columnas para los subplots
rows = math.ceil(n_categories / cols)

fig_hist, axes = plt.subplots(rows, cols, figsize=(cols * 5, rows * 4))
axes = axes.flatten()

for i, imesi in enumerate(unique_imesi):
    ax = axes[i]
    subset = df_clean[df_clean['IMESI_x100'] == imesi]
    ax.hist(subset['CO2 NEDC (g/km)'], bins=50, weights=subset['Procesados'],
            color=colors[i % len(colors)], edgecolor='black')
    # Agregar subtítulo especial para IMESI = 34.5
    if imesi == 34.5:
        title = (f"{imesi}%\n - {map_imesi.get(imesi, '')}\n"
                 "Incluye PHEV y HEV (2.500 c.c. - ∞), MHEV (2.000 c.c. - ∞)")
    else:
        title = f"{imesi}%\n - {map_imesi.get(imesi, '')}"
    ax.set_title(title, fontsize=9)
    ax.set_xlabel("CO₂ NEDC (g/km)", fontsize=9)
    ax.set_ylabel("Ventas", fontsize=9)

for j in range(i+1, len(axes)):
    fig_hist.delaxes(axes[j])

# Cambiar "fig" por "fig_hist" en las siguientes líneas:
fig_hist.suptitle('Histograma de emisiones de CO₂ según volumen de ventas', fontsize=14)
#fig_hist.text(0.5, 0.04, 'CO₂ NEDC (g/km)', ha='center', fontsize=12)
#fig_hist.text(0.04, 0.5, 'Ventas', va='center', rotation='vertical', fontsize=12)

plt.subplots_adjust(top=0.88, bottom=0.08, left=0.08, right=0.97, hspace=0.5, wspace=0.35)
fig_hist.savefig(r'C:\Users\emili\PycharmProjects\TesisUY\Gráficos\Histograma de CO2, según volumen de ventas.png', dpi=300)
#plt.show()

##############################################
# Paso 5: Distribución de CO₂ acumulado por ventas, por franja de IMESI
##############################################

# Colores para cada franja de IMESI
colors_dict = {val: colors[i % len(colors)] for i, val in enumerate(unique_imesi)}

# Número total de franjas
n_franjas = len(unique_imesi)
cols = 3
rows = math.ceil(n_franjas / cols)

fig, axes = plt.subplots(rows, cols, figsize=(cols * 6, rows * 4))
axes = axes.flatten()

for i, imesi in enumerate(unique_imesi):
    ax = axes[i]
    subset = df_clean[df_clean['IMESI_x100'] == imesi][['CO2 NEDC (g/km)', 'Procesados']].dropna()

    if subset.empty:
        continue

    # Expandir filas según ventas
    expanded = subset.loc[subset.index.repeat(subset['Procesados'].astype(int))].copy()
    expanded.sort_values('CO2 NEDC (g/km)', inplace=True)

    # % acumulado
    expanded['Acumulado'] = np.arange(1, len(expanded) + 1) / len(expanded) * 100

    ax.plot(expanded['Acumulado'], expanded['CO2 NEDC (g/km)'],
            color=colors_dict[imesi], lw=2)

    if imesi == 34.5:
        titulo = (f"{imesi}%\n{map_imesi.get(imesi, '')}\n"
                  "Incluye PHEV y HEV (2.500 c.c. - ∞), MHEV (2.000 c.c. - ∞)")
    else:
        titulo = f"{imesi}%\n{map_imesi.get(imesi, '')}"

    ax.set_title(titulo, fontsize=9)
    ax.set_xlabel('% acumulado de ventas', fontsize=9)
    ax.set_ylabel('CO₂ NEDC (g/km)', fontsize=9)
    ax.grid(True)

# Eliminar ejes vacíos
for j in range(i + 1, len(axes)):
    fig.delaxes(axes[j])

fig.suptitle('Distribución de CO₂ según acumulado de ventas por franja IMESI', fontsize=14)
plt.subplots_adjust(top=0.88, bottom=0.08, left=0.08, right=0.97, hspace=0.5, wspace=0.35)
fig.savefig(r'C:\Users\emili\PycharmProjects\TesisUY\Gráficos\Distribucion CO2 acumulado.png', dpi=300)
plt.close(fig)


