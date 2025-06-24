import pandas as pd
import matplotlib.pyplot as plt
import os

# --------------------------------------------
# 1) CONFIGURACIÓN DE RUTAS Y LECTURA DE DATOS
# --------------------------------------------
excel_path = r"C:\Users\emili\PycharmProjects\TesisUY\Gráficos\Parque vehicular 2017-2024.xlsx"
output_dir = r"C:\Users\emili\PycharmProjects\TesisUY\Gráficos"
os.makedirs(output_dir, exist_ok=True)

# Leemos la hoja "Hoja2" usando openpyxl
df = pd.read_excel(excel_path, sheet_name="Hoja2", engine="openpyxl")

# Convertimos la columna 'año' en índice
df.set_index('año', inplace=True)


# --------------------------------------------
# 2) GRÁFICO APILADO (2017–2024)
#    Total de vehículos en parque automotor
#    (“Combustión y otros” + “Eléctricos”) por año,
#    anotando el % de eléctricos con dos decimales,
#    y eje Y en formato plain (sin notación 1e6).
# --------------------------------------------

# Extraer sólo las columnas “Combustión y otros” y “Eléctricos”
df_stack = df[['Combustión y otros', 'Eléctricos']].copy()

# Calcular el porcentaje de eléctricos sobre el total del parque cada año
percent_electric = df['Eléctricos'] / (df['Combustión y otros'] + df['Eléctricos']) * 100

# Crear figura y dibujar barras apiladas
plt.figure(figsize=(10, 6))
ax1 = df_stack.plot(
    kind='bar',
    stacked=True,
    figsize=(10, 6),
    color=['#FF8C00', '#DA70D6'],
    legend=True
)

# Desactivar notación científica en el eje Y
ax1.ticklabel_format(style='plain', axis='y')

# Anotar porcentaje de eléctricos con dos decimales encima de cada barra
max_val = (df_stack['Combustión y otros'] + df_stack['Eléctricos']).max()
offset = max_val * 0.01  # pequeño desplazamiento para colocar el texto
for i, year in enumerate(df_stack.index):
    total_height = df_stack.loc[year, 'Combustión y otros'] + df_stack.loc[year, 'Eléctricos']
    pct = percent_electric.loc[year]
    ax1.text(
        i,
        total_height + offset,
        f"{pct:.2f}%",
        ha='center',
        va='bottom',
        fontsize=10,
        fontweight='bold'
    )

# Configuración de título y ejes
ax1.set_title(
    'Total de vehículos en parque automotor\npor tecnología (2017–2024)',
    fontsize=14,
    pad=15
)
ax1.set_xlabel('Año', fontsize=12)
ax1.set_ylabel('Cantidad de vehículos', fontsize=12)
plt.xticks(rotation=0)
ax1.legend(title='Tecnología')

plt.tight_layout()

# Guardar figura en la ruta especificada
output_path1 = os.path.join(output_dir, 'parque_combustion_electricos_2017_2024.png')
plt.savefig(output_path1, dpi=300)
plt.close()


# --------------------------------------------
# 3) GRÁFICO DE % PARTICIPACIÓN ELÉCTRICA POR CATEGORÍA (2024)
# --------------------------------------------
# Tomar la fila correspondiente a 2024
df_2024 = df.loc[2024].copy()

# Definir categorías y sus columnas “eléctricas” (con nombres exactos de la hoja)
categorias = [
    'Automóviles',
    'Pick Up',
    'Utilitarios',
    'SUV, Crossover y Rural',
    'Taxis',
    'Remises'
]
electric_cols = [
    'Automóviles Eléctrico',
    'Pick Up Eléctrico',
    'Utilitarios Eléctrico',
    'SUV, Crossover y Rural Eléctrico',
    'Taxis Eléctrico',
    'Remises Eléctrico'
]

# Calcular % de eléctricos para cada categoría en 2024
percent_por_categoria = []
for cat, el_col in zip(categorias, electric_cols):
    total_cat = df_2024[cat]
    electr_cat = df_2024[el_col]
    if total_cat != 0:
        pct = (electr_cat / total_cat) * 100
    else:
        pct = 0.0
    percent_por_categoria.append(pct)

# Construir DataFrame auxiliar para graficar
df_pct_cat = pd.DataFrame({
    'Categoría': categorias,
    '% Eléctricos': percent_por_categoria
})

# Dibujar gráfico de barras del % eléctrico por categoría
plt.figure(figsize=(10, 6))
ax2 = df_pct_cat.plot(
    x='Categoría',
    y='% Eléctricos',
    kind='bar',
    legend=False,
    figsize=(10, 6),
    color='#DA70D6'
)

# Anotar el porcentaje con dos decimales encima de cada barra
for p in ax2.patches:
    height = p.get_height()
    ax2.annotate(
        f"{height:.1f}%",
        (p.get_x() + p.get_width() / 2, height),
        ha='center',
        va='bottom',
        fontsize=10,
        fontweight='bold'
    )

# Configuración de títulos y ejes
ax2.set_title(
    'Porcentaje de participación de vehículos eléctricos\npor categoría en 2024',
    fontsize=14,
    pad=15
)
ax2.set_xlabel('Categoría', fontsize=12)
ax2.set_ylabel('% Eléctricos', fontsize=12)
plt.xticks(rotation=30, ha='right')

plt.tight_layout()

# Guardar figura
output_path2 = os.path.join(output_dir, 'participacion_electrico_por_categoria_2024.png')
plt.savefig(output_path2, dpi=300)
plt.close()


# --------------------------------------------
# 4) Mensaje de confirmación (opcional)
# --------------------------------------------
print(f"Gráficos guardados en:\n  • {output_path1}\n  • {output_path2}")
