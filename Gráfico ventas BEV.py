import pandas as pd
import matplotlib.pyplot as plt
import os

# Definir los datos según la tabla
data = {
    'Total': [49265, 52857, 58367, 65909],
    'Combustión y otros': [48408, 51316, 55852, 60053],
    'Eléctricos': [857, 1541, 2515, 5856]
}
years = ['2021', '2022', '2023', '2024']

# Crear DataFrame
df = pd.DataFrame(data, index=years)

# Seleccionar solo las columnas de combustión y eléctricos (el Total solo sirve para calcular el porcentaje)
df_stack = df[['Combustión y otros', 'Eléctricos']]

# Calcular el porcentaje de Eléctricos sobre el Total
percent_electric = df['Eléctricos'] / df['Total'] * 100

# Definir ruta de destino y crear carpeta si no existe
output_dir = r"C:\Users\emili\PycharmProjects\TesisUY\Gráficos"
os.makedirs(output_dir, exist_ok=True)

# Dibujar gráfico de barras apiladas sin la serie "Total"
plt.figure(figsize=(10, 6))
ax = df_stack.plot(
    kind='bar',
    stacked=True,
    figsize=(10, 6),
    color=['#FF8C00', '#DA70D6']
)

# Anotar el porcentaje de Eléctricos sobre el Total encima de cada barra
for i, year in enumerate(years):
    total_height = df['Total'].iloc[i]
    pct = percent_electric.iloc[i]
    ax.text(
        i,
        total_height + 1000,
        f"{pct:.1f}%",
        ha='center',
        va='bottom',
        fontsize=10,
        fontweight='bold'
    )

# Configurar títulos y etiquetas
ax.set_title('Venta anual de vehículos cero kilómetro por tecnología', fontsize=14)
ax.set_xlabel('Año', fontsize=12)
ax.set_ylabel('Cantidad de vehículos', fontsize=12)
plt.xticks(rotation=0)
plt.legend(title='Tecnología')

# Ajustar diseño
plt.tight_layout()

# Guardar el gráfico en la ruta especificada
output_path = os.path.join(output_dir, 'Ventas por tecnologia 2021 a 2024.png')
plt.savefig(output_path, dpi=300)

# Mostrar el gráfico
plt.show()

