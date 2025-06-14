import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import numpy as np
import re

archivo_resultados = "Salida Escenarios.xlsx"
df_resumen = pd.read_excel(archivo_resultados, sheet_name="Resumen Escenarios")




ventas_no_bev_base = df_resumen["Ventas Año Base (Sin BEV)"].iloc[0]
ventas_bev_base = df_resumen["Ventas BEV Año Base"].iloc[0]
consumo_sin_bev_año_base_unidad = df_resumen["Consumo Año Base (sin BEV) (tep)/Unidad"].iloc[0]
consumo_bev_año_base_unidad = df_resumen["Consumo eléctrico Año Base BEV (tep)/Unidad"].iloc[0]
consumo_año_base = (ventas_no_bev_base * consumo_sin_bev_año_base_unidad) + (ventas_bev_base * consumo_bev_año_base_unidad)

# Escenarios porcentaje de participación de eléctricos de 8.9% en 2024. Dato real.
escenarios_penetracion = {
    "Pesimista": [8.9, 10, 11, 12, 13],
    "Tendencial": [8.9, 12, 15, 18, 22],
    "Acelerado": [8.9, 15, 23, 32, 40]
}

# Ventas totales de vehículos
ventas_base_2023 = 57183  # ventas año base
ventas_totales_2024_adelante = 57183  # ventas proyectadas. Se toman constantes.
# 65909 ventas reales en 2024

resultados_proyeccion = []

# Iterar cada escenario IMESI de de la hoja Resumen Escenarios
for idx, fila in df_resumen.iterrows():
    escenario_base = fila["Escenario"]

    # Variables constantes IMESI desde Excel
    imesi_sin_bev_escenario_unidad = fila["Recaudación IMESI Escenario (Sin BEV) / Unidad"]
    imesi_bev_escenario_unidad = fila["Recaudación IMESI Escenario BEV / Unidad"]

    # Recaudación total IMESI año base (2023) desde Excel original para cada escenario
    recaudacion_año_base_total = fila["Recaudación IMESI Año Base (Sin BEV) (USD)"] + fila["Recaudación IMESI Año Base BEV (USD)"]

    # Consumo sin BEV y BEV
    consumo_sin_bev_tep_escenario = fila["Consumo Escenario (sin BEV) (tep)/Unidad"]
    consumo_bev_tep_escenario = fila["Consumo eléctrico Escenario BEV (tep)/Unidad"]


    #Factores de emision por unidad
    emisiones_bev_año_base_unidad = fila["Emisiones BEV CO2 Año Base (ton)/Unidad"]
    emisiones_bev_escenario_unidad = fila["Emisiones BEV CO2 Escenario (ton)/Unidad"]
    emisiones_sin_bev_año_base_unidad = fila["Emisiones Sin BEV CO2 Año Base (ton)/Unidad"]
    emisiones_sin_bev_escenario_unidad = fila["Emisiones Sin BEV CO2 Escenario (ton)/Unidad"]

    # Calcular emisiones en el año base (fijas para cada escenario)
    emisiones_año_base = (ventas_bev_base * emisiones_bev_año_base_unidad) + (
                ventas_no_bev_base * emisiones_sin_bev_año_base_unidad)

    # Proyección para cada escenario de penetración BEV desde 2024 a 2028
    for tipo_penetracion, penetraciones in escenarios_penetracion.items():
        for año_offset, bev_pct in enumerate(penetraciones, start=1):
            año = 2023 + año_offset

            # Ventas anuales: 2024 en adelante se usan ventas actualizadas
            ventas_totales = ventas_totales_2024_adelante if año >= 2024 else ventas_base_2023

            ventas_bev_escenario = ventas_totales * (bev_pct / 100)
            ventas_sin_bev_escenario = ventas_totales - ventas_bev_escenario

            recaudacion_imesi_escenario = (ventas_bev_escenario * imesi_bev_escenario_unidad) + (ventas_sin_bev_escenario * imesi_sin_bev_escenario_unidad)

            # Diferencia respecto al año base real del escenario original (2023)
            diferencia_recaudacion = recaudacion_imesi_escenario - recaudacion_año_base_total

            # -----------------------------------------------------------------
            # Cálculo del consumo energético en tep
            # -----------------------------------------------------------------
            consumo_energetico_sin_bev_escenario = ventas_sin_bev_escenario * consumo_sin_bev_tep_escenario
            consumo_energetico_bev_escenario = ventas_bev_escenario * consumo_bev_tep_escenario
            consumo_energetico_total_escenario = consumo_energetico_sin_bev_escenario + consumo_energetico_bev_escenario

            # -----------------------------------------------------------------
            # Cálculo del consumo energético en GWh
            # Factor de conversión: 1 tep = 0.01163 GWh
            # -----------------------------------------------------------------
            consumo_energetico_sin_bev_escenario_GWh = consumo_energetico_sin_bev_escenario*0.01163
            consumo_energetico_bev_escenario_GWh = consumo_energetico_bev_escenario*0.01163
            consumo_energetico_total_escenario_GWh = consumo_energetico_total_escenario*0.01163
            consumo_año_base_GWh = consumo_año_base*0.01163

            # -----------------------------------------------------------------
            # Cálculo de emisiones (CO2 en ton)
            # -----------------------------------------------------------------
            emisiones_sin_bev = ventas_sin_bev_escenario * emisiones_sin_bev_escenario_unidad
            emisiones_bev = ventas_bev_escenario * emisiones_bev_escenario_unidad
            emisiones_total = emisiones_sin_bev + emisiones_bev

            resultados_proyeccion.append({
                "Escenario": escenario_base,
                "Tipo de penetración": tipo_penetracion,
                "Año": año,
                "% Participación BEV": bev_pct,
                "Ventas BEV": ventas_bev_escenario,
                "Ventas No-BEV": ventas_sin_bev_escenario,
                "Recaudación IMESI BEV (USD)": ventas_bev_escenario * imesi_bev_escenario_unidad,
                "Recaudación IMESI No-BEV (USD)": ventas_sin_bev_escenario * imesi_sin_bev_escenario_unidad,
                "Recaudación IMESI Total (USD)": recaudacion_imesi_escenario,
                "Diferencia IMESI vs Año Base 2023 (USD)": diferencia_recaudacion,
                "Consumo Energético Sin BEV (tep)": consumo_energetico_sin_bev_escenario,
                "Consumo Energético BEV (tep)": consumo_energetico_bev_escenario,
                "Consumo Energético Total (tep)": consumo_energetico_total_escenario,
                "Consumo Energético Sin BEV (GWh)": consumo_energetico_sin_bev_escenario_GWh,
                "Consumo Energético BEV (GWh)": consumo_energetico_bev_escenario_GWh,
                "Consumo Energético Total (GWh)": consumo_energetico_total_escenario_GWh,
                "Consumo Año Base (tep)": consumo_año_base,
                "Consumo Año Base (GWh)": consumo_año_base_GWh,
                "Emisiones CO2 Año Base (ton)": emisiones_año_base,
                "Emisiones CO2 Sin BEV (ton)": emisiones_sin_bev,
                "Emisiones CO2 BEV (ton)": emisiones_bev,
                "Emisiones CO2 Total (ton)": emisiones_total
            })

# Generar DataFrame y Excel final
df_proyecciones = pd.DataFrame(resultados_proyeccion)

# Extraer el valor de elasticidad de la columna "Escenario"
# Se asume que el formato es "Lineal Escenario x (E=y)"
df_proyecciones["Elasticidad"] = df_proyecciones["Escenario"].str.extract(r"\(E=([-\d\.]+)\)").astype(float)

# -------------------------- Guardado del Excel con dos pestañas --------------------------
# Definir columnas para cada pestaña
cols_recaudacion = [
    "Escenario", "Tipo de penetración", "Año", "% Participación BEV",
    "Ventas BEV", "Ventas No-BEV",
    "Recaudación IMESI BEV (USD)", "Recaudación IMESI No-BEV (USD)",
    "Recaudación IMESI Total (USD)", "Diferencia IMESI vs Año Base 2023 (USD)"
]

cols_energia = [
    "Escenario", "Tipo de penetración", "Año", "% Participación BEV",
    "Ventas BEV", "Ventas No-BEV",
    "Consumo Energético Sin BEV (tep)", "Consumo Energético BEV (tep)",
    "Consumo Energético Total (tep)", "Consumo Año Base (tep)", "Consumo Energético Sin BEV (GWh)",
    "Consumo Energético BEV (GWh)", "Consumo Energético Total (GWh)", "Consumo Año Base (GWh)",
    "Emisiones CO2 Año Base (ton)", "Emisiones CO2 Sin BEV (ton)",
    "Emisiones CO2 BEV (ton)", "Emisiones CO2 Total (ton)"
]

df_recaudacion = df_proyecciones[cols_recaudacion].copy()
df_energia = df_proyecciones[cols_energia].copy()

# Escribir ambos DataFrames en un mismo archivo Excel con dos hojas
with pd.ExcelWriter("Proyección 5 años.xlsx", engine="openpyxl") as writer:
    df_recaudacion.to_excel(writer, sheet_name="Proyección Recaudación", index=False)
    df_energia.to_excel(writer, sheet_name="Proyección Energía", index=False)

print("Archivo Excel generado con las pestañas 'Proyección Recaudación' y 'Proyección Energía y Emisiones'.")

# -------------------------------------------------------------------------------------------------------------------
# Parte 2: Gráficos relacionados
# -------------------------------------------------------------------------------------------------------------------
plt.rcParams['figure.figsize'] = (10, 6)
escenarios = df_proyecciones["Tipo de penetración"].unique()

# -------------------------------------------------------------------------------------------------------------------
# Gráfico 1 - Evolución de la participación BEV (%) por escenario
# -------------------------------------------------------------------------------------------------------------------
# Eliminar filas duplicadas para el gráfico 1
df_plot1 = df_proyecciones.drop_duplicates(subset=["Tipo de penetración", "Año", "% Participación BEV"])

escenarios = df_plot1["Tipo de penetración"].unique()

plt.figure()
for escenario in escenarios:
    datos = df_plot1[df_plot1["Tipo de penetración"] == escenario]
    plt.plot(datos["Año"], datos["% Participación BEV"], marker='o', label=escenario)

plt.title('Evolución de la participación BEV (%) por escenario')
plt.xlabel('Año')
plt.ylabel('% Participación BEV')
plt.grid(True, linestyle='--', alpha=0.6)
plt.legend()
ax = plt.gca()
ax.xaxis.set_major_locator(ticker.MaxNLocator(integer=True))
plt.tight_layout()
plt.savefig(r"C:\Users\emili\PycharmProjects\TesisUY\Gráficos\Proyección 5 años\grafico participacion BEV.png", dpi=300)
##plt.show()

# -------------------------------------------------------------------------------------------------------------------
# Gráfico 2: Diferencia IMESI vs Año Base millones 2023 (USD). Filtrado por Elasticidad y Escenario deseado
# Se generan tantos subplots como Escenarios deseados
# -------------------------------------------------------------------------------------------------------------------

# Filtrar por elasticidad deseada
elasticidad_objetivo = -1.87
df_grafico2 = df_proyecciones[df_proyecciones["Elasticidad"] == elasticidad_objetivo]

# Filtrar escenarios específicos: se usa .str.contains para buscar cualquiera de los patrones indicados
escenarios_deseados = ["Lineal Escenario 6", "Lineal Escenario 4", "Lineal Escenario 3", "Lineal Escenario 7"]
pattern = '|'.join(escenarios_deseados)
df_filtrado = df_grafico2[df_grafico2["Escenario"].str.contains(pattern)]

# Extraer escenarios únicos después de filtrar
escenarios_unicos = df_filtrado["Escenario"].unique()
n_escenarios = len(escenarios_unicos)

fig, axs = plt.subplots(nrows=1, ncols=n_escenarios, figsize=(5 * n_escenarios, 5), sharey=True)

# Si solo hay un escenario, se fuerza a que axs sea una lista para iterar
if n_escenarios == 1:
    axs = [axs]

for i, escenario in enumerate(escenarios_unicos):
    ax = axs[i]
    # Filtrar data para el escenario actual
    df_escenario = df_filtrado[df_filtrado["Escenario"] == escenario]
    # Obtener los años ordenados
    anios = sorted(df_escenario["Año"].unique())
    x = np.arange(len(anios))
    # Obtener los tipos de penetración disponibles en este escenario
    tipos_penetracion = df_escenario["Tipo de penetración"].unique()
    n_tipos = len(tipos_penetracion)
    ancho_barra = 0.8 / n_tipos  # Se usa el 80% del espacio disponible

    for j, tipo in enumerate(tipos_penetracion):
        df_tipo = df_escenario[df_escenario["Tipo de penetración"] == tipo].sort_values("Año")
        # Convertir la diferencia a millones de USD
        valores = df_tipo["Diferencia IMESI vs Año Base 2023 (USD)"].values / 1e6
        ax.bar(x + j * ancho_barra, valores, width=ancho_barra, label=tipo)

    ax.set_title(escenario)
    ax.set_xlabel('Año')
    if i == 0:
        ax.set_ylabel('Diferencia IMESI vs Año Base 2023 (millones USD)')
    ax.set_xticks(x + ancho_barra * (n_tipos - 1) / 2)
    ax.set_xticklabels(anios)
    ax.axhline(0, color='black', linewidth=0.8)
    ax.grid(True, linestyle='--', alpha=0.6)
    ax.legend()

plt.suptitle('Diferencia IMESI vs Año Base 2023')
plt.tight_layout()
plt.savefig(r"C:\Users\emili\PycharmProjects\TesisUY\Gráficos\Proyección 5 años\Diferencia IMESI vs Año Base 2023.png", dpi=300)
##plt.show()

# -------------------------------------------------------------------------------------------------------------------
# Gráfico 3: Comparativo acumulado de recaudación IMESI total (2024-2028) para excenarios lineales y tipo de penetración
# -------------------------------------------------------------------------------------------------------------------

# Filtrar elasticidad específica. Al filtrar por elasticidad deja afuera el escalonado ya que ese escenario no esta
# separado por elasticidad.
elasticidad_objetivo = -1.87
df_filtrado = df_proyecciones[df_proyecciones["Elasticidad"] == elasticidad_objetivo]

# Filtrar años 2024-2028 y sumar acumuladamente
df_acumulado = df_filtrado[df_filtrado["Año"].between(2024, 2028)].groupby(
    ["Escenario", "Tipo de penetración"]
)["Recaudación IMESI Total (USD)"].sum().reset_index()

# Añadir columna del valor base constante (5 años)
valor_base_constante = df_resumen["Recaudación IMESI Año Base (Sin BEV) (USD)"].iloc[0] + df_resumen["Recaudación IMESI Año Base BEV (USD)"].iloc[0]
valor_base_5anios = valor_base_constante * 5

# Ordenar los escenarios fiscales para facilitar comparación
def obtener_numero(escenario):
    match = re.search(r'\d+', escenario)
    return int(match.group()) if match else 0

escenarios_fiscales = sorted(df_acumulado["Escenario"].unique(), key=obtener_numero)

# Tipos de penetración claramente definidos
tipos_penetracion = ["Pesimista", "Tendencial", "Acelerado"]
ancho_barra = 0.2

# Preparar posiciones para barras agrupadas
posiciones = np.arange(len(escenarios_fiscales))

plt.figure(figsize=(14, 7))

# Dibujar barras para cada tipo de penetración
for i, tipo in enumerate(tipos_penetracion):
    datos_tipo = []
    for esc in escenarios_fiscales:
        valor = df_acumulado[
            (df_acumulado["Escenario"] == esc) &
            (df_acumulado["Tipo de penetración"] == tipo)
        ]["Recaudación IMESI Total (USD)"].values[0] / 1e6  # millones USD
        datos_tipo.append(valor)

    plt.bar(posiciones + i * ancho_barra, datos_tipo, ancho_barra, label=tipo)

# Barra adicional para valor base
datos_base = [valor_base_5anios / 1e6] * len(escenarios_fiscales)
plt.bar(posiciones + 3 * ancho_barra, datos_base, ancho_barra, label="Recaudación Base 2023 (5 años)", color='gray', alpha=0.6)

# Configuración visual
plt.xlabel("Escenario", fontsize=12)
plt.ylabel("Recaudación Acumulada IMESI 2024-2028 (millones USD)", fontsize=12)
plt.title(f"Comparación acumulada de recaudación IMESI (2024-2028)\nElasticidad = {elasticidad_objetivo}", fontsize=14)
plt.xticks(posiciones + ancho_barra, escenarios_fiscales, rotation=45, ha='right')
plt.grid(axis='y', linestyle='--', alpha=0.7)
plt.legend(title="Tipo de penetración", loc='upper left', bbox_to_anchor=(1, 1))
plt.tight_layout(rect=[0, 0, 1, 1])
plt.savefig(r"C:\Users\emili\PycharmProjects\TesisUY\Gráficos\Proyección 5 años\comparacion acumulada recaudacion IMESI.png", dpi=300)
#plt.show()

# -------------------------------------------------------------------------------------------------------------------
# Gráfico 4,5 y 6: Comparativo acumulado consumos energéticos (2024-2028) por tipo de penetración y Elasticidad única
# -------------------------------------------------------------------------------------------------------------------

# Filtrar elasticidad específica
elasticidad_objetivo = -1.87
df_filtrado = df_proyecciones[df_proyecciones["Elasticidad"] == elasticidad_objetivo]

# Extraer datos base exactos desde archivo original
ventas_no_bev_base = df_resumen["Ventas Año Base (Sin BEV)"].iloc[0]
ventas_bev_base = df_resumen["Ventas BEV Año Base"].iloc[0]

# Consumo unitario desde el archivo original
consumo_sin_bev_unidad_base = df_resumen["Consumo Año Base (sin BEV) (tep)/Unidad"].iloc[0]
consumo_bev_unidad_base = df_resumen["Consumo eléctrico Año Base BEV (tep)/Unidad"].iloc[0]

# Cálculo exacto del consumo base (multiplicado por 5 años)
consumo_base_sin_bev_5anios = consumo_sin_bev_unidad_base * ventas_no_bev_base * 5
consumo_base_bev_5anios = consumo_bev_unidad_base * ventas_bev_base * 5
consumo_base_total_5anios = consumo_base_sin_bev_5anios + consumo_base_bev_5anios

# Acumulación datos proyección (2024-2028)
df_acumulado = df_filtrado[df_filtrado["Año"].between(2024, 2028)].groupby(
    ["Escenario", "Tipo de penetración"]
).agg({
    "Consumo Energético Sin BEV (tep)": "sum",
    "Consumo Energético BEV (tep)": "sum",
    "Consumo Energético Total (tep)": "sum"
}).reset_index()

# Escenarios fiscales ordenados
escenarios_fiscales = sorted(df_acumulado["Escenario"].unique(), key=obtener_numero)
tipos_penetracion = ["Pesimista", "Tendencial", "Acelerado"]
ancho_barra = 0.2
posiciones = np.arange(len(escenarios_fiscales))

# Función reutilizable para gráficos
def generar_grafico(columna, valor_base, titulo, nombre_archivo):
    plt.figure(figsize=(14, 7))

    for i, tipo in enumerate(tipos_penetracion):
        datos_tipo = [
            df_acumulado[
                (df_acumulado["Escenario"] == esc) &
                (df_acumulado["Tipo de penetración"] == tipo)
            ][columna].values[0] / 1e3  # miles de tep
            for esc in escenarios_fiscales
        ]
        plt.bar(posiciones + i * ancho_barra, datos_tipo, ancho_barra, label=tipo)

    # Barra adicional base
    datos_base = [valor_base / 1e3] * len(escenarios_fiscales)
    plt.bar(posiciones + 3 * ancho_barra, datos_base, ancho_barra, label="Consumo Base 2023 (5 años)", color='gray', alpha=0.6)

    # Detalles visuales
    plt.xlabel("Escenario", fontsize=12)
    plt.ylabel("Consumo acumulado (miles de tep)", fontsize=12)
    plt.title(f"{titulo}\nElasticidad = {elasticidad_objetivo}", fontsize=14)
    plt.xticks(posiciones + 1.5 * ancho_barra, escenarios_fiscales, rotation=45, ha='right')
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    plt.legend(title="Tipo de penetración", loc='upper left', bbox_to_anchor=(1, 1))
    plt.tight_layout(rect=[0, 0, 1, 1])
    plt.savefig(nombre_archivo, dpi=300)
    #plt.show()

# 1 - Consumo Energético Sin BEV (tep)
generar_grafico(
    columna="Consumo Energético Sin BEV (tep)",
    valor_base=consumo_base_sin_bev_5anios,
    titulo="Comparación acumulada del Consumo Energético Sin BEV (2024-2028)",
    nombre_archivo=r"C:\Users\emili\PycharmProjects\TesisUY\Gráficos\Proyección 5 años\comparacion consumo sin BEV.png"
)

# 2 - Consumo Energético BEV (tep)
generar_grafico(
    columna="Consumo Energético BEV (tep)",
    valor_base=consumo_base_bev_5anios,
    titulo="Comparación acumulada del Consumo Energético BEV (2024-2028)",
    nombre_archivo=r"C:\Users\emili\PycharmProjects\TesisUY\Gráficos\Proyección 5 años\comparacion consumo BEV.png"
)

# 3 - Consumo Energético Total (tep)
generar_grafico(
    columna="Consumo Energético Total (tep)",
    valor_base=consumo_base_total_5anios,
    titulo="Comparación acumulada del Consumo Energético Total (2024-2028)",
    nombre_archivo=r"C:\Users\emili\PycharmProjects\TesisUY\Gráficos\Proyección 5 años\comparacion consumo total.png"
)

# --------------------------------------------------
# Gráfico 7: Evolución de las emisiones CO2 totales (ton) por tipo de penetración
# Filtrado por elasticidad ingresada
# --------------------------------------------------
# Filtrar el DataFrame por la elasticidad deseada (-1.87 en este ejemplo)
df_grafico7 = df_proyecciones[df_proyecciones["Elasticidad"] == -1.87]

# Escenarios que se desean filtrar (asegúrate de que estas subcadenas estén presentes en los nombres)
escenarios_deseados = ["Lineal Escenario 4", "Lineal Escenario 5", "Lineal Escenario 6"]

# Crear subplots: 1 fila, 3 columnas
fig, axs = plt.subplots(nrows=1, ncols=3, figsize=(18, 6), sharey=True)

# Iterar sobre cada subplot y escenario
for i, escenario in enumerate(escenarios_deseados):
    # Filtrar las filas correspondientes al escenario actual usando str.contains
    df_escenario = df_grafico7[df_grafico7["Escenario"].str.contains(escenario, na=False)]

    # Iterar por cada tipo de penetración dentro del escenario filtrado
    for tipo_penetracion in df_escenario["Tipo de penetración"].unique():
        df_tipo = df_escenario[df_escenario["Tipo de penetración"] == tipo_penetracion].sort_values("Año")
        axs[i].plot(df_tipo["Año"], df_tipo["Emisiones CO2 Total (ton)"],
                    marker='o', label=tipo_penetracion)

    axs[i].set_title(escenario)
    axs[i].set_xlabel("Año")
    axs[i].grid(True, linestyle='--', alpha=0.6)
    axs[i].legend()
    if i == 0:
        axs[i].set_ylabel("Emisiones CO2 Total (ton)")

plt.suptitle("Evolución de las emisiones CO2 totales (ton) para tres escenarios específicos")
plt.tight_layout()
#plt.show()
# --------------------------------------------------
# Gráfico 8: Comparativo acumulado de emisiones CO2 totales (2024-2028)
# Filtrado por elasticidad ingresada
# --------------------------------------------------
df_emisiones_acumulado = df_proyecciones[(df_proyecciones["Elasticidad"] == -1.87) &
                                         (df_proyecciones["Año"].between(2024, 2028))].groupby(
    ["Escenario", "Tipo de penetración"]
)["Emisiones CO2 Total (ton)"].sum().reset_index()

base_emisiones_total = ((ventas_bev_base * df_resumen["Emisiones BEV CO2 Año Base (ton)/Unidad"].iloc[0]) +
                        (ventas_no_bev_base * df_resumen["Emisiones Sin BEV CO2 Año Base (ton)/Unidad"].iloc[0])) * 5

escenarios_fiscales_emisiones = sorted(df_emisiones_acumulado["Escenario"].unique(), key=obtener_numero)
tipos_penetracion_emisiones = ["Pesimista", "Tendencial", "Acelerado"]
ancho_barra = 0.2
posiciones = np.arange(len(escenarios_fiscales_emisiones))

plt.figure(figsize=(14,7))
for i, tipo in enumerate(tipos_penetracion_emisiones):
    datos_tipo = []
    for esc in escenarios_fiscales_emisiones:
        valor = df_emisiones_acumulado[
            (df_emisiones_acumulado["Escenario"] == esc) &
            (df_emisiones_acumulado["Tipo de penetración"] == tipo)
        ]["Emisiones CO2 Total (ton)"].values[0]
        datos_tipo.append(valor / 1e3)  # Convertir a miles de ton
    plt.bar(posiciones + i * ancho_barra, datos_tipo, ancho_barra, label=tipo)

datos_base = [base_emisiones_total / 1e3] * len(escenarios_fiscales_emisiones)
plt.bar(posiciones + 3 * ancho_barra, datos_base, ancho_barra, label="Emisiones Base 2023 (5 años)", color='gray', alpha=0.6)

plt.xlabel("Escenario", fontsize=12)
plt.ylabel("Emisiones CO2 acumuladas (miles de ton)", fontsize=12)
plt.title(f"Comparación acumulada de emisiones CO2 totales (2024-2028)\nElasticidad = {-1.87}", fontsize=14)
plt.xticks(posiciones + ancho_barra, escenarios_fiscales_emisiones, rotation=45, ha='right')
plt.grid(axis='y', linestyle='--', alpha=0.7)
plt.legend(title="Tipo de penetración", loc='upper left', bbox_to_anchor=(1, 1))
plt.tight_layout(rect=[0, 0, 1, 1])
plt.savefig(r"C:\Users\emili\PycharmProjects\TesisUY\Gráficos\Proyección 5 años\comparacion acumulada emisiones CO2 filtrado.png", dpi=300)
#plt.show()

# --------------------------------------------------
# Generación de gráficos de consumo energético en GWh
# --------------------------------------------------

# Factor de conversión: 1 tep = 0.01163 GWh
factor_tep_to_GWh = 0.01163

def generar_grafico_GWh(columna, valor_base, titulo, nombre_archivo):
    plt.figure(figsize=(14, 7))
    for i, tipo in enumerate(tipos_penetracion):
        datos_tipo = [
            df_acumulado[
                (df_acumulado["Escenario"] == esc) &
                (df_acumulado["Tipo de penetración"] == tipo)
                ][columna].values[0] * factor_tep_to_GWh
            for esc in escenarios_fiscales
        ]
        plt.bar(posiciones + i * ancho_barra, datos_tipo, ancho_barra, label=tipo)

    # Barra adicional para el valor base (consumo del año base multiplicado por 5 años)
    datos_base = [valor_base * factor_tep_to_GWh] * len(escenarios_fiscales)
    plt.bar(posiciones + 3 * ancho_barra, datos_base, ancho_barra, label="Consumo Base 2023 (5 años)", color='gray',
            alpha=0.6)

    # Configuración visual
    plt.xlabel("Escenario", fontsize=12)
    plt.ylabel("Consumo acumulado (GWh)", fontsize=12)
    plt.title(f"{titulo}\nElasticidad = {elasticidad_objetivo}", fontsize=14)
    plt.xticks(posiciones + 1.5 * ancho_barra, escenarios_fiscales, rotation=45, ha='right')
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    plt.legend(title="Tipo de penetración", loc='upper left', bbox_to_anchor=(1, 1))
    plt.tight_layout(rect=[0, 0, 1, 1])
    plt.savefig(nombre_archivo, dpi=300)
    #plt.show()


# --- Gráfico 1: Consumo Energético Sin BEV en GWh ---
generar_grafico_GWh(
    columna="Consumo Energético Sin BEV (tep)",
    valor_base=consumo_base_sin_bev_5anios,
    titulo="Comparación acumulada del Consumo Energético Sin BEV (2024-2028) en GWh",
    nombre_archivo=r"C:\Users\emili\PycharmProjects\TesisUY\Gráficos\Proyección 5 años\comparacion consumo sin BEV GWh.png"
)


# --- Gráfico 2: Consumo Energético BEV en GWh ---
generar_grafico_GWh(
    columna="Consumo Energético BEV (tep)",
    valor_base=consumo_base_bev_5anios,
    titulo="Comparación acumulada del Consumo Energético BEV (2024-2028) en GWh",
    nombre_archivo=r"C:\Users\emili\PycharmProjects\TesisUY\Gráficos\Proyección 5 años\comparacion consumo BEV GWh.png"
)
# --- Gráfico 3: Consumo Energético Total en GWh ---
generar_grafico_GWh(
    columna="Consumo Energético Total (tep)",
    valor_base=consumo_base_total_5anios,
    titulo="Comparación acumulada del Consumo Energético Total (2024-2028) en GWh",
    nombre_archivo=r"C:\Users\emili\PycharmProjects\TesisUY\Gráficos\Proyección 5 años\comparacion consumo total GWh.png"
)

# --------------------------------------------------
# Pie Chart: Distribución de emisiones BEV vs Sin BEV
# Solo para los Escenarios 4, 5 y 6 y filtrado por la elasticidad ingresada,
# acumulando datos del período 2024-2028.
# Se crean 3 subplots, uno para cada escenario.
# --------------------------------------------------

# Filtrar el DataFrame para elasticidad deseada y el período 2024-2028
df_pie = df_proyecciones[(df_proyecciones["Elasticidad"] == -1.87) &
                         (df_proyecciones["Año"].between(2024, 2028))].copy()

# Extraer el número del escenario (asumiendo que aparece un dígito en el nombre)
df_pie["Numero_Escenario"] = df_pie["Escenario"].str.extract(r'(\d+)').astype(int)

# Filtrar solo para los escenarios 4, 5 y 6
df_pie = df_pie[df_pie["Numero_Escenario"].isin([4, 5, 6])]

# Agrupar por "Escenario" y sumar las emisiones BEV y Sin BEV
df_pie_grouped = df_pie.groupby("Escenario").agg({
    "Emisiones CO2 BEV (ton)": "sum",
    "Emisiones CO2 Sin BEV (ton)": "sum"
}).reset_index()

# Crear los subplots (pie charts) para cada escenario filtrado
scenarios_for_pie = df_pie_grouped["Escenario"].unique()
fig, axes = plt.subplots(1, len(scenarios_for_pie), figsize=(6*len(scenarios_for_pie), 6))
if len(scenarios_for_pie) == 1:
    axes = [axes]
for ax, esc in zip(axes, scenarios_for_pie):
    row = df_pie_grouped[df_pie_grouped["Escenario"] == esc].iloc[0]
    bev_emissions = row["Emisiones CO2 BEV (ton)"]
    sin_bev_emissions = row["Emisiones CO2 Sin BEV (ton)"]
    data = [bev_emissions, sin_bev_emissions]
    labels = ["BEV", "Sin BEV"]
    ax.pie(data, labels=labels, autopct='%1.1f%%', startangle=90)
    ax.set_title(f"Escenario {esc}")
plt.suptitle(f"Distribución de emisiones CO2 (2024-2028)\nElasticidad = {-1.87}")
plt.tight_layout()
plt.savefig(r"C:\Users\emili\PycharmProjects\TesisUY\Gráficos\Proyección 5 años\pie emisiones scenarios filtrado.png", dpi=300)
#plt.show()