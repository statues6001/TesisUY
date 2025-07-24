import numpy as np
import matplotlib.pyplot as plt

# Puntos de referencia
x1, y1 = 133.2, 0.23         # (g/km, IMESI)
x2_base, y2_base = 258.08, 0.46   # sin +10%
x2_pen, y2_pen = 258.08, 0.56    # con +10%

# Cálculo de pendientes e interceptos
m_base = (y2_base - y1) / (x2_base - x1)
b_base = y1 - m_base * x1
m_pen = (y2_pen - y1) / (x2_pen - x1)
b_pen = y1 - m_pen * x1

# Dominio de CO2
x = np.linspace(0, 350, 200)
y_base = m_base * x + b_base
y_pen = m_pen * x + b_pen

# Punto de intersección
x_int = (b_pen - b_base) / (m_base - m_pen)

# Gráfico
plt.figure(figsize=(8, 5))
plt.plot(x, y_base * 100, label='Sin +10 %')      # convierte a %
plt.plot(x, y_pen * 100, label='Con +10 %')
plt.axvline(x_int, linestyle='--', label=f'Intersección ≈ {x_int:.1f} g/km')
# Puntos de referencia
plt.scatter([x1, x2_base], [y1*100, y2_base*100])
plt.scatter([x1, x2_pen], [y1*100, y2_pen*100])
plt.xlabel('CO$_2$ (g/km)')
plt.ylabel('Tasa IMESI (%)')
plt.title('Comparación de funciones IMESI vs CO$_2$')
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.savefig(r'C:\Users\emili\PycharmProjects\TesisUY\Gráficos\Comparación de pendientes.png', dpi=300)
plt.show()

# -----------------------------------------------
# Segunda gráfica: comparación Escenario 8 vs 11
# -----------------------------------------------

# Puntos de Escenario 8
x1_8, y1_8 = 133.2, 0.23
x2_8, y2_8 = 258.08, 0.66

# Puntos de Escenario 11
x1_11, y1_11 = 133.2, 0.28
x2_11, y2_11 = 258.08, 0.96

# Pendientes e intersecciones
m_8 = (y2_8 - y1_8) / (x2_8 - x1_8)
b_8 = y1_8 - m_8 * x1_8
m_11 = (y2_11 - y1_11) / (x2_11 - x1_11)
b_11 = y1_11 - m_11 * x1_11

x = np.linspace(0, 300, 200)
y_8 = m_8 * x + b_8
y_11 = m_11 * x + b_11

x_int2 = (b_11 - b_8) / (m_8 - m_11)

# Gráfico 2
plt.figure(figsize=(8, 5))
plt.plot(x, y_8 * 100, label='Escenario 8', linewidth=2, color='orange')
plt.plot(x, y_11 * 100, label='Escenario 11', linewidth=2, color='orangered')
plt.axvline(x_int2, linestyle='--', color='gray', label=f'Intersección ≈ {x_int2:.1f} g/km')
plt.scatter([x1_8, x2_8], [y1_8 * 100, y2_8 * 100], color='blue')
plt.scatter([x1_11, x2_11], [y1_11 * 100, y2_11 * 100], color='orange')
plt.xlabel('CO$_2$ (g/km)')
plt.ylabel('Tasa IMESI (\%)')
plt.title('Comparación de funciones IMESI vs CO$_2$ (Escenarios 8 y 11)')
plt.ylim(0, 100)
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.savefig(r'C:\Users\emili\PycharmProjects\TesisUY\Gráficos\Comparación Escenario 8 vs 11.png', dpi=300)
plt.show()


