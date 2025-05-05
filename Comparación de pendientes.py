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
