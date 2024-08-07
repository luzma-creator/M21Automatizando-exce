# Este programa lee un archivo Excel con las calificaciones de los alumnos en varias materias
# y genera una gráfica de barras para cada alumno, asegurándose de que las etiquetas en el eje X
# no se encimen.

import pandas as pd  # Para manejar los datos del archivo Excel
import matplotlib.pyplot as plt  # Para crear las gráficas
import matplotlib.ticker as ticker  # Para ajustar el formato de las etiquetas del eje X

# Leer el archivo Excel
archivo_excel = 'calificaciones_alumnos.xlsx'
df = pd.read_excel(archivo_excel, engine='openpyxl')

# Rellenar valores NaN con 0 o cualquier otro valor adecuado
df.fillna(0, inplace=True)

# Convertir todas las etiquetas del eje X a cadenas de texto
df.columns = df.columns.astype(str)

# Configuración del tamaño de la figura
plt.figure(figsize=(12, 8))

# Recorrer cada fila del DataFrame (cada alumno)
for index, row in df.iterrows():
    nombre = row['Nombre']
    calificaciones = row[1:]  # Seleccionar todas las columnas de calificaciones

    # Crear una gráfica de barras para cada alumno
    plt.bar(calificaciones.index.astype(str), calificaciones.values, label=nombre)

# Configurar el formato de las etiquetas del eje X para evitar que se encimen
plt.xticks(rotation=45, ha='right')

# Configurar un intervalo en las etiquetas del eje X para evitar la superposición
ax = plt.gca()
ax.xaxis.set_major_locator(ticker.MaxNLocator(nbins=15))

# Agregar título y etiquetas
plt.title('Calificaciones de Alumnos')
plt.xlabel('Materias')
plt.ylabel('Calificaciones')

# Agregar leyenda fuera de la gráfica
plt.legend(loc='upper left', bbox_to_anchor=(1, 1))

# Ajustar el diseño para que todo se vea bien
plt.tight_layout()

# Guardar la gráfica como una imagen
plt.savefig('calificaciones_alumnos.png')

# Mostrar la gráfica
plt.show()