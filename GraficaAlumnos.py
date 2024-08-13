import pandas as pd
import matplotlib.pyplot as plt

# Leer el archivo Excel
archivo_excel = 'calificaciones_alumnos.xlsx'
df = pd.read_excel(archivo_excel, engine='openpyxl')

# Mostrar las primeras filas del DataFrame para revisar su estructura
print("Primeras filas del DataFrame:")
print(df.head())

# Rellenar valores faltantes con 0
df.fillna(0, inplace=True)

# Asegurarse de que las etiquetas (materias) sean palabras
df.columns = df.columns.astype(str)

# Mostrar las columnas del DataFrame para revisar los nombres y la posición
print("\nNombres de las columnas:")
print(df.columns)

# Identificar las columnas de calificaciones automáticamente
# Se asume que las columnas de calificaciones tienen nombres que contienen la palabra 'Mat' o que empiezan después de las columnas como 'Nombre'
calificaciones = df.loc[:, df.columns.str.contains('Mat')]

# Mostrar las primeras filas de las calificaciones seleccionadas para revisar
print("\nPrimeras filas de las calificaciones seleccionadas:")
print(calificaciones.head())

# Verificar si hay columnas seleccionadas correctamente
if calificaciones.empty:
    print("\nNo se encontraron columnas de calificaciones con los criterios dados.")
else:
    # Calcular el promedio de calificaciones para cada materia
    promedios = calificaciones.mean()

    # Mostrar los promedios calculados
    print("\nPromedios de calificaciones por materia:")
    print(promedios)

    # Configuración del tamaño de la figura
    plt.figure(figsize=(12, 8))

    # Crear la gráfica de barras horizontal con los promedios
    plt.barh(promedios.index, promedios.values)

    # Configurar las etiquetas del eje X y Y
    plt.xlabel('Promedio de Calificaciones')
    plt.ylabel('Materias')

    # Agregar título
    plt.title('Promedio de Calificaciones por Materia')

    # Ajustar el diseño
    plt.tight_layout()

    # Guardar y mostrar la gráfica
    plt.savefig('promedio_calificaciones_materias.png')
    plt.show()