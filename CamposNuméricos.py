# Este programa lee un archivo Excel y determina qué columnas contienen datos numéricos.
# Luego imprime los nombres de las columnas que son de tipo numérico.

import pandas as pd  # Para manejar los datos del archivo Excel

# Leer el archivo Excel
archivo_excel = 'calificaciones_alumnos.xlsx'
df = pd.read_excel(archivo_excel, engine='openpyxl')

# Identificar columnas numéricas
columnas_numericas = df.select_dtypes(include=['number']).columns

# Imprimir los nombres de las columnas que son de tipo numérico
print("Las siguientes columnas contienen datos numéricos:")
for columna in columnas_numericas:
    print(columna)