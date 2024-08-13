import pandas as pd  # Para manejar los datos del archivo Excel
import numpy as np  # Para generar valores aleatorios
import openpyxl  # Para manejar la lectura y escritura de archivos Excel

# Este programa lee un archivo Excel con las calificaciones de los alumnos en varias materias,
# agrega una columna llamada 'Mat_Fisica' con valores aleatorios entre 0 y 100 con un decimal,
# ordena la tabla por el campo 'Nombre' y guarda el archivo modificado.
# Luego, vuelve a leer el archivo modificado y lo ordena nuevamente por el campo 'Nombre'
# para asegurarse de que esté ordenado correctamente, guardando el resultado final en un archivo Excel.

#  Aqui lee el archivo Excel original
archivo_excel = 'calificaciones_alumnos.xlsx'
df = pd.read_excel(archivo_excel, engine='openpyxl')  # Carga los datos del archivo Excel en un DataFrame

# Aqui genera valores aleatorios entre 0 y 100 con un decimal para la nueva columna 'Mat_Fisica'
df['Mat_Fisica'] = np.round(np.random.uniform(0, 100, len(df)), 1)  # Agrega la columna con los valores generados

# Aqui ordena la tabla por la columna 'Nombre'
df.sort_values(by='Nombre', inplace=True)  # Ordena el DataFrame por el campo 'Nombre'

# Aqui guardar el DataFrame modificado en un nuevo archivo Excel
archivo_modificado = 'calificaciones_alumnos_modificado.xlsx'
df.to_excel(archivo_modificado, index=False, engine='openpyxl')  # Guarda el DataFrame en un nuevo archivo Excel

# Aqui lee nuevamente el archivo modificado
df_modificado = pd.read_excel(archivo_modificado, engine='openpyxl')  # Carga el archivo modificado en un DataFrame

# En esta parte  nuevamente la tabla por la columna 'Nombre'
df_modificado.sort_values(by='Nombre', inplace=True)  # Asegura que el DataFrame esté ordenado correctamente por 'Nombre'

# Finalmente guarda el DataFrame ordenado en el mismo archivo Excel modificado
df_modificado.to_excel(archivo_modificado, index=False, engine='openpyxl')  # Sobrescribe el archivo modificado con el DataFrame ordenado

print("Se ha agregado la columna 'Mat_Fisica', ordenado la tabla por 'Nombre' y guardado el archivo modificado dos veces para asegurar el orden.")