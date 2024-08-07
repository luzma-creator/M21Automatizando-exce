 Este programa lee un archivo Excel con las calificaciones de los alumnos en varias materias,
# agrega una columna llamada 'Mat_Fisica' con valores aleatorios entre 0 y 100 con un decimal,
# ordena la tabla por el campo 'Nombre' y guarda el archivo modificado.

import pandas as pd  # Para manejar los datos del archivo Excel
import numpy as np  # Para generar valores aleatorios
import openpyxl  # Para manejar la lectura y escritura de archivos Excel

# Leer el archivo Excel
archivo_excel = 'calificaciones_alumnos.xlsx'
df = pd.read_excel(archivo_excel, engine='openpyxl')

# Generar valores aleatorios entre 0 y 100 con un decimal para la nueva columna 'Mat_Fisica'
df['Mat_Fisica'] = np.round(np.random.uniform(0, 100, len(df)), 1)

# Ordenar la tabla por la columna 'Nombre'
df.sort_values(by='Nombre', inplace=True)

# Guardar el DataFrame modificado en un nuevo archivo Excel
archivo_modificado = 'calificaciones_alumnos_modificado.xlsx'
df.to_excel(archivo_modificado, index=False, engine='openpyxl')

print("Se ha agregado la columna 'Mat_Fisica', ordenado la tabla por 'Nombre' y guardado el archivo modificado.")