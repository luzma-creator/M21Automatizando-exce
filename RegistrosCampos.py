import pandas as pd  # Para manejar los datos del archivo Excel
import numpy as np  # Para generar valores aleatorios
import openpyxl  # Para manejar la lectura y escritura de archivos Excel

# Este programa lee un archivo Excel con las calificaciones de los alumnos en varias materias,
# agrega una columna llamada 'Mat_Fisica' con valores aleatorios entre 0 y 100 con un decimal,
# ordena la tabla por el campo 'Nombre', cuenta la cantidad de registros y campos en la tabla,
# cuenta de la tabla cuantos registros tiene y cuantos campos desplegando el dato en la consola,
# y guarda el archivo modificado.

# Leer el archivo Excel original
archivo_excel = 'calificaciones_alumnos.xlsx'
df = pd.read_excel(archivo_excel, engine='openpyxl')  # Carga los datos del archivo Excel en un DataFrame

# Generar valores aleatorios entre 0 y 100 con un decimal para la nueva columna 'Mat_Fisica'
df['Mat_Fisica'] = np.round(np.random.uniform(0, 100, len(df)), 1)  # Agrega la columna con los valores generados

# Ordenar la tabla por la columna 'Nombre'
df.sort_values(by='Nombre', inplace=True)  # Ordena el DataFrame por el campo 'Nombre'

# Contar el número de registros (filas) y campos (columnas)
num_registros = len(df)  # Número de filas en el DataFrame
num_campos = len(df.columns)  # Número de columnas en el DataFrame

print(f"La tabla tiene {num_registros} registros y {num_campos} campos.")  # Muestra la cantidad de registros y campos

# Guardar el DataFrame modificado en un nuevo archivo Excel
archivo_modificado = 'calificaciones_alumnos_modificado.xlsx'
df.to_excel(archivo_modificado, index=False, engine='openpyxl')  # Guarda el DataFrame en un nuevo archivo Excel

print("Se ha agregado la columna 'Mat_Fisica', ordenado la tabla por 'Nombre' y guardado el archivo modificado.")