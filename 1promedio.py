import pandas as pd
import openpyxl

# Leer el archivo Excel
archivo_excel = 'calificaciones_alumnos.xlsx'
df = pd.read_excel(archivo_excel, engine='openpyxl')

# Eliminar la columna 'No' del DataFrame si existe
if 'No' in df.columns:
    df = df.drop(columns=['No'])

# Filtrar las columnas numéricas (que contienen las calificaciones)
df_numerico = df.select_dtypes(include=['number'])

# Mostrar las columnas que se están utilizando para calcular el promedio
print("Columnas usadas para calcular el promedio:", df_numerico.columns.tolist())

# Calcular el promedio de las calificaciones por alumno
df['Promedio'] = df_numerico.mean(axis=1)

# Guardar el archivo modificado
df.to_excel('calificaciones_alumnos_promedio.xlsx', index=False, engine='openpyxl')

print("Se ha calculado el promedio de calificaciones por alumno y guardado en un nuevo archivo.")