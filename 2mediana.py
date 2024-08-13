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

# Mostrar las columnas que se están utilizando para calcular el promedio y la mediana
print("Columnas usadas para calcular el promedio y la mediana:", df_numerico.columns.tolist())

# Calcular el promedio de las calificaciones por alumno
df['Promedio'] = df_numerico.mean(axis=1)

# Calcular la mediana de las calificaciones por materia
mediana_por_materia = df_numerico.median()

# Agregar la mediana de cada materia al DataFrame como una fila nueva (opcional)
df_medianas = pd.DataFrame(mediana_por_materia).T
df_medianas.index = ['Mediana']
df = pd.concat([df, df_medianas], sort=False)

# Guardar el archivo modificado
df.to_excel('calificaciones_alumnos_promedio_mediana.xlsx', index=False, engine='openpyxl')

print("Se ha calculado el promedio de calificaciones por alumno y la mediana por materia, y se ha guardado en un nuevo archivo.")
