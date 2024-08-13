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

# Mostrar las columnas que se están utilizando para calcular las estadísticas
print("Columnas usadas para calcular las estadísticas:", df_numerico.columns.tolist())

# Calcular el promedio de las calificaciones por materia
promedio_por_materia = df_numerico.mean()

# Calcular la mediana de las calificaciones por materia
mediana_por_materia = df_numerico.median()

# Calcular la desviación estándar de las calificaciones por materia
desviacion_estandar_por_materia = df_numerico.std()

# Calcular la calificación máxima de cada materia
maxima_por_materia = df_numerico.max()

# Calcular la calificación mínima de cada materia
minima_por_materia = df_numerico.min()

# Contar calificaciones aprobatorias y reprobatorias por materia
