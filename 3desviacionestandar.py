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

# Mostrar las columnas que se están utilizando para calcular el promedio, mediana y desviación estándar
print("Columnas usadas para calcular el promedio, mediana y desviación estándar:", df_numerico.columns.tolist())

# Calcular el promedio de las calificaciones por alumno
df['Promedio'] = df_numerico.mean(axis=1)

# Calcular la mediana de las calificaciones por materia
mediana_por_materia = df_numerico.median()

# Calcular la desviación estándar de las calificaciones por materia
desviacion_estandar_por_materia = df_numerico.std()

# Crear un DataFrame para las estadísticas de las materias
estadisticas_materias = pd.DataFrame({
    'Promedio': df_numerico.mean(),
    'Mediana': mediana_por_materia,
    'Desviación Estándar': desviacion_estandar_por_materia
})

# Guardar las estadísticas en un nuevo archivo Excel
estadisticas_materias.to_excel('estadisticas_calificaciones_por_materia.xlsx', engine='openpyxl')

print("Se ha calculado el promedio, la mediana y la desviación estándar por materia, y se ha guardado en un nuevo archivo.")
