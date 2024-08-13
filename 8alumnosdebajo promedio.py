import pandas as pd
import openpyxl

# Leer el archivo Excel
archivo_excel = 'calificaciones_alumnos.xlsx'
df = pd.read_excel(archivo_excel, engine='openpyxl')

# Eliminar la columna 'No' si existe
if 'No' in df.columns:
    df = df.drop(columns=['No'])

# Definir las materias de interés
materias_interes = ['Mat_CalculoIntegral', 'Mat_ProgramacionOOP', 'Mat_EstructuraDatos']

# Calcular el promedio de calificaciones por materia en las materias de interés
promedio_por_materia = df[materias_interes].mean()

# Filtrar alumnos que tengan calificaciones por debajo del promedio en cualquier materia de interés
# Para cada fila (alumno), verificamos si al menos una calificación está por debajo del promedio
filtro = (df[materias_interes] < promedio_por_materia).any(axis=1)

# Aplicar el filtro al DataFrame original
df_bajo_promedio = df[filtro]

# Guardar los resultados en un nuevo archivo Excel
df_bajo_promedio.to_excel('alumnos_bajo_promedio_materias_interes.xlsx', index=False, engine='openpyxl')

print("Se ha identificado a los alumnos cuyas calificaciones en 'Mat_CalculoIntegral', 'Mat_ProgramacionOOP', o 'Mat_EstructuraDatos' están por debajo del promedio de esas materias. Los resultados se han guardado en un nuevo archivo.")
