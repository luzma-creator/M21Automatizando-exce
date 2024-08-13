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

# Definir el rango de calificaciones
rango_minimo = 91
rango_maximo = 100

# Filtrar las calificaciones que están dentro del rango en las materias de interés
df_filtrado = pd.DataFrame()
for materia in materias_interes:
    df_temp = df[(df[materia] >= rango_minimo) & (df[materia] <= rango_maximo)]
    df_filtrado = pd.concat([df_filtrado, df_temp], axis=0)

# Eliminar duplicados en caso de que un alumno aparezca en múltiples filtros
df_filtrado = df_filtrado.drop_duplicates()

# Contar el número de alumnos con calificaciones dentro del rango
num_alumnos = len(df_filtrado)

# Mostrar el resultado
print(f"Número de alumnos con calificaciones dentro del rango de {rango_minimo} a {rango_maximo} en al menos una materia: {num_alumnos}")
