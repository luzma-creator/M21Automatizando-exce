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

# Calcular el promedio de calificaciones para cada materia
promedios = df[materias_interes].mean()

# Filtrar a los alumnos cuyas calificaciones en todas las materias sean mayores de 90
df_filtrado = df.copy()
for materia in materias_interes:
    df_filtrado = df_filtrado[df_filtrado[materia] > 85]

# Verificar que las columnas 'NombreAlumno' y 'ApellidoAlumno' están presentes
if 'NombreAlumno' in df.columns and 'ApellidoAlumno' in df.columns:
    # Seleccionar las columnas relevantes, incluyendo nombre y apellido del alumno
    df_filtrado = df_filtrado[['NombreAlumno', 'ApellidoAlumno'] + materias_interes]
else:
    raise ValueError("Las columnas 'NombreAlumno' y 'ApellidoAlumno' no se encuentran en el archivo.")

# Agregar detalles de los alumnos
df_filtrado['Promedio_Calificaciones'] = df_filtrado[materias_interes].mean(axis=1)

# Guardar los resultados en un nuevo archivo Excel
archivo_salida = 'alumnos_con_calificaciones_mayores_85.xlsx'
df_filtrado.to_excel(archivo_salida, index=False, engine='openpyxl')

# Contar el número de alumnos que cumplen la condición
num_alumnos = len(df_filtrado)

# Mostrar el resultado
print(f"Número de alumnos con calificaciones mayores de 90 en todas las materias: {num_alumnos}")
print(f"Detalles de los alumnos han sido guardados en el archivo: {archivo_salida}")
