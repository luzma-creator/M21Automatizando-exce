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

# Filtrar a los alumnos que tienen al menos una calificación menor de 70 en alguna materia
df_filtrado = df.copy()
df_filtrado = df_filtrado[df_filtrado[materias_interes].lt(70).any(axis=1)]

# Verificar que las columnas 'NombreAlumno' y 'ApellidoAlumno' están presentes
if 'NombreAlumno' in df.columns and 'ApellidoAlumno' in df.columns:
    # Seleccionar las columnas relevantes, incluyendo nombre y apellido del alumno
    df_filtrado = df_filtrado[['NombreAlumno', 'ApellidoAlumno'] + materias_interes]
else:
    raise ValueError("Las columnas 'NombreAlumno' y 'ApellidoAlumno' no se encuentran en el archivo.")

# Guardar los resultados en un nuevo archivo Excel
archivo_salida = 'alumnos_con_calificaciones_por_debajo_de_70.xlsx'
df_filtrado.to_excel(archivo_salida, index=False, engine='openpyxl')

# Contar el número de alumnos que cumplen la condición
num_alumnos = len(df_filtrado)

# Mostrar el resultado
print(f"Número de alumnos con calificaciones por debajo de 70 en alguna materia: {num_alumnos}")
print(f"Detalles de los alumnos han sido guardados en el archivo: {archivo_salida}")
