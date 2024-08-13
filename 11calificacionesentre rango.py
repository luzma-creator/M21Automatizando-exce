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
rango_minimo = 70
rango_maximo = 90

# Filtrar las calificaciones que están dentro del rango en las materias de interés
# Para cada materia, filtramos las calificaciones que están dentro del rango especificado
df_filtrado = pd.DataFrame()
for materia in materias_interes:
    df_temp = df[(df[materia] >= rango_minimo) & (df[materia] <= rango_maximo)]
    df_filtrado = pd.concat([df_filtrado, df_temp], axis=0)

# Eliminar duplicados en caso de que un alumno aparezca en múltiples filtros
df_filtrado = df_filtrado.drop_duplicates()

# Verificar que las columnas 'NombreAlumno' y 'ApellidoAlumno' están presentes
if 'NombreAlumno' in df.columns and 'ApellidoAlumno' in df.columns:
    # Seleccionar las columnas relevantes, incluyendo nombre y apellido del alumno
    df_filtrado = df_filtrado[['NombreAlumno', 'ApellidoAlumno'] + materias_interes]
else:
    raise ValueError("Las columnas 'NombreAlumno' y 'ApellidoAlumno' no se encuentran en el archivo.")

# Guardar los resultados en un nuevo archivo Excel
df_filtrado.to_excel('calificaciones_rango_70_90_con_nombres.xlsx', index=False, engine='openpyxl')

print(f"Se han filtrado y guardado las calificaciones que se encuentran dentro del rango de {rango_minimo} a {rango_maximo} en las materias de interés, incluyendo nombre y apellido del alumno.")
