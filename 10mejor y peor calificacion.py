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

# Crear un DataFrame para almacenar los resultados
mejor_peor_calificacion = pd.DataFrame()

# Iterar sobre cada materia para encontrar la mejor y peor calificación
for materia in materias_interes:
    # Identificar la mejor calificación y los alumnos correspondientes
    mejor_calificacion = df[materia].max()
    alumnos_mejor_calificacion = df[df[materia] == mejor_calificacion]

    # Identificar la peor calificación y los alumnos correspondientes
    peor_calificacion = df[materia].min()
    alumnos_peor_calificacion = df[df[materia] == peor_calificacion]

    # Agregar resultados al DataFrame, indicando si es la mejor o peor calificación
    alumnos_mejor_calificacion['Resultado'] = f'Mejor calificación en {materia}'
    alumnos_peor_calificacion['Resultado'] = f'Peor calificación en {materia}'

    # Concatenar los resultados en el DataFrame final
    mejor_peor_calificacion = pd.concat(
        [mejor_peor_calificacion, alumnos_mejor_calificacion, alumnos_peor_calificacion])

# Guardar los resultados en un nuevo archivo Excel
mejor_peor_calificacion.to_excel('mejor_peor_calificacion_por_materia.xlsx', index=False, engine='openpyxl')

print(
    "Se han identificado los alumnos con la mejor y peor calificación en cada materia. Los resultados se han guardado en un nuevo archivo.")