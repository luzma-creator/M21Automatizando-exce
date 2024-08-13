import pandas as pd
import openpyxl
import matplotlib.pyplot as plt

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
aprobatorias = (df_numerico >= 70).sum()
reprobatorias = (df_numerico < 70).sum()

# Calcular el promedio global de todas las calificaciones
suma_total = df_numerico.sum().sum()
total_calificaciones = df_numerico.size
promedio_global = suma_total / total_calificaciones

# Crear un DataFrame para las estadísticas de las materias
estadisticas_materias = pd.DataFrame({
    'Promedio': promedio_por_materia,
    'Mediana': mediana_por_materia,
    'Desviación Estándar': desviacion_estandar_por_materia,
    'Máxima': maxima_por_materia,
    'Mínima': minima_por_materia,
    'Aprobatorias (>=70)': aprobatorias,
    'Reprobatorias (<70)': reprobatorias
})

# Guardar las estadísticas en un nuevo archivo Excel
estadisticas_materias.to_excel('estadisticas_calificaciones_completas.xlsx', engine='openpyxl')

# Generar un gráfico de barras con los promedios de calificaciones por materia
plt.figure(figsize=(10, 6))
promedio_por_materia.plot(kind='bar', color='skyblue')
plt.title('Promedio de Calificaciones por Materia')
plt.xlabel('Materia')
plt.ylabel('Promedio de Calificación')
plt.xticks(rotation=45, ha='right')
plt.tight_layout()  # Ajusta el diseño para que todo encaje bien
plt.savefig('promedio_calificaciones_por_materia.png')  # Guarda el gráfico como archivo PNG
plt.show()  # Muestra el gráfico

print("Se ha calculado el promedio, la mediana, la desviación estándar, la máxima, la mínima calificación, y el conteo de calificaciones aprobatorias y reprobatorias por materia. También se ha generado un gráfico de barras con los promedios de calificaciones por materia y se ha guardado en un archivo.")
