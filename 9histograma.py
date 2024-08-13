import pandas as pd
import matplotlib.pyplot as plt

# Leer el archivo Excel
archivo_excel = 'calificaciones_alumnos.xlsx'
df = pd.read_excel(archivo_excel, engine='openpyxl')

# Eliminar la columna 'No' si existe
if 'No' in df.columns:
    df = df.drop(columns=['No'])

# Definir las materias de interés
materias_interes = ['Mat_CalculoIntegral', 'Mat_ProgramacionOOP', 'Mat_EstructuraDatos']

# Crear un histograma para cada materia de interés
for materia in materias_interes:
    plt.figure(figsize=(8, 6))  # Tamaño de la figura
    plt.hist(df[materia], bins=10, color='skyblue', edgecolor='black')  # Crear el histograma
    plt.title(f'Distribución de Calificaciones en {materia}')  # Título del gráfico
    plt.xlabel('Calificaciones')  # Etiqueta del eje X
    plt.ylabel('Número de Alumnos')  # Etiqueta del eje Y
    plt.grid(True)  # Mostrar la cuadrícula
    plt.show()  # Mostrar el gráfico

print("Se han generado los histogramas de distribución de calificaciones para cada materia.")
