import pandas as pd
import openpyxl
import locale
from datetime import datetime

# Configura la localización para los nombres de los meses en español
try:
    locale.setlocale(locale.LC_TIME, "es_ES.utf8")  # En Windows podrías necesitar "Spanish_Spain.1252"
except locale.Error:
    print("Configuración regional no soportada. Usando traducción manual para los meses.")

def obtener_mes_actual():
    """Devuelve el nombre del mes actual en español."""
    mes_actual = datetime.now().strftime('%B')  # Obtiene el nombre del mes actual
    return mes_actual.capitalize()  # Devuelve el mes con la primera letra en mayúscula

# Asignar el mes actual a una variable
mes_en_curso = obtener_mes_actual()
print(mes_en_curso)

# Configuración inicial
csv_file = r"F:\python\projectos\Simulacion Power BI\CSV_Septiembre_2024.csv"  # Archivo CSV de entrada
excel_file = r"F:\python\projectos\Simulacion Power BI\Septiembre_2024.xlsx"  # Archivo Excel de salida
nuevo_excel_file = r"F:\python\projectos\Simulacion Power BI\Reporte de Eventos Banco Rioja.xlsx"  # Archivo Excel para los datos extraídos

# Leer el archivo CSV
df = pd.read_csv(csv_file)

# Reemplazar 'CrÃ­-tico' por 'Critico' en todas las celdas
df.replace('CrÃ­-tico', 'Critico', inplace=True)

# Separar la columna 'Time' por espacio y asignar nombres a las columnas resultantes
if 'Time' in df.columns:
    time_split = df['Time'].str.split(' ', expand=True)
    df['Fecha'] = time_split[0]  # Asignar la primera parte a 'Fecha'
    df.drop(columns=['Time'], inplace=True)  # Eliminar la columna 'Time'

# Cambiar el nombre de la columna 'Problem' y moverla a la columna J
if 'Problem' in df.columns:
    problem_data = df['Problem']
    df.drop(columns=['Problem'], inplace=True)
    df.insert(len(df.columns), 'Problem', problem_data)

# Separar la columna 'Host' por el delimitador '/' y eliminar la original
if 'Host' in df.columns:
    host_split = df['Host'].str.split('/', expand=True)
    df['Host'] = host_split[1]  # Usar la primera parte como nueva columna


# Convertir la columna 'Fecha' al formato de fecha corta
df['Fecha'] = pd.to_datetime(df['Fecha'], format='%Y-%m-%d').dt.strftime('%d-%m-%Y')

# Crear la columna 'Meses' con los nombres de los meses en español
try:
    df['Meses'] = pd.to_datetime(df['Fecha'], format='%d-%m-%Y').dt.strftime('%B').str.capitalize()
except locale.Error:
    # Traducción manual si la configuración regional no está disponible
    meses = {
        "January": "Enero", "February": "Febrero", "March": "Marzo", "April": "Abril",
        "May": "Mayo", "June": "Junio", "July": "Julio", "August": "Agosto",
        "September": "Septiembre", "October": "Octubre", "November": "Noviembre", "December": "Diciembre"
    }
    df['Meses'] = pd.to_datetime(df['Fecha'], format='%d-%m-%Y').dt.strftime('%B')
    df['Meses'] = df['Meses'].map(meses)

df.drop(columns=['Recovery time'], inplace=True)  # Eliminar la columna 'Time'
columnas = list(df.columns)  # Lista de las columnas actuales
columnas.remove('Fecha')  # Elimina 'Fecha' de su posición actual
columnas.insert(1, 'Fecha')  # Inserta 'Fecha' en la posición deseada (columna B)
df = df[columnas]  # Reordena el DataFrame según la nueva lista de columnas


# Guardar el resultado en un archivo Excel
df.to_excel(excel_file, index=False)


# Leer las filas desde la fila 2 hacia abajo
datos_desde_fila_2 = df.iloc[1:]  # Selecciona todas las filas desde el índice 1 hacia abajo

# Verificar si el archivo Excel ya existe
try:
    # Cargar el archivo Excel existente
    with pd.ExcelWriter(nuevo_excel_file, mode='a', if_sheet_exists='overlay', engine='openpyxl') as writer:
        # Leer el archivo existente
        existing_df = pd.read_excel(nuevo_excel_file)
        # Concatenar los datos nuevos al archivo existente
        datos_actualizados = pd.concat([existing_df, datos_desde_fila_2], ignore_index=True)
        # Guardar los datos actualizados en el archivo
        datos_actualizados.to_excel(writer, index=False, sheet_name='Eventos Banco Rioja 2024')
except FileNotFoundError:
    # Si el archivo no existe, crea uno nuevo con los datos desde la fila 2
    datos_desde_fila_2.to_excel(nuevo_excel_file, index=False, sheet_name='Eventos Banco Rioja 2024')
