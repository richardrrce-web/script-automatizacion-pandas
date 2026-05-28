import pandas as pd
from datetime import datetime

print("🔄 Iniciando proceso de automatización de datos...")

try:
    # 1. LEER LOS DATOS: Cargamos el archivo CSV que nos dio el cliente
    # (Pandas lo convierte automáticamente en una tabla organizada en memoria)
    df = pd.read_csv('datos_ventas.csv')
    
    # 2. FILTRAR LOS DATOS:
    # Condición A: El Estado de la venta debe ser 'Completado'
    # Condición B: El Total debe ser mayor o igual a 100
    ventas_filtradas = df[(df['Estado'] == 'Completado') & (df['Total'] >= 100)]
    
    # 3. PROCESAR / MODIFICAR: Añadimos una columna nueva con la fecha del reporte
    fecha_actual = datetime.now().strftime('%Y-%m-%d')
    ventas_filtradas['Fecha_Reporte'] = fecha_actual
    
    # 4. EXPORTAR EL RESULTADO: Creamos un archivo Excel nuevo e impecable
    nombre_archivo_salida = 'reporte_ventas_VIP.xlsx'
    
    # Exportamos a Excel sin el índice numérico de la izquierda
    ventas_filtradas.to_excel(nombre_archivo_salida, index=False, engine='openpyxl')
    
    print(f"✅ ¡Proceso completado con éxito!")
    print(f"📊 Se filtraron los datos y se generó el archivo: '{nombre_archivo_salida}'")

except FileNotFoundError:
    print("❌ Error: No se encontró el archivo 'datos_ventas.csv'. Asegúrate de que esté en la misma carpeta.")
except Exception as e:
    print(f"❌ Ocurrió un error inesperado: {e}")
    