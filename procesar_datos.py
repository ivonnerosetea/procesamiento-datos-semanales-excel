
import pandas as pd

def procesar_datos(file_path, sheet_name="Raw Data"):
    # 1. Carga los datos desde el archivo Excel, omitiendo la primera fila
    raw_data = pd.read_excel(file_path, sheet_name=sheet_name).iloc[1:].copy()
    
    # 2. Renombra las columnas para facilitar su uso
    raw_data.columns = ["Fecha de creación", "Argentina", "Brasil", "Chile", "Perú", "Subtotal"]
    
    # 3. Convierte la columna de fecha a formato datetime y elimina valores nulos
    raw_data["Fecha de creación"] = pd.to_datetime(raw_data["Fecha de creación"], errors="coerce")
    raw_data = raw_data.dropna(subset=["Fecha de creación"])

    # 4. Convierte las columnas numéricas a valores numéricos, manejando errores
    columnas_numericas = ["Argentina", "Brasil", "Chile", "Perú", "Subtotal"]
    for col in columnas_numericas:
        raw_data[col] = pd.to_numeric(raw_data[col], errors="coerce")
    
    # 5. Calcula el inicio de la semana para cada fecha (siempre lunes)
    raw_data["FECHA_SEMANA"] = raw_data["Fecha de creación"] - pd.to_timedelta(raw_data["Fecha de creación"].dt.weekday, unit='D')

    # 6. Agrupa por semana y suma los valores de las columnas numéricas
    weekly_data = raw_data.groupby("FECHA_SEMANA", sort=False)[columnas_numericas].sum().reset_index()
    
    # 7. Ordena las semanas de menor a mayor y da formato a la fecha
    weekly_data = weekly_data.sort_values(by="FECHA_SEMANA", ascending=True)
    weekly_data["FECHA_SEMANA"] = weekly_data["FECHA_SEMANA"].dt.strftime("%-m/%-d/%Y")

    # 8. Renombra las columnas para mayor claridad
    weekly_data.rename(columns={
        "Argentina": "FUERZA_VENTAS_ARGENTINA_COLORITO",
        "Brasil": "FUERZA_VENTAS_BRASIL_COLORITO",
        "Chile": "FUERZA_VENTAS_CHILE_COLORITO",
        "Perú": "FUERZA_VENTAS_PERU_COLORITO",
        "Subtotal": "FUERZA_VENTAS_TOTAL_COLORITO"
    }, inplace=True)
    
    # 9. Detecta registros duplicados y valores nulos
    duplicados = raw_data.duplicated().sum()
    valores_nulos = raw_data.isnull().sum().sum()
    
    # 10. Reemplaza valores nulos con 0
    raw_data.fillna(0, inplace=True)

    # 11. Genera alertas si hay registros duplicados o valores nulos
    alertas = []
    if duplicados:
        alertas.append(f"Se han identificado {duplicados} registros duplicados.")
    if valores_nulos:
        alertas.append(f"Se han encontrado {valores_nulos} valores nulos, reemplazados por 0.")
    
    return weekly_data, alertas

# 12. Define la ruta del archivo de entrada
file_path = "datos.xlsx"

# 13. Ejecuta la función de procesamiento de datos
resultado, alertas = procesar_datos(file_path)

# 14. Muestra alertas en la terminal
for alerta in alertas:
    print(alerta)
if not alertas:
    print("No se encontraron más problemas en los datos.")

# 15. Guarda los datos procesados en un nuevo archivo Excel
resultado.to_excel("resultado_semanal.xlsx", index=False)

print(resultado.to_string())



























