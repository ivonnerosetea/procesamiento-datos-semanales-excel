
import pandas as pd

def procesar_datos(file_path, sheet_name="Raw Data"):
    raw_data = pd.read_excel(file_path, sheet_name=sheet_name).iloc[1:].copy()
    
    raw_data.columns = ["Fecha de creación", "Argentina", "Brasil", "Chile", "Perú", "Subtotal"]
    
    raw_data["Fecha de creación"] = pd.to_datetime(raw_data["Fecha de creación"], errors="coerce")
    raw_data = raw_data.dropna(subset=["Fecha de creación"])

    columnas_numericas = ["Argentina", "Brasil", "Chile", "Perú", "Subtotal"]
    for col in columnas_numericas:
        raw_data[col] = pd.to_numeric(raw_data[col], errors="coerce")
    
    raw_data["FECHA_SEMANA"] = raw_data["Fecha de creación"] - pd.to_timedelta(raw_data["Fecha de creación"].dt.weekday, unit='D')

    raw_data = raw_data.sort_values(by="Fecha de creación", ascending=True)

    weekly_data = raw_data.groupby("FECHA_SEMANA", sort=False)[columnas_numericas].sum().reset_index()
    
    weekly_data = weekly_data.sort_values(by="FECHA_SEMANA", ascending=True)
    weekly_data["FECHA_SEMANA"] = weekly_data["FECHA_SEMANA"].dt.strftime("%-m/%-d/%Y")

    weekly_data.rename(columns={
        "Argentina": "FUERZA_VENTAS_ARGENTINA_COLORITO",
        "Brasil": "FUERZA_VENTAS_BRASIL_COLORITO",
        "Chile": "FUERZA_VENTAS_CHILE_COLORITO",
        "Perú": "FUERZA_VENTAS_PERU_COLORITO",
        "Subtotal": "FUERZA_VENTAS_TOTAL_COLORITO"
    }, inplace=True)
    
    duplicados = raw_data.duplicated().sum()
    valores_nulos = raw_data.isnull().sum().sum()
    
    raw_data.fillna(0, inplace=True)

    alertas = []
    if duplicados:
        alertas.append(f"Se han identificado {duplicados} registros duplicados.")
    if valores_nulos:
        alertas.append(f"Se han encontrado {valores_nulos} valores nulos, reemplazados por 0.")
    
    return weekly_data, alertas

file_path = "datos.xlsx"

resultado, alertas = procesar_datos(file_path)

for alerta in alertas:
    print(alerta)
if not alertas:
    print("No se encontraron alertas.")

resultado.to_excel("resultado_semanal.xlsx", index=False)

print(resultado.to_string())



























