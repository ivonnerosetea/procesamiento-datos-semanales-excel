# 📊 Procesamiento de Datos Semanales en Excel

## Descripción
Este script en Python procesa datos de ventas semanales a partir de un archivo Excel. Extrae información de una hoja específica, limpia los datos, los agrupa por semanas y genera un nuevo archivo Excel con los resultados procesados, listos para su análisis.

## Requisitos Previos
Antes de ejecutar el script, asegúrate de tener instalado **Python 3** y las siguientes librerías:
- **pandas** (para manipulación de datos en Excel)
- **openpyxl** (para leer y escribir archivos Excel)

Puedes instalarlas con:
```sh
pip install pandas openpyxl
```

## Entrada y Salida
- **Entrada:** Un archivo Excel (`datos.xlsx`) con una hoja llamada `Raw Data`, que contiene datos de ventas diarias.
- **Salida:** Un nuevo archivo Excel (`resultado_semanal.xlsx`) con los datos procesados y organizados por semana.

## Funcionamiento del Código
1. **Carga los datos de Excel:**
   - Se lee la hoja `Raw Data`, omitiendo la primera fila que generalmente contiene encabezados repetidos o información no relevante.
   - Se renombran las columnas para facilitar su manejo en Python.

2. **Limpieza y preprocesamiento de datos:**
   - Convierte la columna de fechas (`Fecha de creación`) a formato `datetime` para facilitar su manipulación.
   - Elimina filas con fechas inválidas para evitar errores en el procesamiento.
   - Convierte columnas numéricas (`Argentina`, `Brasil`, `Chile`, `Perú`, `Subtotal`) a tipo numérico para realizar operaciones matemáticas correctamente.
   - **Elimina registros duplicados**, asegurando que no haya filas repetidas en el dataset.
   - **Reemplaza valores nulos con `0`** para evitar inconsistencias en los cálculos.

3. **Agrupación por Semana:**
   - Se calcula el inicio de cada semana (lunes) basándose en `Fecha de creación`.
   - Se ordenan las fechas de menor a mayor para mantener un orden cronológico correcto.
   - Se agrupan los datos sumando los valores por semana, asegurando que cada semana tenga sus datos consolidados.

4. **Formato de salida:**
   - Se cambia el formato de la fecha de la semana a `M/D/YYYY`, eliminando ceros innecesarios para mayor claridad.
   - Se renombran las columnas con nombres más descriptivos, facilitando la comprensión del reporte.
   - Se detectan registros duplicados y valores nulos, generando alertas si es necesario.

5. **Generación del archivo final:**
   - Se exporta el DataFrame procesado a un nuevo archivo Excel llamado `resultado_semanal.xlsx`.

## Alertas
El script muestra alertas en la terminal si:
- **Existen registros duplicados**, lo que puede afectar la precisión del análisis.
- **Hay valores nulos en los datos**, los cuales se reemplazan por `0` para evitar inconsistencias.

### Ejemplo de Salida en Terminal
```sh
Se han identificado 3 registros duplicados.
Se han encontrado 5 valores nulos, reemplazados por 0.
No se encontraron más problemas en los datos.
```

---

## Formato de Salida
| FECHA_SEMANA | FUERZA_VENTAS_ARGENTINA_COLORITO | FUERZA_VENTAS_BRASIL_COLORITO | FUERZA_VENTAS_CHILE_COLORITO | FUERZA_VENTAS_PERU_COLORITO | FUERZA_VENTAS_TOTAL_COLORITO |
|-------------|---------------------------------|--------------------------------|-------------------------------|------------------------------|------------------------------|
| 4/1/2019    | 722                             | 2016                           | 2270                          | 2287                         | 7295                         |
| 4/8/2019    | 349                             | 3598                           | 2390                          | 2041                         | 8378                         |
| 4/15/2019   | 417                             | 2788                           | 4195                          | 2249                         | 9649                         |

