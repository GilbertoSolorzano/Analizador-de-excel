import pandas as pd 
import os

def pedir_archivo():
    while True:
        nombre = input("Ingrese el nombre del archivo Excel (con extensión .xlsx o .xls): ").strip()
        if not (nombre.lower().endswith(".xlsx") or nombre.lower().endswith(".xls")):
            print("El archivo debe tener extensión .xlsx o .xls.")
            continue
        if not os.path.isfile(nombre):
            print(f" El archivo '{nombre}' no existe en la carpeta actual")
            continue
        return nombre
    
def pedir_hoja(archivo):
    try:
        xls = pd.ExcelFile(archivo)
        hojas = xls.sheet_names
    except Exception as e:
        print(f"No se pudo leer el archivo: {e}")
        return None

    hojas_lower = [h.lower() for h in hojas]
    while True:
        nombre_hoja = input("Ingresa el nombre de la hoja: ").strip()
        if nombre_hoja.lower() in hojas_lower:
            return hojas[hojas_lower.index(nombre_hoja.lower())]
        print(f"La hoja '{nombre_hoja}' no existe. Hojas disponibles: {', '.join(hojas)}")
    
def leer_archivo(archivo, hoja): 
    try:
        datos = pd.read_excel(archivo, sheet_name=hoja, header=6, engine='openpyxl') 
        return datos
    except FileNotFoundError as e:
        print(f"Error: archivo no encontrado: {e}")
    except ValueError as e:
        print(f"Error leyendo la hoja: {e}")
    except Exception as e:
        print(f"Ocurrió un error inesperado: {e}")
    return None   

def archivo_filtrado_factory_datos(datos, save_path=None, col='Factory', factories=None):
    if factories is None:
        factories = ['Ensenada', 'El Sauzal', 'Olathe']

    if datos is None:
        raise ValueError("El DataFrame es None. Asegúrate de haber leído el archivo correctamente.")
    if col not in datos.columns:
        raise KeyError(f"La columna '{col}' no existe en el DataFrame.")

    valores_norm = [v.strip().lower() for v in factories]
    mask = datos[col].astype(str).str.strip().str.lower().isin(valores_norm)
    datos_filtrado = datos.loc[mask].copy()

    if save_path:
        if not save_path.lower().endswith('.xlsx'):
            save_path += '.xlsx'
        datos_filtrado.to_excel(save_path, index=False)
        return datos_filtrado, save_path
    print("HOLA")
    return datos_filtrado

if __name__ == "__main__": 
    archivo = pedir_archivo() 
    hoja = pedir_hoja(archivo)
    print(f"archivo seleccionado: {archivo}")
    print(f"hoja seleccionada: {hoja}")

    datos = leer_archivo(archivo, hoja) 

    if datos is None:
        print("No se pudo leer el DataFrame. Saliendo.")
    else:
        #datos_filtrado = archivo_filtrado_factory_datos(datos)
        datos_filtrado, ruta = archivo_filtrado_factory_datos(datos, save_path="datos_filtrados2.xlsx")
        print(f"Archivo guardado en: {ruta}")