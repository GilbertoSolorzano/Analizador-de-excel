
import openpyxl
import pandas as pd 

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