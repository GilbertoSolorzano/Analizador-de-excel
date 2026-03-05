import pandas as pd 
import os
import openpyxl

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


def sanitize_sheet_name(name: str) -> str:
    # Excel limita los nombres de hoja a 31 caracteres y no permite algunos caracteres
    invalid = ['\\', '/', '*', '[', ']', ':', '?']
    for c in invalid:
        name = name.replace(c, '')
    return name[:31]

def guardar_filtros_en_hojas(datos: pd.DataFrame, save_path: str = 'nuevoFiltrado.xlsx'):
    
    #Aplica varios filtros y guarda cada resultado en una hoja distinta del mismo archivo Excel.
    
    # Define tus filtros: (nombre_hoja, columna, lista_de_valores)
    filtros = [
        ('Factory_Ensenada_Sauzal_Olathe', 'Factory', ['Ensenada', 'El Sauzal', 'Olathe']),
        ('Schlage_Residential_Mechanical', 'Brand / Category', ['Schlage Residential Mechanical']),
        ('Schlage_Residential_Electronic', 'Brand / Category', ['Schlage Residential Electronic']),
        ('Schlage_Electronic_Locks', 'Brand / Category', ['Schlage Electronic Locks']),
        ('Falcon_Lock', 'Brand / Category', ['Falcon - Lock']),
        ('Schlage_Commercial', 'Brand / Category', ['Schlage Commercial'])
    ]

    with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
        for nombre, col, valores in filtros:
            sheet_name = sanitize_sheet_name(nombre)
            # Comprueba que columna exista
            if col not in datos.columns:
                print(f"Advertencia: la columna '{col}' no existe en el DataFrame. Hoja '{sheet_name}' vacía.")
                # opcional: escribir un DataFrame vacío o con aviso
                pd.DataFrame({'Aviso': [f"Columna '{col}' no encontrada"]}).to_excel(writer, sheet_name=sheet_name, index=False)
                continue

            # Filtrado (maneja valores nulos sin error)
            mask = datos[col].isin(valores)
            df_filtrado = datos[mask].copy()

            if df_filtrado.empty:
                print(f"No se encontraron filas para {nombre}. Se escribirá hoja vacía con mensaje.")
                pd.DataFrame({'Aviso': [f"No se encontraron filas para filtro: {valores}"]}).to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                df_filtrado.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"Hoja '{sheet_name}' guardada con {len(df_filtrado)} filas.")

    print(f"Archivo guardado en: {save_path}")
#generar tabla 1
#def generar_tabla1(archivo_limpio):


if __name__ == "__main__": 
    archivo = pedir_archivo() 
    hoja = pedir_hoja(archivo)
    print(f"archivo seleccionado: {archivo}")
    print(f"hoja seleccionada: {hoja}")
    datos = leer_archivo(archivo, hoja) 

    if datos is None:
        print("No se pudo leer el DataFrame. Saliendo.")
    else:
        guardar_filtros_en_hojas(datos, save_path='nuevoFiltrado.xlsx')
        