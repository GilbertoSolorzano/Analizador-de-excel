from acciones_archivo.leer import leer_archivo
from acciones_archivo.guardar import guardar_por_hojas
from acciones_archivo.pedir import pedir_archivo
from dotenv import load_dotenv
if __name__ == "__main__": 
    archivo = pedir_archivo()
    if not archivo:
        sys.exit()
    hoja = 'IPL - Cases'
    print(f"archivo seleccionado: {archivo}")
    print(f"hoja seleccionada: {hoja}")
    datos = leer_archivo(archivo, hoja) 
    guardar_por_hojas(datos, original_path=archivo)
        