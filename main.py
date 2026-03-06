from acciones_archivo.leer import leer_archivo
from acciones_archivo.guardar import guardar_por_hojas
if __name__ == "__main__": 
    archivo = 'archivos/Week9.xlsx'
    hoja = 'IPL - Cases'
    print(f"archivo seleccionado: {archivo}")
    print(f"hoja seleccionada: {hoja}")
    datos = leer_archivo(archivo, hoja) 
    guardar_por_hojas(datos, original_path=archivo)
        