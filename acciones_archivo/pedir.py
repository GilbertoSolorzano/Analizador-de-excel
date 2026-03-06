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