import pandas as pd
import openpyxl
from tablas.m_tabla import autofit_columns
from tablas.tabla1 import tabla_1
from tablas.tabla2 import tabla_2
from tablas.tabla3 import tabla_3
from tablas.tabla4 import tabla_4
from acciones_archivo.buscar_columnas import match_column_by_keywords
from acciones_archivo.obtener_nombre import build_save_path
def guardar_por_hojas(datos: pd.DataFrame,  original_path: str):
    
    save_path = build_save_path(original_path, suffix='_filtrado', out_ext='.xlsx')
    #Aplica varios filtros y guarda cada resultado en una hoja distinta del mismo archivo Excel.
    
    # Define tus filtros: (nombre_hoja, columna, lista_de_valores)
    filtros = [
        #('Factory_Ensenada_Sauzal_Olathe', 'Factory', ['Ensenada', 'El Sauzal', 'Olathe']),
        ('Schlage_Residential_Mechanical', 'Brand / Category', ['Schlage Residential Mechanical']),
        ('Schlage_Residential_Electronic', 'Brand / Category', ['Schlage Residential Electronic']),
        ('Schlage_Electronic_Locks', 'Brand / Category', ['Schlage Electronic Locks']),
        ('Falcon_Lock', 'Brand / Category', ['Falcon - Lock']),
        ('Schlage_Commercial', 'Brand / Category', ['Schlage Commercial'])
    ]
    print("detecta filtros")
    with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
        for nombre, col, valores in filtros:
            sheet_name = nombre
            # Comprueba que columna exista
            if col not in datos.columns:
                print(f"Advertencia: la columna '{col}' no existe en el DataFrame. Hoja '{sheet_name}' vacía.")
                # opcional: f un DataFrame vacío o con aviso
                pd.DataFrame({'Aviso': [f"Columna '{col}' no encontrada"]}).to_excel(writer, sheet_name=sheet_name, index=False)
                continue

            # Filtrado (maneja valores nulos sin error)
            mask = datos[col].isin(valores)
            df_filtrado = datos[mask].copy()

            if df_filtrado.empty:
                print(f"No se encontraron filas para {nombre}. Se escribirá hoja vacía con mensaje.")
                pd.DataFrame({'Aviso': [f"No se encontraron filas para filtro: {valores}"]}).to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                # detecta columnas para el resumen: serie, case, cantidad
                col_serie = match_column_by_keywords(df_filtrado, ['serie', 'serial', 'serial number', 's/n'])
                col_case  = match_column_by_keywords(df_filtrado, ['case', 'case of', 'case number', 'case_of', 'case#'])
                col_qty   = match_column_by_keywords(df_filtrado, ['quality', 'qty', 'quantity', 'quiality', 'cant', 'count'])
                col_customer = match_column_by_keywords(df_filtrado, ['customer', 'cliente', 'client', 'cust'])
                col_reason = match_column_by_keywords(df_filtrado, ['reason (english)', 'razon', 'razón'])
                col_detail_reason = match_column_by_keywords(df_filtrado, ['detail reason (english)', 'detalle', 'detail'])

                if col_serie is None or col_case is None or col_qty is None:
                    # Si falta alguna columna, avisamos y escribimos el df completo como antes
                    print(f"Advertencia: no se detectaron las columnas necesarias para el resumen: serie={col_serie}, case={col_case}, qty={col_qty}")
                    df_filtrado.to_excel(writer, sheet_name=sheet_name, index=False)
                else:
                    next_row = tabla_1(df_filtrado, writer, sheet_name, col_serie, col_case, col_qty)

                    if col_reason is not None and col_detail_reason is not None:
                        next_row = tabla_2(df_filtrado, writer, sheet_name, col_serie, col_case, col_qty, col_reason, col_detail_reason, startrow=next_row)
                    if col_reason is not None and col_detail_reason is not None:
                        next_row = tabla_3(df_filtrado, writer, sheet_name, col_case, col_qty, col_reason, col_detail_reason, startrow=next_row)
                    if col_customer is not None:
                        tabla_4(df_filtrado, writer, sheet_name, col_customer, col_case, startrow=next_row)
                        
                    ws = writer.book[sheet_name]
                    autofit_columns(ws)
