import pandas as pd
from tablas.m_tabla import _make_table 
#tabla 1
def tabla_1(df_filtrado: pd.DataFrame, writer, sheet_name: str, col_serie, col_case, col_qty):
    # Preparar resumen
    resumen = df_filtrado.loc[:, [col_serie, col_case, col_qty]].copy()
    resumen.columns = ['serie', 'case_of_numer', 'quality']
    resumen['quality'] = pd.to_numeric(resumen['quality'], errors='coerce').fillna(0)

    # Tabla detalle por serie + case (solo para calcular, no se escribe)
    tabla_resumen = (
        resumen
        .groupby(['serie', 'case_of_numer'], dropna=False, as_index=False)
        .agg({'quality': 'sum'})
        .rename(columns={'quality': 'sum_of_quality'})
    )

    # Tabla agrupada por serie
    tabla_por_serie = (
        tabla_resumen
        .groupby('serie', as_index=False)
        .agg(
            cantidad_de_casos=('case_of_numer', 'nunique'),
            suma_de_las_cantidades=('sum_of_quality', 'sum')
        )
    )

    # Escribir solo la tabla por serie
    tabla_por_serie.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0)

    # Dar formato de tabla de Excel
    wb = writer.book
    ws = wb[sheet_name]

    _make_table(ws, tabla_por_serie, 0, "por_serie")

    print(f"Hoja '{sheet_name}': {len(tabla_por_serie)} series escritas.")
    return len(tabla_por_serie) + 2  # retorna la siguiente fila disponible
