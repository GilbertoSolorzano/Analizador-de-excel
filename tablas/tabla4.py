import pandas as pd
from tablas.m_tabla import _make_table 
def tabla_4(df_filtrado: pd.DataFrame, writer, sheet_name: str, col_customer, col_case, startrow=0):
    # Preparar resumen solo con customer y case
    resumen = df_filtrado.loc[:, [col_customer, col_case]].copy()
    resumen.columns = ['customer', 'case_of_number']

    tabla_por_customer = (
        resumen
        .groupby('customer', as_index=False)
        .agg(total_de_casos=('case_of_number', 'nunique'))
    )

    # Escribir a partir de startrow (no en 0)
    tabla_por_customer.to_excel(writer, sheet_name=sheet_name, index=False, startrow=startrow)

    wb = writer.book
    ws = wb[sheet_name]

    _make_table(ws, tabla_por_customer, startrow, "por_customer")
    print(f"Hoja '{sheet_name}': {len(tabla_por_customer)} clientes escritos.")
    return startrow + len(tabla_por_customer) + 2