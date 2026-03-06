import pandas as pd
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

def tabla_5(df_filtrado: pd.DataFrame, writer, sheet_name: str, col_brand, col_serie, col_detail_reason, col_case, col_qty, startrow=0):
    # Preparar resumen
    resumen = df_filtrado.loc[:, [col_brand, col_serie, col_detail_reason, col_case, col_qty]].copy()
    resumen.columns = ['brand', 'serie', 'defect', 'case_number', 'quantity']
    resumen['quantity'] = pd.to_numeric(resumen['quantity'], errors='coerce').fillna(0)

    wb = writer.book
    ws = wb[sheet_name]
    
    current_row = startrow
    
    # Obtener combinaciones únicas de brand + serie, ordenadas por cantidad
    totales_grupo = (
        resumen
        .groupby(['brand', 'serie'], dropna=False, as_index=False)
        .agg(total_qty=('quantity', 'sum'))
        .sort_values('total_qty', ascending=False)
    )
    
    # Estilos
    border_thin = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    fill_brand = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")  # Verde
    fill_serie = PatternFill(start_color="00B0F0", end_color="00B0F0", fill_type="solid")  # Azul
    fill_header = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")  # Azul claro
    
    font_bold = Font(bold=True, size=11)
    font_header = Font(bold=True, size=10)
    center = Alignment(horizontal='center')
    
    tabla_count = 0
    
    for _, grupo_row in totales_grupo.iterrows():
        brand = grupo_row['brand']
        serie = grupo_row['serie']
        
        # Filtrar defectos de este brand + serie
        defectos = (
            resumen[(resumen['brand'] == brand) & (resumen['serie'] == serie)]
            .groupby('defect', dropna=False, as_index=False)
            .agg(
                cas_qty=('case_number', 'nunique'),
                qty_rejected=('quantity', 'sum')
            )
            .sort_values('qty_rejected', ascending=False)
        )
        
        if defectos.empty:
            continue
        
        tabla_count += 1
        
        # === FILA 1: Brand Category ===
        ws.merge_cells(start_row=current_row + 1, start_column=1, end_row=current_row + 1, end_column=3)
        cell_brand = ws.cell(row=current_row + 1, column=1, value=brand)
        cell_brand.font = font_bold
        cell_brand.fill = fill_brand
        cell_brand.alignment = center
        cell_brand.border = border_thin
        for col in range(2, 4):
            ws.cell(row=current_row + 1, column=col).border = border_thin
        current_row += 1
        
        # === FILA 2: Serie ===
        ws.merge_cells(start_row=current_row + 1, start_column=1, end_row=current_row + 1, end_column=3)
        cell_serie = ws.cell(row=current_row + 1, column=1, value=serie)
        cell_serie.font = font_bold
        cell_serie.fill = fill_serie
        cell_serie.alignment = center
        cell_serie.border = border_thin
        for col in range(2, 4):
            ws.cell(row=current_row + 1, column=col).border = border_thin
        current_row += 1
        
        # === FILA 3: Encabezados ===
        headers = ['Defect', 'CAS Qty', 'Qty Rejected']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=current_row + 1, column=col, value=header)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = center
            cell.border = border_thin
        current_row += 1
        
        # === FILAS DE DATOS ===
        for _, row in defectos.iterrows():
            ws.cell(row=current_row + 1, column=1, value=row['defect']).border = border_thin
            ws.cell(row=current_row + 1, column=2, value=int(row['cas_qty'])).border = border_thin
            ws.cell(row=current_row + 1, column=2).alignment = center
            ws.cell(row=current_row + 1, column=3, value=int(row['qty_rejected'])).border = border_thin
            ws.cell(row=current_row + 1, column=3).alignment = center
            current_row += 1
        
        # === FILA TOTAL ===
        total_cas = defectos['cas_qty'].sum()
        total_qty = defectos['qty_rejected'].sum()
        
        cell_total = ws.cell(row=current_row + 1, column=1, value="Total")
        cell_total.font = font_bold
        cell_total.border = border_thin
        
        cell_total_cas = ws.cell(row=current_row + 1, column=2, value=int(total_cas))
        cell_total_cas.font = font_bold
        cell_total_cas.alignment = center
        cell_total_cas.border = border_thin
        
        cell_total_qty = ws.cell(row=current_row + 1, column=3, value=int(total_qty))
        cell_total_qty.font = font_bold
        cell_total_qty.alignment = center
        cell_total_qty.border = border_thin
        
        current_row += 1
        
        # Espacio entre tablas
        current_row += 2

    print(f"Hoja '{sheet_name}': {tabla_count} tablas de defectos creadas.")
    return current_row + 2