import pandas as pd
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

def tabla_5(df_filtrado: pd.DataFrame, writer, sheet_name: str, col_serie, col_reason, col_detail_reason, col_case, col_qty, startrow=0, startcol=5):
    # Preparar resumen
    resumen = df_filtrado.loc[:, [col_serie, col_reason, col_detail_reason, col_case, col_qty]].copy()
    resumen.columns = ['serie', 'reason', 'defect', 'case_number', 'quantity']
    resumen['quantity'] = pd.to_numeric(resumen['quantity'], errors='coerce').fillna(0)

    wb = writer.book
    ws = wb[sheet_name]
    
    # Obtener combinaciones únicas de brand + serie
    totales_grupo = (
        resumen
        .groupby(['serie', 'reason'], dropna=False, as_index=False)
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
    
    fill_serie = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
    fill_reason= PatternFill(start_color="00B0F0", end_color="00B0F0", fill_type="solid")
    fill_header = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
    
    font_bold = Font(bold=True, size=11)
    font_header = Font(bold=True, size=10)
    center = Alignment(horizontal='center')
    
    tabla_count = 0
    current_col = startcol  # Columna inicial (E = 5)
    max_row_used = startrow  # Para trackear la fila máxima usada
    
    for _, grupo_row in totales_grupo.iterrows():
        serie = grupo_row['serie']
        reason = grupo_row['reason']
        
        defectos = (
            resumen[(resumen['serie'] == serie) & (resumen['reason'] == reason)]
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
        current_row = startrow  # Resetear fila para cada tabla
        col_start = current_col
        col_end = current_col + 2  # 3 columnas
        
        # === FILA 1: Brand Category ===
        ws.merge_cells(start_row=current_row + 1, start_column=col_start, end_row=current_row + 1, end_column=col_end)
        cell_serie = ws.cell(row=current_row + 1, column=col_start, value=serie)
        cell_serie.font = font_bold
        cell_serie.fill = fill_serie
        cell_serie.alignment = center
        cell_serie.border = border_thin
        for col in range(col_start, col_end + 1):
            ws.cell(row=current_row + 1, column=col).border = border_thin
        current_row += 1
        
        # === FILA 2: Serie ===
        ws.merge_cells(start_row=current_row + 1, start_column=col_start, end_row=current_row + 1, end_column=col_end)
        cell_reason = ws.cell(row=current_row + 1, column=col_start, value=reason)
        cell_reason.font = font_bold
        cell_reason.fill = fill_reason
        cell_reason.alignment = center
        cell_reason.border = border_thin
        for col in range(col_start, col_end + 1):
            ws.cell(row=current_row + 1, column=col).border = border_thin
        current_row += 1
        
        # === FILA 3: Encabezados ===
        headers = ['Defect', 'CAS Qty', 'Qty Rejected']
        for i, header in enumerate(headers):
            cell = ws.cell(row=current_row + 1, column=col_start + i, value=header)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = center
            cell.border = border_thin
        current_row += 1
        
        # === FILAS DE DATOS ===
        for _, row in defectos.iterrows():
            ws.cell(row=current_row + 1, column=col_start, value=row['defect']).border = border_thin
            
            cell_cas = ws.cell(row=current_row + 1, column=col_start + 1, value=int(row['cas_qty']))
            cell_cas.alignment = center
            cell_cas.border = border_thin
            
            cell_qty = ws.cell(row=current_row + 1, column=col_start + 2, value=int(row['qty_rejected']))
            cell_qty.alignment = center
            cell_qty.border = border_thin
            
            current_row += 1
        
    
        '''# === FILA TOTAL ===
        total_cas = defectos['cas_qty'].sum()
        total_qty = defectos['qty_rejected'].sum()
        
        cell_total = ws.cell(row=current_row + 1, column=col_start, value="Total")
        cell_total.font = font_bold
        cell_total.border = border_thin
        
        cell_total_cas = ws.cell(row=current_row + 1, column=col_start + 1, value=int(total_cas))
        cell_total_cas.font = font_bold
        cell_total_cas.alignment = center
        cell_total_cas.border = border_thin
        
        cell_total_qty = ws.cell(row=current_row + 1, column=col_start + 2, value=int(total_qty))
        cell_total_qty.font = font_bold
        cell_total_qty.alignment = center
        cell_total_qty.border = border_thin
        
        current_row += 1'''
        
        # Actualizar fila máxima usada
        if current_row > max_row_used:
            max_row_used = current_row
        
        # Mover a la siguiente columna (espacio de 1 columna entre tablas)
        current_col += 4  # 3 columnas de datos + 1 de espacio

    print(f"Hoja '{sheet_name}': {tabla_count} tablas de defectos creadas (horizontal).")
    return max_row_used + 2