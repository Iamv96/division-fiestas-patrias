import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def update_excel(file_path):
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)
    
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid") # Dark blue
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    bold_font = Font(name="Calibri", size=11, bold=True)
    regular_font = Font(name="Calibri", size=11)
    money_format = "$#,##0"
    
    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )
    
    # ----------------------------------------------------
    # HOJA 1: Compras
    # ----------------------------------------------------
    ws1 = wb.create_sheet(title="Compras")
    
    # Titulos
    ws1["A1"] = "Total primera compra (Ignacio)"
    ws1["B1"] = 122251
    ws1["A2"] = "Total segunda compra (Ignacio)"
    ws1["B2"] = 61141
    ws1["A3"] = "Total tercera compra (Mindy)"
    ws1["B3"] = 11790
    ws1["A4"] = "Total cuarta compra (Joaquín)"
    ws1["B4"] = 5570
    
    for r in range(1, 5):
        ws1[f"A{r}"].font = bold_font
        ws1[f"B{r}"].font = bold_font
        ws1[f"B{r}"].number_format = money_format
    
    items = [
        "Chorizo", "Lomo Liso", "Lomo vetado", "Costillar Chileno", "Costillar cerdo",
        "Trutro", "Pechuga", "Carbón", "Pan Choripán", "Trigo Mote (1kg)", "Huesillos (1kg)", "Chancaca y Azúcar", "Restante"
    ]
    
    c1_values = [14670, 21226, 19560, 19238, 14957, 0, 0, 0, 0, 2490, 10990, 0, 19120]
    c2_values = [4890, 0, 33693, 0, 0, 16353, 6205, 0, 0, 0, 0, 0, 0]
    c3_values = [0, 0, 0, 0, 0, 0, 0, 8400, 0, 0, 0, 3390, 0]
    c4_values = [0, 0, 0, 0, 0, 0, 0, 0, 5570, 0, 0, 0, 0]
    
    # Headers
    header_row_idx = 6
    ws1.cell(row=header_row_idx, column=1, value="Detalle").font = header_font
    ws1.cell(row=header_row_idx, column=1).fill = header_fill
    ws1.cell(row=header_row_idx, column=1).alignment = Alignment(horizontal="center")
    
    for col_idx, item in enumerate(items, start=2):
        cell = ws1.cell(row=header_row_idx, column=col_idx, value=item)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
    
    total_col_idx = len(items) + 2
    cell_tot = ws1.cell(row=header_row_idx, column=total_col_idx, value="Total Fila")
    cell_tot.font = header_font
    cell_tot.fill = header_fill
    cell_tot.alignment = Alignment(horizontal="center")
    
    # Rows
    rows_data = [
        ("Primera compra (Ignacio)", c1_values, "=B1"),
        ("Segunda compra (Ignacio)", c2_values, "=B2"),
        ("Tercera compra (Mindy)", c3_values, "=B3"),
        ("Cuarta compra (Joaquín)", c4_values, "=B4"),
    ]
    
    for r_offset, (label, vals, expected_ref) in enumerate(rows_data, start=7):
        ws1.cell(row=r_offset, column=1, value=label).font = bold_font
        for col_idx, val in enumerate(vals, start=2):
            c = ws1.cell(row=r_offset, column=col_idx, value=val)
            c.font = regular_font
            c.number_format = money_format
            c.border = thin_border
        
        # Total Fila formula
        end_letter = get_column_letter(len(items) + 1)
        tot_cell = ws1.cell(row=r_offset, column=total_col_idx, value=f"=SUM(B{r_offset}:{end_letter}{r_offset})")
        tot_cell.font = bold_font
        tot_cell.number_format = money_format
        tot_cell.border = thin_border
    
    # Total Columnas
    tot_row_idx = 11
    ws1.cell(row=tot_row_idx, column=1, value="Total").font = bold_font
    for col_idx in range(2, len(items) + 2):
        col_letter = get_column_letter(col_idx)
        c = ws1.cell(row=tot_row_idx, column=col_idx, value=f"=SUM({col_letter}7:{col_letter}10)")
        c.font = bold_font
        c.number_format = money_format
        c.border = thin_border
        
    tot_general = ws1.cell(row=tot_row_idx, column=total_col_idx, value=f"=SUM({get_column_letter(total_col_idx)}7:{get_column_letter(total_col_idx)}10)")
    tot_general.font = bold_font
    tot_general.number_format = money_format
    tot_general.fill = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")
    tot_general.border = thin_border
    
    # Auto-adjust column widths
    for col in ws1.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws1.column_dimensions[col_letter].width = max(max_len + 3, 13)
        
    # ----------------------------------------------------
    # HOJA 2: Consumo por Integrante
    # ----------------------------------------------------
    ws2 = wb.create_sheet(title="Consumo por Integrante")
    
    integrantes = [
        "Pamela", "Cote", "Joaquín", "Ignacio", "Kena",
        "Ale", "Mauri", "Mindy", "Monse", "Gustavo",
        "Coni", "Miriam", "Abuelo Coni", "Mamá Mauri", "Papá Mauri"
    ]
    
    ws2["A1"] = "Integrante"
    ws2["A1"].font = header_font
    ws2["A1"].fill = header_fill
    ws2["A1"].alignment = Alignment(horizontal="center")
    
    for col_idx, item in enumerate(items, start=2):
        c = ws2.cell(row=1, column=col_idx, value=item)
        c.font = header_font
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center")
        
    si_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid") # green
    no_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid") # red
    
    for r_idx, persona in enumerate(integrantes, start=2):
        ws2.cell(row=r_idx, column=1, value=persona).font = bold_font
        for col_idx, item in enumerate(items, start=2):
            item_num = col_idx - 1 # 1 to 13
            if persona == "Monse":
                # Vegetariana: No carnes, no carbón, no pan choripán (1 a 9 = No). Sí Trigo Mote, Huesillos, Chancaca y Azúcar, Restante (10 a 13 = Sí)
                is_si = (item_num >= 10)
            elif persona == "Miriam":
                # No chorizo, lomo liso, lomo vetado, no pan choripán (1, 2, 3, 9 = No). Sí costillares, pollo, carbón, mote, huesillos, chancaca, restante.
                is_si = (item_num not in [1, 2, 3, 9])
            else:
                is_si = True
                
            val_text = "Sí" if is_si else "No"
            c = ws2.cell(row=r_idx, column=col_idx, value=val_text)
            c.alignment = Alignment(horizontal="center")
            c.font = regular_font
            c.fill = si_fill if is_si else no_fill
            c.border = thin_border
            
    # Fila de conteo
    count_row = len(integrantes) + 2
    ws2.cell(row=count_row, column=1, value="Total Consumidores").font = bold_font
    for col_idx in range(2, len(items) + 2):
        col_letter = get_column_letter(col_idx)
        c = ws2.cell(row=count_row, column=col_idx, value=f'=COUNTIF({col_letter}2:{col_letter}{count_row-1}, "Sí")')
        c.font = bold_font
        c.alignment = Alignment(horizontal="center")
        c.border = thin_border
        
    for col in ws2.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws2.column_dimensions[col_letter].width = max(max_len + 3, 14)

    # ----------------------------------------------------
    # HOJA 3: División de Cuentas & Liquidación
    # ----------------------------------------------------
    ws3 = wb.create_sheet(title="División de Cuentas")
    
    headers_ws3 = [
        "Integrante", "Consumo Propio", "Aporte Cuota Kena (1/4)", 
        "Total Cuota a Pagar", "Aportado en Compras", "Saldo Final a Transferir", 
        "Destinatario", "Estado Pago"
    ]
    
    for c_idx, h in enumerate(headers_ws3, start=1):
        cell = ws3.cell(row=1, column=c_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        
    sponsors_kena = ["Pamela", "Ale", "Mindy", "Joaquín"]
    # Row for Kena in ws3 is row 6 (Kena is index 4 in integrantes list -> row 6)
    kena_row = integrantes.index("Kena") + 2
    
    for r_idx, persona in enumerate(integrantes, start=2):
        ws3.cell(row=r_idx, column=1, value=persona).font = bold_font
        
        # Col B: Consumo Propio
        formula_terms = []
        for col_idx in range(2, len(items) + 2):
            col_letter = get_column_letter(col_idx)
            term = f"IF('Consumo por Integrante'!{col_letter}{r_idx}=\"Sí\", Compras!{col_letter}$11 / 'Consumo por Integrante'!{col_letter}${count_row}, 0)"
            formula_terms.append(term)
        
        exact_formula = "=" + " + ".join(formula_terms)
        c_exact = ws3.cell(row=r_idx, column=2, value=f"=ROUND({exact_formula[1:]}, 0)")
        c_exact.font = regular_font
        c_exact.number_format = money_format
        c_exact.border = thin_border
        
        # Col C: Aporte Cuota Kena (1/4)
        if persona in sponsors_kena:
            c_kena = ws3.cell(row=r_idx, column=3, value=f"=ROUND(B${kena_row}/4, 0)")
            c_kena.font = Font(name="Calibri", size=11, color="1E3A8A", bold=True)
        elif persona == "Kena":
            c_kena = ws3.cell(row=r_idx, column=3, value=f"=-B{r_idx}")
            c_kena.font = Font(name="Calibri", size=11, color="DC2626", bold=True)
        else:
            c_kena = ws3.cell(row=r_idx, column=3, value=0)
            c_kena.font = regular_font
        c_kena.number_format = money_format
        c_kena.border = thin_border
        
        # Col D: Total Cuota a Pagar
        if persona == "Kena":
            c_tot_pay = ws3.cell(row=r_idx, column=4, value=0)
        else:
            c_tot_pay = ws3.cell(row=r_idx, column=4, value=f"=B{r_idx}+C{r_idx}")
        c_tot_pay.font = bold_font
        c_tot_pay.number_format = money_format
        c_tot_pay.border = thin_border
        
        # Col E: Aportado en Compras
        aporte = 0
        if persona == "Ignacio":
            aporte = 183392
        elif persona == "Mindy":
            aporte = 11790
        elif persona == "Joaquín":
            aporte = 5570
            
        c_paid = ws3.cell(row=r_idx, column=5, value=aporte)
        c_paid.font = regular_font
        c_paid.number_format = money_format
        c_paid.border = thin_border
        
        # Col F: Saldo Final a Transferir
        if persona == "Ignacio":
            c_net = ws3.cell(row=r_idx, column=6, value=0)
            c_dest = ws3.cell(row=r_idx, column=7, value="Organizador (Recibe)")
            c_status = ws3.cell(row=r_idx, column=8, value="Organizador")
        elif persona == "Kena":
            c_net = ws3.cell(row=r_idx, column=6, value=0)
            c_dest = ws3.cell(row=r_idx, column=7, value="Cubierto por familiares")
            c_status = ws3.cell(row=r_idx, column=8, value="Cubierto")
        else:
            c_net = ws3.cell(row=r_idx, column=6, value=f"=IF(D{r_idx}-E{r_idx}>0, D{r_idx}-E{r_idx}, 0)")
            c_dest = ws3.cell(row=r_idx, column=7, value="Ignacio")
            c_status = ws3.cell(row=r_idx, column=8, value="Pendiente")
            
        c_net.font = Font(name="Calibri", size=11, bold=True, color="047857" if persona not in ["Ignacio", "Kena"] else "6B7280")
        c_net.number_format = money_format
        c_net.border = thin_border
        
        c_dest.font = regular_font
        c_dest.alignment = Alignment(horizontal="center")
        c_dest.border = thin_border
        
        c_status.font = regular_font
        c_status.alignment = Alignment(horizontal="center")
        c_status.border = thin_border
        
    tot_row = len(integrantes) + 2
    ws3.cell(row=tot_row, column=1, value="Total Recaudado").font = bold_font
    
    t_propio = ws3.cell(row=tot_row, column=2, value=f"=SUM(B2:B{tot_row-1})")
    t_propio.font = bold_font
    t_propio.number_format = money_format
    t_propio.fill = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")
    t_propio.border = thin_border
    
    t_kena = ws3.cell(row=tot_row, column=3, value=f"=SUM(C2:C{tot_row-1})")
    t_kena.font = bold_font
    t_kena.number_format = money_format
    t_kena.fill = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")
    t_kena.border = thin_border

    t_tot_pay = ws3.cell(row=tot_row, column=4, value=f"=SUM(D2:D{tot_row-1})")
    t_tot_pay.font = bold_font
    t_tot_pay.number_format = money_format
    t_tot_pay.fill = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")
    t_tot_pay.border = thin_border

    t_paid = ws3.cell(row=tot_row, column=5, value=f"=SUM(E2:E{tot_row-1})")
    t_paid.font = bold_font
    t_paid.number_format = money_format
    t_paid.fill = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")
    t_paid.border = thin_border

    t_net = ws3.cell(row=tot_row, column=6, value=f"=SUM(F2:F{tot_row-1})")
    t_net.font = bold_font
    t_net.number_format = money_format
    t_net.fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    t_net.border = thin_border

    ws3.cell(row=tot_row, column=7, value="")
    ws3.cell(row=tot_row, column=8, value="")

    for col in ws3.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws3.column_dimensions[col_letter].width = max(max_len + 4, 16)
        
    wb.save(file_path)
    print(f"Excel guardado con éxito en: {file_path}")

if __name__ == "__main__":
    update_excel("Cuentas Fiestas Patrias.xlsx")
    try:
        update_excel("C:/Users/Ignac/OneDrive/Escritorio/Cuentas Fiestas Patrias.xlsx")
    except Exception as e:
        print("No se pudo actualizar la copia en el escritorio:", e)
