import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

def load_existing_statuses(excel_path):
    """
    Reads the [VIEW] sheet if it exists, mapping task_id to its current status.
    This preserves human-entered statuses across script executions.
    """
    status_map = {}
    if not os.path.exists(excel_path):
        return status_map
    
    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        if "VIEW - Fila de Operacao" in wb.sheetnames:
            ws = wb["VIEW - Fila de Operacao"]
            headers = [cell.value for cell in ws[1]]
            
            if "task_id" in headers and "status" in headers:
                task_id_col = headers.index("task_id") + 1
                status_col = headers.index("status") + 1
                
                for r in range(2, ws.max_row + 1):
                    tid = ws.cell(row=r, column=task_id_col).value
                    status_val = ws.cell(row=r, column=status_col).value
                    if tid:
                        status_map[str(tid).strip()] = status_val
    except Exception as e:
        print(f"Error loading existing statuses: {e}")
        
    return status_map

def load_all_raw_ids(excel_path):
    """
    Reads the [RAW] sheet and returns a set of all task_ids that have already been ingested.
    """
    raw_ids = set()
    if not os.path.exists(excel_path):
        return raw_ids
    
    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        if "RAW - Ingestao" in wb.sheetnames:
            ws = wb["RAW - Ingestao"]
            # First column is task_id
            for r in range(2, ws.max_row + 1):
                tid = ws.cell(row=r, column=1).value
                if tid:
                    raw_ids.add(str(tid).strip())
    except Exception as e:
        print(f"Error reading raw task_ids: {e}")
        
    return raw_ids

def create_or_update_excel(excel_path, new_items):
    """
    Updates the Excel file with the new items.
    Appends all new items to [RAW].
    Updates [VIEW] to contain all active (non-completed) items, preserving their statuses.
    Regenerates [DASH] with dynamic KPIs and critical alerts.
    """
    # Load or create workbook
    if os.path.exists(excel_path):
        try:
            wb = openpyxl.load_workbook(excel_path)
        except Exception as e:
            print(f"Error loading workbook, creating new one: {e}")
            wb = openpyxl.Workbook()
    else:
        wb = openpyxl.Workbook()

    # Ensure sheets exist
    for sheet_name in ["RAW - Ingestao", "VIEW - Fila de Operacao", "DASH - Painel Executivo"]:
        if sheet_name not in wb.sheetnames:
            wb.create_sheet(sheet_name)
    
    # Remove default "Sheet" if present
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])
        
    ws_raw = wb["RAW - Ingestao"]
    ws_view = wb["VIEW - Fila de Operacao"]
    ws_dash = wb["DASH - Painel Executivo"]

    # Define headers
    raw_headers = [
        "task_id", "data_ingestao", "remetente_raiz", "categoria_ia", 
        "assunto_padronizado", "resumo_executivo", "valor_certame", 
        "prazo_fatal", "roteamento_responsavel", "link_anexo_seguro"
    ]
    
    view_headers = raw_headers + ["status"]

    # 1. Setup [RAW] Ingestion Layer
    if ws_raw.max_row == 1 and ws_raw.cell(row=1, column=1).value is None:
        for col_num, header in enumerate(raw_headers, 1):
            ws_raw.cell(row=1, column=col_num, value=header)
    
    # Check what is already in RAW
    existing_raw_ids = set()
    for r in range(2, ws_raw.max_row + 1):
        tid = ws_raw.cell(row=r, column=1).value
        if tid:
            existing_raw_ids.add(str(tid).strip())

    # Append only new items to RAW
    for item in new_items:
        tid = item["task_id"]
        if tid not in existing_raw_ids:
            ws_raw.append([
                item["task_id"],
                item["data_ingestao"],
                item["remetente_raiz"],
                item["categoria_ia"],
                item["assunto_padronizado"],
                item["resumo_executivo"],
                item["valor_certame"],
                item["prazo_fatal"],
                item["roteamento_responsavel"],
                item["link_anexo_seguro"]
            ])
            existing_raw_ids.add(tid)

    # 2. Setup [VIEW] Operational Layer
    # Read current statuses from VIEW before clearing it
    status_map = {}
    if ws_view.max_row > 1:
        headers = [cell.value for cell in ws_view[1]]
        if "task_id" in headers and "status" in headers:
            tid_idx = headers.index("task_id") + 1
            st_idx = headers.index("status") + 1
            for r in range(2, ws_view.max_row + 1):
                tid = ws_view.cell(row=r, column=tid_idx).value
                status_val = ws_view.cell(row=r, column=st_idx).value
                if tid:
                    status_map[str(tid).strip()] = status_val

    # Clear VIEW and re-populate
    ws_view.delete_rows(1, ws_view.max_row + 10)
    
    # Write headers to row 1
    for col_num, header in enumerate(view_headers, 1):
        ws_view.cell(row=1, column=col_num, value=header)

    # Load all items from RAW to reconstruct VIEW (excluding completed ones)
    view_row_idx = 2
    for r in range(2, ws_raw.max_row + 1):
        tid = str(ws_raw.cell(row=r, column=1).value or "").strip()
        if not tid or tid == "task_id":
            continue
            
        # Get status: if it was in our status map, keep it. Otherwise, default to "Pendente".
        status = status_map.get(tid, "Pendente")
        
        # If the status is "Concluído", do not add it to [VIEW]
        if status == "Concluído":
            continue
            
        # Exclude Informativos (Level 1) from the operational queue
        cat = ws_raw.cell(row=r, column=4).value
        if cat == "Informativo":
            continue

        # Add to VIEW row-by-row
        for col_num in range(1, len(raw_headers) + 1):
            val = ws_raw.cell(row=r, column=col_num).value
            ws_view.cell(row=view_row_idx, column=col_num, value=val)
        ws_view.cell(row=view_row_idx, column=len(view_headers), value=status)
        view_row_idx += 1

    # 3. Format sheets [RAW] and [VIEW]
    style_data_sheet(ws_raw, is_raw=True)
    style_data_sheet(ws_view, is_raw=False)

    # 4. Setup [DASH] Dashboard Layer
    build_dashboard(ws_dash, ws_view)

    # Save
    try:
        wb.save(excel_path)
        print(f"Excel file updated successfully at: {excel_path}")
    except PermissionError:
        print(f"\n[ERRO DE PERMISSÃO] Não foi possível salvar a planilha em: {excel_path}")
        print("Certifique-se de que o arquivo não está aberto no Microsoft Excel ou em outro programa e tente novamente.\n")
        raise

def style_data_sheet(ws, is_raw=False):
    """
    Applies colors, borders, gridlines, fonts, and dropdown list validations.
    """
    # Enable grid lines
    ws.views.sheetView[0].showGridLines = True
    
    font_header = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    font_data = Font(name="Segoe UI", size=10)
    
    # Navy fill for headers
    fill_header = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    
    # Light gray border
    thin_border = Border(
        left=Side(style='thin', color='E0E0E0'),
        right=Side(style='thin', color='E0E0E0'),
        top=Side(style='thin', color='E0E0E0'),
        bottom=Side(style='thin', color='E0E0E0')
    )
    
    # Alignments
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    
    # Format Headers
    ws.row_dimensions[1].height = 26
    for col_num in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center

    # Format Data Rows
    for r in range(2, ws.max_row + 1):
        ws.row_dimensions[r].height = 20
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = font_data
            cell.border = thin_border
            
            # Numeric columns
            if c == 7:  # valor_certame
                cell.number_format = 'R$ #,##0.00'
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif c in [1, 2, 4, 8, 9]:  # IDs, dates, categories, responsible
                cell.alignment = align_center
            else:
                cell.alignment = align_left

    # Auto-fit columns
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        
        # Don't auto-fit executive summary or link columns to huge widths
        is_large_col = col[0].value in ["resumo_executivo", "link_anexo_seguro", "assunto_padronizado"]
        
        for cell in col:
            val_str = str(cell.value or '')
            if len(val_str) > max_len:
                max_len = len(val_str)
                
        # Limit column widths to reasonable values
        if is_large_col:
            ws.column_dimensions[col_letter].width = min(40, max(max_len + 3, 15))
        else:
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    # Add Dropdown Data Validation for Status column in [VIEW]
    if not is_raw and ws.max_row > 1:
        # Status column is K (column 11)
        dv = DataValidation(type="list", formula1='"Pendente,Em Análise,Aguardando Assinatura,Concluído"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"K2:K1000")

def build_dashboard(ws, ws_view):
    """
    Builds the executive dashboard on DASH using data from VIEW.
    """
    # Clear dashboard
    ws.delete_rows(1, ws.max_row + 20)
    ws.views.sheetView[0].showGridLines = True

    # Title
    ws["B2"] = "PAINEL EXECUTIVO - GESTÃO E COMPLIANCE DE LICITAÇÕES"
    ws["B2"].font = Font(name="Segoe UI", size=16, bold=True, color="1B365D")
    
    # ----------------------------------------------------
    # KPI Cards Section
    # ----------------------------------------------------
    # Card 1: Volume Financeiro Pendente (B4:C5)
    ws.merge_cells("B4:C4")
    ws.merge_cells("B5:C5")
    ws["B4"] = "VOLUME EM LICITAÇÕES PENDENTES"
    ws["B5"] = "=SUM('VIEW - Fila de Operacao'!G2:G1000)"
    
    # Card 2: Licitações em Operação (E4:F4, E5:F5)
    ws.merge_cells("E4:F4")
    ws.merge_cells("E5:F5")
    ws["E4"] = "LICITAÇÕES ATIVAS NA OPERAÇÃO"
    ws["E5"] = "=COUNTA('VIEW - Fila de Operacao'!A2:A1000)"
    
    # Card 3: Risco Crítico - Nível 4 (H4:I4, H5:I5)
    ws.merge_cells("H4:I4")
    ws.merge_cells("H5:I5")
    ws["H4"] = "ALERTAS DE RISCO CRÍTICO (N4)"
    ws["H5"] = "=COUNTIF('VIEW - Fila de Operacao'!D2:D1000, \"Risco Crítico\")"

    # Style KPI Cards
    fill_blue = PatternFill(start_color="E6EEF8", end_color="E6EEF8", fill_type="solid")
    fill_gray = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")
    fill_red = PatternFill(start_color="FDF2F2", end_color="FDF2F2", fill_type="solid")
    
    font_lbl = Font(name="Segoe UI", size=9, bold=True, color="555555")
    font_num = Font(name="Segoe UI", size=18, bold=True, color="1B365D")
    font_num_red = Font(name="Segoe UI", size=18, bold=True, color="D9534F")
    
    align_center = Alignment(horizontal="center", vertical="center")
    
    card_border = Border(
        left=Side(style='thin', color='B0C4DE'),
        right=Side(style='thin', color='B0C4DE'),
        top=Side(style='thin', color='B0C4DE'),
        bottom=Side(style='thin', color='B0C4DE')
    )
    
    # Apply card 1 styles
    for row in range(4, 6):
        for col in range(2, 4):
            cell = ws.cell(row=row, column=col)
            cell.fill = fill_blue
            cell.border = card_border
    ws["B4"].font = font_lbl
    ws["B4"].alignment = align_center
    ws["B5"].font = font_num
    ws["B5"].alignment = align_center
    ws["B5"].number_format = 'R$ #,##0.00'

    # Apply card 2 styles
    for row in range(4, 6):
        for col in range(5, 7):
            cell = ws.cell(row=row, column=col)
            cell.fill = fill_gray
            cell.border = card_border
    ws["E4"].font = font_lbl
    ws["E4"].alignment = align_center
    ws["E5"].font = font_num
    ws["E5"].alignment = align_center
    ws["E5"].number_format = '#,##0'

    # Apply card 3 styles
    for row in range(4, 6):
        for col in range(8, 10):
            cell = ws.cell(row=row, column=col)
            cell.fill = fill_red
            cell.border = Border(
                left=Side(style='thin', color='FFC1C1'),
                right=Side(style='thin', color='FFC1C1'),
                top=Side(style='thin', color='FFC1C1'),
                bottom=Side(style='thin', color='FFC1C1')
            )
    ws["H4"].font = Font(name="Segoe UI", size=9, bold=True, color="A94442")
    ws["H4"].alignment = align_center
    ws["H5"].font = font_num_red
    ws["H5"].alignment = align_center
    ws["H5"].number_format = '#,##0'

    # ----------------------------------------------------
    # Level 4 Alert Table Section
    # ----------------------------------------------------
    ws["B7"] = "DETALHES DE ALERTAS DE RISCO CRÍTICO (NÍVEL 4)"
    ws["B7"].font = Font(name="Segoe UI", size=12, bold=True, color="A94442")
    
    headers = ["Órgão/Cliente", "Assunto", "Resumo Executivo", "Valor (R$)", "Prazo Fatal", "Link do Anexo"]
    for col_idx, h in enumerate(headers, start=2):
        cell = ws.cell(row=8, column=col_idx)
        cell.value = h
        cell.font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="D9534F", end_color="D9534F", fill_type="solid")
        cell.alignment = align_center
        cell.border = Border(bottom=Side(style='medium', color='A94442'))
        
    # Populate Level 4 items statically from ws_view
    row_idx = 9
    thin_red_border = Border(
        left=Side(style='thin', color='FFC1C1'),
        right=Side(style='thin', color='FFC1C1'),
        top=Side(style='thin', color='FFC1C1'),
        bottom=Side(style='thin', color='FFC1C1')
    )
    fill_row_red = PatternFill(start_color="FFF8F8", end_color="FFF8F8", fill_type="solid")
    
    for r in range(2, ws_view.max_row + 1):
        cat = ws_view.cell(row=r, column=4).value
        if cat == "Risco Crítico":
            # Extract cells: remetente_raiz, assunto_padronizado, resumo_executivo, valor_certame, prazo_fatal, link_anexo_seguro
            remetente = ws_view.cell(row=r, column=3).value
            assunto = ws_view.cell(row=r, column=5).value
            resumo = ws_view.cell(row=r, column=6).value
            valor = ws_view.cell(row=r, column=7).value
            prazo = ws_view.cell(row=r, column=8).value
            link = ws_view.cell(row=r, column=10).value
            
            row_vals = [remetente, assunto, resumo, valor, prazo, link]
            for col_offset, val in enumerate(row_vals):
                cell = ws.cell(row=row_idx, column=2 + col_offset)
                cell.value = val
                cell.font = Font(name="Segoe UI", size=9)
                cell.fill = fill_row_red
                cell.border = thin_red_border
                
                # Alignments and formats
                if col_offset == 3:  # Valor
                    cell.number_format = 'R$ #,##0.00'
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                elif col_offset in [4]:  # Prazo
                    cell.alignment = align_center
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                    
            ws.row_dimensions[row_idx].height = 22
            row_idx += 1
            
    # Set column widths for Dashboard
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 20 # Remetente
    ws.column_dimensions["C"].width = 30 # Assunto
    ws.column_dimensions["D"].width = 40 # Resumo
    ws.column_dimensions["E"].width = 18 # Valor
    ws.column_dimensions["F"].width = 15 # Prazo
    ws.column_dimensions["G"].width = 25 # Link Anexo
    ws.column_dimensions["H"].width = 18
    ws.column_dimensions["I"].width = 18
