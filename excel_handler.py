"""
excel_handler.py — Motor de Planilha Executiva v3.0
4 Abas: Painel · Operação · Histórico · Arquivo
Paleta: Azul Marinho (#1B365D) + Verde Esmeralda (#2E7D32)
3 Dropdowns: Tema, Área Destino, Status
Dashboard: KPIs + Semáforo SLA + Timeline + Alertas Diretoria
"""

import os
from datetime import datetime, timezone
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from routing_tables import (
    THEMES, TARGET_AREAS, WORKFLOW_STATES, THEME_COLORS,
    format_br_datetime, format_br_date, calculate_sla_days,
    resolve_theme, resolve_target,
)

# ============================================================
#  CONSTANTES VISUAIS
# ============================================================

# Paleta de Cores
NAVY = "1B365D"
EMERALD = "2E7D32"
WHITE = "FFFFFF"
LIGHT_GRAY = "F9F9F9"
MID_GRAY = "E0E0E0"
DARK_TEXT = "333333"
RED_ALERT = "C62828"
AMBER = "F9A825"
GREEN_OK = "2E7D32"

# Fontes
FONT_HEADER = Font(name="Segoe UI", size=11, bold=True, color=WHITE)
FONT_DATA = Font(name="Segoe UI", size=10, color=DARK_TEXT)
FONT_DATA_BOLD = Font(name="Segoe UI", size=10, bold=True, color=DARK_TEXT)
FONT_STRIKETHROUGH = Font(name="Segoe UI", size=10, color="999999", strikethrough=True)
FONT_TITLE = Font(name="Segoe UI", size=16, bold=True, color=NAVY)
FONT_KPI_LABEL = Font(name="Segoe UI", size=9, bold=True, color="555555")
FONT_KPI_NUM = Font(name="Segoe UI", size=20, bold=True, color=NAVY)
FONT_KPI_RED = Font(name="Segoe UI", size=20, bold=True, color=RED_ALERT)
FONT_SECTION = Font(name="Segoe UI", size=12, bold=True, color=NAVY)
FONT_LINK = Font(name="Segoe UI", size=10, color="1565C0", underline="single")
FONT_CHECK_YES = Font(name="Segoe UI", size=11, bold=True, color=EMERALD)
FONT_CHECK_NO = Font(name="Segoe UI", size=10, color="BDBDBD")

# Fills
FILL_HEADER = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
FILL_ZEBRA = PatternFill(start_color=LIGHT_GRAY, end_color=LIGHT_GRAY, fill_type="solid")
FILL_WHITE = PatternFill(start_color=WHITE, end_color=WHITE, fill_type="solid")
FILL_FINALIZED = PatternFill(start_color=MID_GRAY, end_color=MID_GRAY, fill_type="solid")
FILL_KPI_BLUE = PatternFill(start_color="E8EAF6", end_color="E8EAF6", fill_type="solid")
FILL_KPI_GREEN = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")
FILL_KPI_RED = PatternFill(start_color="FFEBEE", end_color="FFEBEE", fill_type="solid")
FILL_KPI_AMBER = PatternFill(start_color="FFF8E1", end_color="FFF8E1", fill_type="solid")

# Borders
THIN_BORDER = Border(
    left=Side(style='thin', color='E0E0E0'),
    right=Side(style='thin', color='E0E0E0'),
    top=Side(style='thin', color='E0E0E0'),
    bottom=Side(style='thin', color='E0E0E0'),
)
ACCENT_BORDER = Border(
    left=Side(style='thin', color=EMERALD),
    right=Side(style='thin', color=EMERALD),
    top=Side(style='thin', color=EMERALD),
    bottom=Side(style='thin', color=EMERALD),
)

# Alignments
ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=False)
ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")

# ============================================================
#  LABELS DOS CABEÇALHOS (português profissional)
# ============================================================

# Colunas da aba Histórico (RAW)
RAW_LABELS = [
    "ID Demanda", "Recebido Em", "Órgão / Cliente", "Tema",
    "Assunto", "Resumo Executivo", "Valor Estimado (R$)",
    "Prazo Limite", "Área Destino", "Diretoria?",
    "Responder", "Anexos",
]
RAW_KEYS = [
    "task_id", "data_ingestao_fmt", "remetente_raiz", "theme_label",
    "assunto_padronizado", "executive_summary", "valor_certame",
    "prazo_fatal_fmt", "target_label", "board_flag",
    "action_link", "link_anexo_seguro",
]

# Colunas da aba Operação (VIEW) = RAW + Status
VIEW_LABELS = RAW_LABELS[:9] + ["Status"] + RAW_LABELS[9:]
VIEW_KEYS = RAW_KEYS[:9] + ["status_workflow"] + RAW_KEYS[9:]

# Dropdown values
DROPDOWN_THEMES = ",".join([THEMES[i] for i in range(1, 7)])
DROPDOWN_TARGETS = ",".join([TARGET_AREAS[i] for i in range(1, 9)])
DROPDOWN_STATUS = ",".join(WORKFLOW_STATES)


# ============================================================
#  FUNÇÕES AUXILIARES
# ============================================================

def _prepare_item_for_excel(item: dict) -> dict:
    """Prepara um item do pipeline para escrita no Excel, formatando datas e flags."""
    row = dict(item)
    row["data_ingestao_fmt"] = format_br_datetime(item.get("data_ingestao", ""))
    row["prazo_fatal_fmt"] = format_br_date(item.get("prazo_fatal", ""))
    row["board_flag"] = "✅" if item.get("requires_board_approval") else "❌"
    return row


def _get_row_values(item: dict, keys: list) -> list:
    """Extrai valores do item na ordem das chaves."""
    return [item.get(k, "") for k in keys]


# ============================================================
#  LOAD EXISTING DATA
# ============================================================

def load_existing_statuses(excel_path: str) -> dict:
    """
    Lê a aba Operação e retorna {task_id: status_workflow}.
    Preserva edições humanas entre execuções.
    """
    status_map = {}
    if not os.path.exists(excel_path):
        return status_map

    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        sheet_name = "Operação"
        if sheet_name not in wb.sheetnames:
            return status_map

        ws = wb[sheet_name]
        headers = [cell.value for cell in ws[1]]

        tid_label = "ID Demanda"
        status_label = "Status"

        if tid_label in headers and status_label in headers:
            tid_col = headers.index(tid_label) + 1
            st_col = headers.index(status_label) + 1
            for r in range(2, ws.max_row + 1):
                tid = ws.cell(row=r, column=tid_col).value
                status_val = ws.cell(row=r, column=st_col).value
                if tid:
                    status_map[str(tid).strip()] = status_val
    except Exception as e:
        print(f"  ⚠ Erro ao ler status existentes: {e}")

    return status_map


def load_all_raw_ids(excel_path: str) -> set:
    """Lê a aba Histórico e retorna set de task_ids já ingeridos."""
    raw_ids = set()
    if not os.path.exists(excel_path):
        return raw_ids

    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        sheet_name = "Histórico"
        if sheet_name not in wb.sheetnames:
            return raw_ids

        ws = wb[sheet_name]
        for r in range(2, ws.max_row + 1):
            tid = ws.cell(row=r, column=1).value
            if tid:
                raw_ids.add(str(tid).strip())
    except Exception as e:
        print(f"  ⚠ Erro ao ler IDs do histórico: {e}")

    return raw_ids


def _load_existing_themes_targets(excel_path: str) -> dict:
    """Lê reclassificações manuais de Tema e Área Destino da aba Operação."""
    overrides = {}
    if not os.path.exists(excel_path):
        return overrides

    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        if "Operação" not in wb.sheetnames:
            return overrides
        ws = wb["Operação"]
        headers = [cell.value for cell in ws[1]]

        tid_label = "ID Demanda"
        theme_label = "Tema"
        target_label = "Área Destino"

        if tid_label in headers and theme_label in headers and target_label in headers:
            tid_col = headers.index(tid_label) + 1
            theme_col = headers.index(theme_label) + 1
            target_col = headers.index(target_label) + 1
            for r in range(2, ws.max_row + 1):
                tid = ws.cell(row=r, column=tid_col).value
                theme_val = ws.cell(row=r, column=theme_col).value
                target_val = ws.cell(row=r, column=target_col).value
                if tid:
                    overrides[str(tid).strip()] = {
                        "theme_label": theme_val,
                        "target_label": target_val,
                    }
    except Exception:
        pass
    return overrides


# ============================================================
#  MAIN: CREATE OR UPDATE EXCEL
# ============================================================

def create_or_update_excel(excel_path: str, new_items: list):
    """
    Atualiza a planilha de monitoramento com novos itens.
    - Histórico: append-only (audit trail)
    - Operação: fila ativa com status editáveis
    - Arquivo: itens finalizados com data de conclusão
    - Painel: dashboard executivo com KPIs
    """
    # Garantir diretório
    os.makedirs(os.path.dirname(excel_path) if os.path.dirname(excel_path) else ".", exist_ok=True)

    # Carregar ou criar workbook
    if os.path.exists(excel_path):
        try:
            wb = openpyxl.load_workbook(excel_path)
        except Exception as e:
            print(f"  ⚠ Erro ao abrir planilha, criando nova: {e}")
            wb = openpyxl.Workbook()
    else:
        wb = openpyxl.Workbook()

    # Garantir as 4 abas na ordem certa
    desired_sheets = ["Painel", "Operação", "Histórico", "Arquivo"]
    for name in desired_sheets:
        if name not in wb.sheetnames:
            wb.create_sheet(name)

    # Remover aba padrão "Sheet"
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])

    # Remover abas antigas se existirem (migração de v1)
    for old_name in ["RAW - Ingestao", "VIEW - Fila de Operacao", "DASH - Painel Executivo"]:
        if old_name in wb.sheetnames:
            wb.remove(wb[old_name])

    # Reordenar abas
    sheet_order = []
    for name in desired_sheets:
        if name in wb.sheetnames:
            sheet_order.append(wb.sheetnames.index(name))
    wb.move_sheet("Painel", offset=-wb.sheetnames.index("Painel"))

    ws_hist = wb["Histórico"]
    ws_op = wb["Operação"]
    ws_arquivo = wb["Arquivo"]
    ws_dash = wb["Painel"]

    # ─────────────────────────────────────────────
    # 1. HISTÓRICO (RAW — append only)
    # ─────────────────────────────────────────────
    if ws_hist.max_row == 1 and ws_hist.cell(row=1, column=1).value is None:
        for col, label in enumerate(RAW_LABELS, 1):
            ws_hist.cell(row=1, column=col, value=label)

    # Coletar IDs existentes no Histórico
    existing_raw_ids = set()
    for r in range(2, ws_hist.max_row + 1):
        tid = ws_hist.cell(row=r, column=1).value
        if tid:
            existing_raw_ids.add(str(tid).strip())

    # Append novos itens ao Histórico
    for item in new_items:
        tid = item.get("task_id", "")
        if tid and tid not in existing_raw_ids:
            prepared = _prepare_item_for_excel(item)
            row_vals = _get_row_values(prepared, RAW_KEYS)
            ws_hist.append(row_vals)
            existing_raw_ids.add(tid)

    # ─────────────────────────────────────────────
    # 2. OPERAÇÃO (VIEW — fila ativa)
    # ─────────────────────────────────────────────
    # Ler status e overrides existentes ANTES de limpar
    status_map = {}
    overrides_map = {}
    if ws_op.max_row > 1 and ws_op.cell(row=1, column=1).value is not None:
        headers = [cell.value for cell in ws_op[1]]
        tid_col_idx = headers.index("ID Demanda") + 1 if "ID Demanda" in headers else 1
        status_col_idx = headers.index("Status") + 1 if "Status" in headers else None
        theme_col_idx = headers.index("Tema") + 1 if "Tema" in headers else None
        target_col_idx = headers.index("Área Destino") + 1 if "Área Destino" in headers else None

        for r in range(2, ws_op.max_row + 1):
            tid = ws_op.cell(row=r, column=tid_col_idx).value
            if tid:
                tid_str = str(tid).strip()
                if status_col_idx:
                    status_map[tid_str] = ws_op.cell(row=r, column=status_col_idx).value
                if theme_col_idx and target_col_idx:
                    overrides_map[tid_str] = {
                        "theme_label": ws_op.cell(row=r, column=theme_col_idx).value,
                        "target_label": ws_op.cell(row=r, column=target_col_idx).value,
                    }

    # Limpar aba Operação
    ws_op.delete_rows(1, ws_op.max_row + 10)

    # Cabeçalhos
    for col, label in enumerate(VIEW_LABELS, 1):
        ws_op.cell(row=1, column=col, value=label)

    # ─────────────────────────────────────────────
    # 3. ARQUIVO (itens finalizados)
    # ─────────────────────────────────────────────
    arquivo_labels = VIEW_LABELS + ["Data de Conclusão"]
    if ws_arquivo.max_row == 1 and ws_arquivo.cell(row=1, column=1).value is None:
        for col, label in enumerate(arquivo_labels, 1):
            ws_arquivo.cell(row=1, column=col, value=label)

    # Coletar IDs já arquivados
    archived_ids = set()
    for r in range(2, ws_arquivo.max_row + 1):
        tid = ws_arquivo.cell(row=r, column=1).value
        if tid:
            archived_ids.add(str(tid).strip())

    # ─────────────────────────────────────────────
    # 4. RECONSTRUIR OPERAÇÃO a partir do Histórico
    # ─────────────────────────────────────────────
    view_row = 2
    finalized_row_indices = []  # Para aplicar estilo cinza+tachado depois

    for r in range(2, ws_hist.max_row + 1):
        tid = str(ws_hist.cell(row=r, column=1).value or "").strip()
        if not tid:
            continue

        # Ler dados do Histórico
        row_data = {}
        for col_idx, key in enumerate(RAW_KEYS, 1):
            row_data[key] = ws_hist.cell(row=r, column=col_idx).value

        # Aplicar overrides manuais (Tema, Área Destino reclassificados pelo humano)
        if tid in overrides_map:
            ov = overrides_map[tid]
            if ov.get("theme_label"):
                row_data["theme_label"] = ov["theme_label"]
            if ov.get("target_label"):
                row_data["target_label"] = ov["target_label"]

        # Determinar status (preservar edição humana)
        status = status_map.get(tid, "Novo")
        # Novos itens que acabaram de ser adicionados nesta execução
        for item in new_items:
            if item.get("task_id") == tid and tid not in status_map:
                status = item.get("status_workflow", "Novo")
                break

        is_finalized = (status == "Finalizado")

        # Skip Informativos/Descartados
        theme = row_data.get("theme_label", "")
        if theme == "Informativo / Não Relevante":
            continue

        # Mover finalizados para Arquivo (se ainda não estão lá)
        if is_finalized and tid not in archived_ids:
            # Escrever na aba Arquivo
            arquivo_row_vals = _get_row_values(row_data, RAW_KEYS[:9]) + \
                              [status] + \
                              _get_row_values(row_data, RAW_KEYS[9:]) + \
                              [format_br_datetime(datetime.now(timezone.utc).isoformat())]
            ws_arquivo.append(arquivo_row_vals)
            archived_ids.add(tid)

        # Escrever na Operação
        view_keys_ordered = RAW_KEYS[:9] + ["status_workflow"] + RAW_KEYS[9:]
        row_data["status_workflow"] = status

        for col_idx, key in enumerate(view_keys_ordered, 1):
            val = row_data.get(key, "")
            ws_op.cell(row=view_row, column=col_idx, value=val)

        if is_finalized:
            finalized_row_indices.append(view_row)

        view_row += 1

    # ─────────────────────────────────────────────
    # 5. ESTILIZAR ABAS
    # ─────────────────────────────────────────────
    _style_data_sheet(ws_hist, "historico")
    _style_data_sheet(ws_op, "operacao", finalized_rows=finalized_row_indices)
    _style_data_sheet(ws_arquivo, "arquivo")

    # ─────────────────────────────────────────────
    # 6. DASHBOARD (Painel)
    # ─────────────────────────────────────────────
    _build_dashboard(ws_dash, ws_op)

    # ─────────────────────────────────────────────
    # 7. REORDENAR ABAS: Painel primeiro
    # ─────────────────────────────────────────────
    try:
        desired_order = ["Painel", "Operação", "Histórico", "Arquivo"]
        for i, name in enumerate(desired_order):
            if name in wb.sheetnames:
                current_idx = wb.sheetnames.index(name)
                wb.move_sheet(name, offset=i - current_idx)
    except Exception:
        pass

    # ─────────────────────────────────────────────
    # 8. SALVAR
    # ─────────────────────────────────────────────
    try:
        wb.save(excel_path)
        print(f"  ✅ Planilha atualizada: {excel_path}")
    except PermissionError:
        print(f"\n  ❌ ERRO: Não foi possível salvar em: {excel_path}")
        print("  Feche o arquivo no Excel e tente novamente.\n")
        raise


# ============================================================
#  ESTILIZAÇÃO DAS ABAS DE DADOS
# ============================================================

def _style_data_sheet(ws, sheet_type: str, finalized_rows: list = None):
    """Aplica formatação profissional completa a uma aba de dados."""
    if finalized_rows is None:
        finalized_rows = []

    max_col = ws.max_column
    max_row = ws.max_row

    if max_row < 1 or max_col < 1:
        return

    # --- Cabeçalhos ---
    ws.row_dimensions[1].height = 28
    for col in range(1, max_col + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = THIN_BORDER

    # --- Linhas de Dados ---
    # Mapear coluna de Tema para colorir por tema
    headers = [ws.cell(row=1, column=c).value for c in range(1, max_col + 1)]
    theme_col = (headers.index("Tema") + 1) if "Tema" in headers else None
    valor_col = (headers.index("Valor Estimado (R$)") + 1) if "Valor Estimado (R$)" in headers else None
    board_col = (headers.index("Diretoria?") + 1) if "Diretoria?" in headers else None
    action_col = (headers.index("Responder") + 1) if "Responder" in headers else None

    for r in range(2, max_row + 1):
        is_finalized = r in finalized_rows
        is_zebra = (r % 2 == 0)
        ws.row_dimensions[r].height = 22

        # Determinar fill por tema
        theme_fill = None
        if theme_col and not is_finalized:
            theme_val = ws.cell(row=r, column=theme_col).value
            for tid, label in THEMES.items():
                if label == theme_val and tid in THEME_COLORS:
                    theme_fill = PatternFill(
                        start_color=THEME_COLORS[tid],
                        end_color=THEME_COLORS[tid],
                        fill_type="solid"
                    )
                    break

        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)

            # Fonte
            if is_finalized:
                cell.font = FONT_STRIKETHROUGH
            elif c == board_col:
                val = cell.value
                cell.font = FONT_CHECK_YES if val == "✅" else FONT_CHECK_NO
            elif c == action_col and cell.value:
                cell.font = FONT_LINK
            else:
                cell.font = FONT_DATA

            # Fill
            if is_finalized:
                cell.fill = FILL_FINALIZED
            elif theme_fill and c == theme_col:
                cell.fill = theme_fill
            elif is_zebra:
                cell.fill = FILL_ZEBRA
            else:
                cell.fill = FILL_WHITE

            # Border
            cell.border = THIN_BORDER

            # Alignment
            if c == valor_col:
                cell.alignment = ALIGN_RIGHT
                cell.number_format = 'R$ #,##0.00'
            elif c in [1, 2, 8] or c == board_col:
                cell.alignment = ALIGN_CENTER
            elif c == action_col:
                cell.alignment = ALIGN_CENTER
            else:
                cell.alignment = ALIGN_LEFT

    # --- Larguras de Coluna ---
    col_widths = {
        "ID Demanda": 18,
        "Recebido Em": 22,
        "Órgão / Cliente": 22,
        "Tema": 26,
        "Assunto": 40,
        "Resumo Executivo": 45,
        "Valor Estimado (R$)": 20,
        "Prazo Limite": 16,
        "Área Destino": 24,
        "Status": 22,
        "Diretoria?": 12,
        "Responder": 35,
        "Anexos": 30,
        "Data de Conclusão": 22,
    }
    for c in range(1, max_col + 1):
        header = ws.cell(row=1, column=c).value
        width = col_widths.get(header, 15)
        ws.column_dimensions[get_column_letter(c)].width = width

    # --- Dropdowns (Data Validation) — somente na aba Operação ---
    if sheet_type == "operacao" and max_row > 1:
        range_end = max(max_row, 500)

        # Dropdown de Tema (coluna D = 4)
        if "Tema" in headers:
            col_letter = get_column_letter(headers.index("Tema") + 1)
            dv_theme = DataValidation(
                type="list",
                formula1=f'"{DROPDOWN_THEMES}"',
                allow_blank=True
            )
            dv_theme.error = "Selecione um tema válido da lista."
            dv_theme.errorTitle = "Tema Inválido"
            ws.add_data_validation(dv_theme)
            dv_theme.add(f"{col_letter}2:{col_letter}{range_end}")

        # Dropdown de Área Destino
        if "Área Destino" in headers:
            col_letter = get_column_letter(headers.index("Área Destino") + 1)
            dv_target = DataValidation(
                type="list",
                formula1=f'"{DROPDOWN_TARGETS}"',
                allow_blank=True
            )
            dv_target.error = "Selecione uma área válida da lista."
            dv_target.errorTitle = "Área Inválida"
            ws.add_data_validation(dv_target)
            dv_target.add(f"{col_letter}2:{col_letter}{range_end}")

        # Dropdown de Status
        if "Status" in headers:
            col_letter = get_column_letter(headers.index("Status") + 1)
            dv_status = DataValidation(
                type="list",
                formula1=f'"{DROPDOWN_STATUS}"',
                allow_blank=True
            )
            dv_status.error = "Selecione um status válido da lista."
            dv_status.errorTitle = "Status Inválido"
            ws.add_data_validation(dv_status)
            dv_status.add(f"{col_letter}2:{col_letter}{range_end}")

    # --- Freeze top row ---
    ws.freeze_panes = "A2"

    # --- Auto-filter ---
    if max_row > 1 and max_col > 1:
        ws.auto_filter.ref = f"A1:{get_column_letter(max_col)}{max_row}"


# ============================================================
#  DASHBOARD EXECUTIVO (Aba "Painel")
# ============================================================

def _build_dashboard(ws_dash, ws_op):
    """Constrói o dashboard executivo com KPIs, Semáforo, Timeline e Alertas."""

    # Limpar dashboard
    ws_dash.delete_rows(1, ws_dash.max_row + 30)

    # Desabilitar grid para look mais limpo
    ws_dash.sheet_view.showGridLines = False

    # ─── TÍTULO ───
    ws_dash.merge_cells("B2:J2")
    ws_dash["B2"] = "PAINEL EXECUTIVO — GESTÃO DE LICITAÇÕES E COMPLIANCE"
    ws_dash["B2"].font = FONT_TITLE

    ws_dash.merge_cells("B3:J3")
    ws_dash["B3"] = f"Última atualização: {format_br_datetime(datetime.now(timezone.utc).isoformat())}"
    ws_dash["B3"].font = Font(name="Segoe UI", size=9, italic=True, color="888888")

    # ─── KPI CARDS (Linha 5-6) ───
    _build_kpi_card(ws_dash, "B", 5, "VOLUME FINANCEIRO PENDENTE",
                    f"=SUM('Operação'!G2:G1000)", FILL_KPI_BLUE, FONT_KPI_NUM,
                    num_format='R$ #,##0.00')

    _build_kpi_card(ws_dash, "E", 5, "DEMANDAS ATIVAS",
                    f"=COUNTA('Operação'!A2:A1000)", FILL_KPI_GREEN, FONT_KPI_NUM,
                    num_format='#,##0')

    _build_kpi_card(ws_dash, "H", 5, "ALERTAS DIRETORIA (✅)",
                    f"=COUNTIF('Operação'!K2:K1000,\"✅\")", FILL_KPI_RED, FONT_KPI_RED,
                    num_format='#,##0')

    # ─── SEMÁFORO SLA (Linha 8-9) ───
    ws_dash.merge_cells("B8:J8")
    ws_dash["B8"] = "SEMÁFORO DE SLA — PRAZOS ATIVOS"
    ws_dash["B8"].font = FONT_SECTION
    ws_dash["B8"].border = Border(bottom=Side(style='medium', color=EMERALD))

    sla_headers = ["🔴 Vencido / <48h", "🟡 3 a 7 dias", "🟢 >7 dias", "⚪ Sem prazo"]
    sla_colors = [FILL_KPI_RED, FILL_KPI_AMBER, FILL_KPI_GREEN, FILL_WHITE]

    # Calcular semáforo a partir dos dados da aba Operação
    sla_counts = {"red": 0, "yellow": 0, "green": 0, "none": 0}
    headers_op = [ws_op.cell(row=1, column=c).value for c in range(1, ws_op.max_column + 1)]
    prazo_col_idx = (headers_op.index("Prazo Limite") + 1) if "Prazo Limite" in headers_op else None

    if prazo_col_idx:
        for r in range(2, ws_op.max_row + 1):
            prazo_val = ws_op.cell(row=r, column=prazo_col_idx).value
            if not prazo_val:
                sla_counts["none"] += 1
            else:
                # Tentar calcular dias restantes
                days = _calculate_days_from_formatted(prazo_val)
                if days <= 2:
                    sla_counts["red"] += 1
                elif days <= 7:
                    sla_counts["yellow"] += 1
                else:
                    sla_counts["green"] += 1

    sla_values = [sla_counts["red"], sla_counts["yellow"], sla_counts["green"], sla_counts["none"]]

    for i, (label, fill, val) in enumerate(zip(sla_headers, sla_colors, sla_values)):
        col_start = 2 + (i * 2)
        col_letter_1 = get_column_letter(col_start)
        col_letter_2 = get_column_letter(col_start + 1)

        # Label
        ws_dash.merge_cells(f"{col_letter_1}9:{col_letter_2}9")
        cell_label = ws_dash.cell(row=9, column=col_start)
        cell_label.value = label
        cell_label.font = FONT_KPI_LABEL
        cell_label.alignment = ALIGN_CENTER
        cell_label.fill = fill

        # Value
        ws_dash.merge_cells(f"{col_letter_1}10:{col_letter_2}10")
        cell_val = ws_dash.cell(row=10, column=col_start)
        cell_val.value = val
        cell_val.font = Font(name="Segoe UI", size=16, bold=True, color=NAVY)
        cell_val.alignment = ALIGN_CENTER
        cell_val.fill = fill

        # Border
        for row in [9, 10]:
            for col in [col_start, col_start + 1]:
                ws_dash.cell(row=row, column=col).border = ACCENT_BORDER

    # ─── TIMELINE DE PRAZOS (Linha 12+) ───
    ws_dash.merge_cells("B12:J12")
    ws_dash["B12"] = "PRÓXIMOS PRAZOS FATAIS — TOP 10 URGÊNCIAS"
    ws_dash["B12"].font = FONT_SECTION
    ws_dash["B12"].border = Border(bottom=Side(style='medium', color=EMERALD))

    timeline_headers = ["Prazo", "Dias Rest.", "Órgão / Cliente", "Assunto", "Valor (R$)", "Área Destino"]
    for i, h in enumerate(timeline_headers):
        cell = ws_dash.cell(row=13, column=2 + i)
        cell.value = h
        cell.font = Font(name="Segoe UI", size=10, bold=True, color=WHITE)
        cell.fill = PatternFill(start_color=EMERALD, end_color=EMERALD, fill_type="solid")
        cell.alignment = ALIGN_CENTER
        cell.border = ACCENT_BORDER

    # Coletar itens com prazo da aba Operação
    items_with_deadline = []
    if prazo_col_idx:
        org_col = (headers_op.index("Órgão / Cliente") + 1) if "Órgão / Cliente" in headers_op else None
        assunto_col = (headers_op.index("Assunto") + 1) if "Assunto" in headers_op else None
        valor_col = (headers_op.index("Valor Estimado (R$)") + 1) if "Valor Estimado (R$)" in headers_op else None
        target_col = (headers_op.index("Área Destino") + 1) if "Área Destino" in headers_op else None

        for r in range(2, ws_op.max_row + 1):
            prazo_val = ws_op.cell(row=r, column=prazo_col_idx).value
            if prazo_val:
                days = _calculate_days_from_formatted(prazo_val)
                items_with_deadline.append({
                    "prazo": prazo_val,
                    "days": days,
                    "org": ws_op.cell(row=r, column=org_col).value if org_col else "",
                    "assunto": ws_op.cell(row=r, column=assunto_col).value if assunto_col else "",
                    "valor": ws_op.cell(row=r, column=valor_col).value if valor_col else "",
                    "target": ws_op.cell(row=r, column=target_col).value if target_col else "",
                })

    # Ordenar por urgência (menos dias primeiro)
    items_with_deadline.sort(key=lambda x: x["days"])

    for idx, item in enumerate(items_with_deadline[:10]):
        row = 14 + idx
        ws_dash.cell(row=row, column=2, value=item["prazo"]).font = FONT_DATA
        ws_dash.cell(row=row, column=2).alignment = ALIGN_CENTER

        days_cell = ws_dash.cell(row=row, column=3, value=item["days"])
        days_cell.alignment = ALIGN_CENTER
        if item["days"] <= 2:
            days_cell.font = Font(name="Segoe UI", size=10, bold=True, color=RED_ALERT)
        elif item["days"] <= 7:
            days_cell.font = Font(name="Segoe UI", size=10, bold=True, color=AMBER)
        else:
            days_cell.font = Font(name="Segoe UI", size=10, color=EMERALD)

        ws_dash.cell(row=row, column=4, value=item["org"]).font = FONT_DATA
        ws_dash.cell(row=row, column=5, value=item["assunto"]).font = FONT_DATA
        valor_cell = ws_dash.cell(row=row, column=6, value=item["valor"])
        valor_cell.font = FONT_DATA
        valor_cell.number_format = 'R$ #,##0.00'
        ws_dash.cell(row=row, column=7, value=item["target"]).font = FONT_DATA

        for c in range(2, 8):
            ws_dash.cell(row=row, column=c).border = THIN_BORDER
            ws_dash.cell(row=row, column=c).alignment = ALIGN_LEFT if c > 3 else ALIGN_CENTER

    # ─── ALERTAS DIRETORIA (abaixo da Timeline) ───
    alert_start_row = 14 + max(len(items_with_deadline[:10]), 1) + 2

    ws_dash.merge_cells(f"B{alert_start_row}:J{alert_start_row}")
    ws_dash[f"B{alert_start_row}"] = "ALERTAS DE DIRETORIA — REQUEREM APROVAÇÃO DO ALOISIO"
    ws_dash[f"B{alert_start_row}"].font = Font(name="Segoe UI", size=12, bold=True, color=RED_ALERT)
    ws_dash[f"B{alert_start_row}"].border = Border(bottom=Side(style='medium', color=RED_ALERT))

    dir_headers = ["Órgão / Cliente", "Assunto", "Resumo Executivo", "Valor (R$)", "Prazo", "Status"]
    h_row = alert_start_row + 1
    for i, h in enumerate(dir_headers):
        cell = ws_dash.cell(row=h_row, column=2 + i)
        cell.value = h
        cell.font = Font(name="Segoe UI", size=10, bold=True, color=WHITE)
        cell.fill = PatternFill(start_color=RED_ALERT, end_color=RED_ALERT, fill_type="solid")
        cell.alignment = ALIGN_CENTER
        cell.border = THIN_BORDER

    # Preencher alertas
    board_col_idx = (headers_op.index("Diretoria?") + 1) if "Diretoria?" in headers_op else None
    org_col_idx = (headers_op.index("Órgão / Cliente") + 1) if "Órgão / Cliente" in headers_op else None
    assunto_col_idx = (headers_op.index("Assunto") + 1) if "Assunto" in headers_op else None
    resumo_col_idx = (headers_op.index("Resumo Executivo") + 1) if "Resumo Executivo" in headers_op else None
    valor_col_idx = (headers_op.index("Valor Estimado (R$)") + 1) if "Valor Estimado (R$)" in headers_op else None
    status_col_idx = (headers_op.index("Status") + 1) if "Status" in headers_op else None

    alert_row = h_row + 1
    fill_alert_row = PatternFill(start_color="FFF5F5", end_color="FFF5F5", fill_type="solid")

    if board_col_idx:
        for r in range(2, ws_op.max_row + 1):
            board_val = ws_op.cell(row=r, column=board_col_idx).value
            if board_val == "✅":
                vals = [
                    ws_op.cell(row=r, column=org_col_idx).value if org_col_idx else "",
                    ws_op.cell(row=r, column=assunto_col_idx).value if assunto_col_idx else "",
                    ws_op.cell(row=r, column=resumo_col_idx).value if resumo_col_idx else "",
                    ws_op.cell(row=r, column=valor_col_idx).value if valor_col_idx else "",
                    ws_op.cell(row=r, column=prazo_col_idx).value if prazo_col_idx else "",
                    ws_op.cell(row=r, column=status_col_idx).value if status_col_idx else "",
                ]
                for i, val in enumerate(vals):
                    cell = ws_dash.cell(row=alert_row, column=2 + i)
                    cell.value = val
                    cell.font = FONT_DATA
                    cell.fill = fill_alert_row
                    cell.border = THIN_BORDER
                    if i == 3:
                        cell.number_format = 'R$ #,##0.00'
                        cell.alignment = ALIGN_RIGHT
                    elif i in [4, 5]:
                        cell.alignment = ALIGN_CENTER
                    else:
                        cell.alignment = ALIGN_LEFT
                alert_row += 1

    # ─── LARGURAS DAS COLUNAS DO DASHBOARD ───
    ws_dash.column_dimensions["A"].width = 3
    ws_dash.column_dimensions["B"].width = 18
    ws_dash.column_dimensions["C"].width = 14
    ws_dash.column_dimensions["D"].width = 22
    ws_dash.column_dimensions["E"].width = 35
    ws_dash.column_dimensions["F"].width = 18
    ws_dash.column_dimensions["G"].width = 22
    ws_dash.column_dimensions["H"].width = 18
    ws_dash.column_dimensions["I"].width = 18
    ws_dash.column_dimensions["J"].width = 18


# ============================================================
#  HELPERS DO DASHBOARD
# ============================================================

def _build_kpi_card(ws, col_letter, row, label, value, fill, value_font, num_format=None):
    """Constrói um card de KPI com label + valor."""
    col2 = get_column_letter(ord(col_letter) - 64 + 1)

    # Label
    ws.merge_cells(f"{col_letter}{row}:{col2}{row}")
    cell_label = ws[f"{col_letter}{row}"]
    cell_label.value = label
    cell_label.font = FONT_KPI_LABEL
    cell_label.alignment = ALIGN_CENTER
    cell_label.fill = fill

    # Value
    val_row = row + 1
    ws.merge_cells(f"{col_letter}{val_row}:{col2}{val_row}")
    cell_val = ws[f"{col_letter}{val_row}"]
    cell_val.value = value
    cell_val.font = value_font
    cell_val.alignment = ALIGN_CENTER
    cell_val.fill = fill
    if num_format:
        cell_val.number_format = num_format

    # Borders
    for r in [row, val_row]:
        for c_offset in range(2):
            col_idx = ord(col_letter) - 64 + c_offset
            ws.cell(row=r, column=col_idx).border = ACCENT_BORDER


def _calculate_days_from_formatted(prazo_str: str) -> int:
    """Calcula dias restantes a partir de prazo formatado (DD MMM YYYY ou YYYY-MM-DD)."""
    from routing_tables import MESES_PT
    if not prazo_str:
        return 9999
    try:
        # Tentar formato "DD MMM YYYY"
        parts = str(prazo_str).strip().split()
        if len(parts) >= 3:
            day = int(parts[0])
            month_str = parts[1]
            year = int(parts[2])
            # Lookup reverso do mês
            month = None
            for m_num, m_name in MESES_PT.items():
                if m_name.lower() == month_str.lower():
                    month = m_num
                    break
            if month:
                deadline = datetime(year, month, day, tzinfo=timezone.utc)
                return (deadline - datetime.now(timezone.utc)).days

        # Tentar formato ISO "YYYY-MM-DD"
        deadline = datetime.strptime(str(prazo_str)[:10], "%Y-%m-%d")
        deadline = deadline.replace(tzinfo=timezone.utc)
        return (deadline - datetime.now(timezone.utc)).days
    except (ValueError, TypeError, IndexError):
        return 9999
