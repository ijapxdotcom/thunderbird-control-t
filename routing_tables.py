"""
routing_tables.py — Single Source of Truth
Tabelas de tradução ID → String para o Routing Engine.
Zero tokens consumidos. Toda lógica de negócio determinística vive aqui.
"""

import re
import email.utils as email_utils
from datetime import datetime, timezone

# ============================================================
#  TABELA P: TEMAS (theme_id → label)
# ============================================================
THEMES = {
    0: "Informativo / Não Relevante",
    1: "Oportunidades / Cotações",
    2: "Editais / Atas de Registro",
    3: "Contratos / Aditivos",
    4: "Técnico / Arquitetura",
    5: "Financeiro / Faturamento",
    6: "Governança / DPO",
}

# ============================================================
#  TABELA Q: ÁREAS-ALVO (target_area_id → label)
# ============================================================
TARGET_AREAS = {
    0: "Descartado",
    1: "Inteligência Comercial",
    2: "Gestão de Contas",
    3: "Engenharia / Pré-vendas",
    4: "Operações",
    5: "Jurídico / Compliance",
    6: "Financeiro",
    7: "TI / Segurança da Informação",
    8: "Diretoria (Aloisio)",
}

# ============================================================
#  WORKFLOW STATES (state machine)
# ============================================================
WORKFLOW_STATES = [
    "Novo",
    "Em Triagem",
    "Em Andamento",
    "Enviado p/ Aprovação",
    "Aguardando Retorno",
    "Finalizado",
]

# ============================================================
#  CORES DE TEMA (para formatação condicional no Excel)
# ============================================================
THEME_COLORS = {
    0: "E0E0E0",  # Cinza — Informativo
    1: "E8F5E9",  # Verde claro — Oportunidades
    2: "E3F2FD",  # Azul claro — Editais
    3: "FFF9C4",  # Amarelo claro — Contratos
    4: "F5F5F5",  # Cinza leve — Técnico
    5: "FFF3E0",  # Laranja claro — Financeiro
    6: "FFEBEE",  # Vermelho claro — Governança
}

# ============================================================
#  CONSTANTES DE NEGÓCIO
# ============================================================
BOARD_THRESHOLD = 5_000_000.0  # R$ 5 milhões — aciona flag Diretoria

# ============================================================
#  MESES PT-BR (para formato de data sem depender de locale)
# ============================================================
MESES_PT = {
    1: "Jan", 2: "Fev", 3: "Mar", 4: "Abr", 5: "Mai", 6: "Jun",
    7: "Jul", 8: "Ago", 9: "Set", 10: "Out", 11: "Nov", 12: "Dez",
}


# ============================================================
#  FUNÇÕES DE TRADUÇÃO
# ============================================================

def resolve_theme(theme_id: int) -> str:
    """Traduz theme_id para label legível. Fallback seguro."""
    if theme_id not in THEMES:
        print(f"[WARN] theme_id={theme_id} fora do range. Clamp para 1.")
        theme_id = 1
    return THEMES[theme_id]


def resolve_target(target_area_id: int) -> str:
    """Traduz target_area_id para label legível. Fallback seguro."""
    if target_area_id not in TARGET_AREAS:
        print(f"[WARN] target_area_id={target_area_id} fora do range. Clamp para 1.")
        target_area_id = 1
    return TARGET_AREAS[target_area_id]


def get_theme_color(theme_id: int) -> str:
    """Retorna o código hexadecimal da cor associada ao tema."""
    return THEME_COLORS.get(theme_id, "FFFFFF")


# ============================================================
#  REGRAS DE NEGÓCIO
# ============================================================

def check_board_approval(valor_certame, target_area_id: int, theme_id: int) -> bool:
    """
    Regra Determinística de Aprovação da Diretoria.
    Aciona se:
      1. Valor > R$ 5 milhões, OU
      2. target_area_id == 8 (Diretoria / Aloisio), OU
      3. theme_id == 6 (Governança / DPO) — risco institucional
    """
    if target_area_id == 8:
        return True
    if theme_id == 6:
        return True
    if valor_certame is not None:
        try:
            if float(valor_certame) >= BOARD_THRESHOLD:
                return True
        except (ValueError, TypeError):
            pass
    return False


def build_action_link(sender_email: str, subject: str) -> str:
    """
    Gera mailto:// link para resposta com 1 clique.
    Retorna string vazia se o e-mail não for válido.
    """
    if not sender_email or "@" not in sender_email:
        return ""

    # Limpar e-mail
    clean_email = sender_email.strip().strip("<>")

    # Limpar assunto para URL encoding básico
    clean_subject = subject.strip()
    for prefix in ["fwd:", "re:", "res:", "enc:", "fw:"]:
        if clean_subject.lower().startswith(prefix):
            clean_subject = clean_subject[len(prefix):].strip()

    encoded_subject = f"RES: {clean_subject}"
    body = "Prezados, acusamos o recebimento da demanda e ela foi encaminhada para análise interna."

    return f"mailto:{clean_email}?subject={encoded_subject}&body={body}"


def extract_pure_email(from_header: str) -> str:
    """
    Extrai o endereço de e-mail puro de um header From.
    Ex: 'João Silva <joao@empresa.com>' → 'joao@empresa.com'
    """
    if not from_header:
        return ""
    _, addr = email_utils.parseaddr(from_header)
    return addr.strip() if addr else ""


# ============================================================
#  FORMATAÇÃO DE DATA PT-BR
# ============================================================

def format_br_datetime(iso_str: str) -> str:
    """
    Converte ISO timestamp para formato brasileiro legível.
    '2026-06-18T14:35:00+00:00' → '18 Jun 2026, 14:35'
    """
    if not iso_str:
        return ""
    try:
        dt = datetime.fromisoformat(iso_str)
        mes = MESES_PT.get(dt.month, str(dt.month))
        return f"{dt.day:02d} {mes} {dt.year}, {dt.hour:02d}:{dt.minute:02d}"
    except (ValueError, TypeError):
        return str(iso_str)[:19]  # Fallback: retorna como está


def format_br_date(date_str: str) -> str:
    """
    Converte data YYYY-MM-DD para formato brasileiro legível.
    '2026-06-22' → '22 Jun 2026'
    """
    if not date_str:
        return ""
    try:
        dt = datetime.strptime(date_str[:10], "%Y-%m-%d")
        mes = MESES_PT.get(dt.month, str(dt.month))
        return f"{dt.day:02d} {mes} {dt.year}"
    except (ValueError, TypeError):
        return str(date_str)


def calculate_sla_days(prazo_fatal: str) -> int:
    """
    Calcula dias restantes até o prazo fatal.
    Retorna negativo se já venceu. Retorna 9999 se não há prazo.
    """
    if not prazo_fatal:
        return 9999
    try:
        deadline = datetime.strptime(prazo_fatal[:10], "%Y-%m-%d")
        deadline = deadline.replace(tzinfo=timezone.utc)
        now = datetime.now(timezone.utc)
        delta = (deadline - now).days
        return delta
    except (ValueError, TypeError):
        return 9999
