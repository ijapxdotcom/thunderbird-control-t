"""
compliance_agent.py — Orquestrador Principal v3.1
Pipeline: Thunderbird MBOX → Routing Engine (Gemini) → PII Masking → Excel Executivo

v3.1 — Issues implementadas:
  #11: Integração com logger.py (logs estruturados + RotatingFileHandler)
  #18: Resiliência no Excel (retry + exponential backoff) — implementado em excel_handler.py
  #19: CLI avançada (--from-date, --area, --dry-run, --log-level)
"""

import os
import argparse
import sys
from datetime import datetime, timezone
from dotenv import load_dotenv

# Logging DEVE ser configurado antes de importar qualquer módulo local
# que use get_logger(), pois eles podem chamar get_logger no nível de módulo.
# O nível do console será ajustado após parse dos args CLI.
from logger import configure_logging, get_logger, get_log_path, log_separator

# Módulos locais
from mbox_parser import parse_mbox, generate_task_id
from pii_masker import sanitize_text
from gemini_pipeline import analyze_email_with_gemini
from excel_handler import create_or_update_excel, load_all_raw_ids
from routing_tables import (
    resolve_theme, resolve_target, check_board_approval,
    build_action_link, format_br_datetime, TARGET_AREAS,
)

# Encoding UTF-8 para terminal Windows
sys.stdout.reconfigure(encoding='utf-8')


# ============================================================
#  BANNER
# ============================================================

BANNER = """
╔══════════════════════════════════════════════════════════╗
║                                                          ║
║   ██████╗ ██████╗ ███╗   ██╗████████╗██████╗  ██████╗    ║
║  ██╔════╝██╔═══██╗████╗  ██║╚══██╔══╝██╔══██╗██╔═══██╗   ║
║  ██║     ██║   ██║██╔██╗ ██║   ██║   ██████╔╝██║   ██║   ║
║  ██║     ██║   ██║██║╚██╗██║   ██║   ██╔══██╗██║   ██║   ║
║  ╚██████╗╚██████╔╝██║ ╚████║   ██║   ██║  ██║╚██████╔╝   ║
║   ╚═════╝ ╚═════╝ ╚═╝  ╚═══╝   ╚═╝   ╚═╝  ╚═╝ ╚═════╝   ║
║                                                          ║
║  CONTROL-T | Compliance & Routing Engine v3.1            ║
║  Gestão Inteligente de Licitações                        ║
║                                                          ║
╚══════════════════════════════════════════════════════════╝
"""


# ============================================================
#  VALIDADORES DE ARGUMENTOS CLI
# ============================================================

def _validate_date(value: str) -> str:
    """Valida que o argumento --from-date está no formato YYYY-MM-DD."""
    try:
        datetime.strptime(value, "%Y-%m-%d")
        return value
    except ValueError:
        raise argparse.ArgumentTypeError(
            f"Data inválida: '{value}'. Use o formato YYYY-MM-DD (ex: 2026-06-19)."
        )


def _validate_area(value: str) -> int:
    """Valida que o argumento --area está no intervalo 1-8."""
    try:
        area_id = int(value)
        if area_id not in TARGET_AREAS or area_id == 0:
            raise ValueError
        return area_id
    except ValueError:
        valid_areas = "\n".join(
            [f"  {k}: {v}" for k, v in TARGET_AREAS.items() if k > 0]
        )
        raise argparse.ArgumentTypeError(
            f"Área '{value}' inválida. Opções válidas:\n{valid_areas}"
        )


# ============================================================
#  MAIN
# ============================================================

def main():
    load_dotenv()

    # ─── CLI ───
    parser = argparse.ArgumentParser(
        description="Control-T — Compliance Agent & Routing Engine v3.1",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Exemplos de uso:
  python compliance_agent.py                         # Produção (últimas 36h)
  python compliance_agent.py --test-mode             # Modo offline com dados simulados
  python compliance_agent.py --mock-gemini --hours 48
  python compliance_agent.py --from-date 2026-06-01  # E-mails a partir de 01/Jun
  python compliance_agent.py --area 3                # Só Engenharia/Pré-vendas
  python compliance_agent.py --dry-run               # Simula sem gravar no Excel
  python compliance_agent.py --log-level DEBUG       # Log detalhado para diagnóstico
        """,
    )

    # Argumentos legados (v3.0) — mantidos para compatibilidade total
    parser.add_argument(
        "--test-mode", "-t", action="store_true",
        help="Executa com dados mockados (offline, sem API Gemini)"
    )
    parser.add_argument(
        "--mock-gemini", "-mg", action="store_true",
        help="Usa e-mails reais mas simula a classificação da IA (economia de tokens)"
    )
    parser.add_argument(
        "--hours", "-hr", type=int,
        help="Janela de ingestão em horas (padrão: lê INGESTION_WINDOW_HOURS do .env ou 36)"
    )

    # Novos argumentos (v3.1 — Issue #19)
    parser.add_argument(
        "--from-date", "-fd", type=_validate_date, metavar="YYYY-MM-DD",
        help=(
            "Processa e-mails a partir de uma data específica (formato: YYYY-MM-DD). "
            "Sobrescreve --hours quando especificado. Ex: --from-date 2026-06-01"
        )
    )
    parser.add_argument(
        "--area", "-a", type=_validate_area, metavar="1-8",
        help=(
            "Filtra resultados para uma área destino específica. "
            "Valores: 1=Comercial 2=Contas 3=Engenharia 4=Operações "
            "5=Jurídico 6=Financeiro 7=TI 8=Diretoria"
        )
    )
    parser.add_argument(
        "--dry-run", "-dr", action="store_true",
        help=(
            "Executa o pipeline completo (parse → Gemini → routing) mas NÃO grava no Excel. "
            "Exibe um resumo completo no terminal. Útil para validação e testes de integração."
        )
    )
    parser.add_argument(
        "--log-level", "-ll",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        default="INFO",
        metavar="NÍVEL",
        help="Nível de log para o console (DEBUG/INFO/WARNING/ERROR). Padrão: INFO"
    )

    args = parser.parse_args()

    # ─── Configurar logging ANTES de qualquer operação ───
    configure_logging(console_level=args.log_level, file_level="DEBUG")
    log = get_logger(__name__)

    print(BANNER)

    log.info(
        "Control-T v3.1 iniciado | modo=%s | log_level=%s | log_file=%s",
        _detect_mode(args), args.log_level, get_log_path(),
    )

    # ─── Configurações de ambiente ───
    excel_path = os.getenv("EXCEL_PATH", "monitoramento_licitacoes.xlsx")
    secure_dir = os.getenv("SECURE_ATTACHMENTS_DIR", "secure_attachments")

    # ─── Calcular janela de ingestão ───
    if args.from_date:
        from_dt = datetime.strptime(args.from_date, "%Y-%m-%d").replace(tzinfo=timezone.utc)
        delta = datetime.now(timezone.utc) - from_dt
        window_hours = max(1, int(delta.total_seconds() / 3600))
        log.info("Modo --from-date: %s → janela de %dh calculada", args.from_date, window_hours)
    elif args.hours:
        window_hours = args.hours
    else:
        try:
            window_hours = int(os.getenv("INGESTION_WINDOW_HOURS", "36"))
        except ValueError:
            window_hours = 36

    # ─── Cabeçalho de status ───
    print(f"  ⏰ Janela de Ingestão: Últimas {window_hours} horas")
    if args.from_date:
        print(f"  📅 A partir de: {args.from_date}")
    print(f"  📊 Planilha: {excel_path}")
    print(f"  📁 Anexos Seguros: {secure_dir}")
    if args.area:
        area_label = TARGET_AREAS.get(args.area, str(args.area))
        print(f"  🔍 Filtro de Área: [{args.area}] {area_label}")
    if args.dry_run:
        print("  🧪 MODO DRY-RUN: Nenhuma alteração será gravada no Excel")
    print(f"  📝 Log: {get_log_path()}")
    print()

    os.makedirs(secure_dir, exist_ok=True)

    # ─── Modo de Execução ───
    if args.test_mode:
        print("  ▶ MODO DE TESTE OFFLINE (dados simulados)\n")
        log.info("Iniciando modo de teste offline")
        run_test_mode(excel_path, area_filter=args.area, dry_run=args.dry_run)
        return

    # ─── Modo Real ───
    zoho_path = os.getenv("ZOHO_INBOX_PATH", "")

    if not zoho_path:
        log.error("ZOHO_INBOX_PATH não configurado no .env")
        print("  ❌ Erro: ZOHO_INBOX_PATH não configurado no .env")
        print("  Execute com --test-mode para validar o sistema.")
        sys.exit(1)

    # Carregar IDs já processados
    processed_ids = load_all_raw_ids(excel_path)
    log.info("IDs já processados carregados: %d registros", len(processed_ids))
    print(f"  📋 Registros já processados: {len(processed_ids)}")

    # Extrair e-mails
    all_emails = []
    if zoho_path and os.path.exists(zoho_path):
        log.info("Processando caixa Zoho: %s", zoho_path)
        print(f"\n  📬 Processando caixa Zoho...")
        all_emails.extend(parse_mbox(zoho_path, window_hours, secure_dir))
    else:
        log.warning("Arquivo INBOX Zoho não encontrado: %s", zoho_path)
        print(f"  ⚠ Arquivo INBOX Zoho não encontrado: {zoho_path}")

    # Filtrar duplicados
    new_emails = [e for e in all_emails if e["task_id"] not in processed_ids]
    log.info(
        "E-mails na janela: %d total, %d novos (após deduplicação)",
        len(all_emails), len(new_emails),
    )
    print(f"\n  📨 E-mails novos na janela: {len(new_emails)}")

    if not new_emails:
        log.info("Nenhum e-mail novo. Atualizando dashboard da planilha existente.")
        print("  ℹ Nenhum e-mail novo. Atualizando planilha existente...")
        if not args.dry_run:
            create_or_update_excel(excel_path, [])
        return

    # ─── Processar cada e-mail ───
    analyzed_items = process_emails(
        new_emails,
        excel_path,
        mock_gemini=args.mock_gemini,
        area_filter=args.area,
    )

    # ─── Dry-run: apenas exibir resumo ───
    if args.dry_run:
        _print_dry_run_summary(analyzed_items, args.area)
        return

    # ─── Gravar na planilha ───
    log.info("Gravando %d registros na planilha: %s", len(analyzed_items), excel_path)
    print(f"\n  💾 Gravando {len(analyzed_items)} registros na planilha...")
    create_or_update_excel(excel_path, analyzed_items)
    log.info("Orquestração concluída com sucesso. %d itens processados.", len(analyzed_items))
    print("\n  ✅ Orquestração concluída com sucesso!")


# ============================================================
#  PROCESSAMENTO DE E-MAILS
# ============================================================

def process_emails(
    emails: list,
    excel_path: str,
    mock_gemini: bool = False,
    area_filter: int | None = None,
) -> list:
    """
    Processa lista de e-mails pelo Routing Engine e retorna payloads formatados.

    Args:
        emails: Lista de dicts de e-mail (saída do parse_mbox)
        excel_path: Caminho da planilha (para calcular paths relativos de anexos)
        mock_gemini: Se True, usa heurísticas locais em vez da API Gemini
        area_filter: Se informado, inclui apenas itens com target_area_id == area_filter
    """
    log = get_logger(__name__)
    analyzed_items = []
    skipped_by_filter = 0

    for idx, email_data in enumerate(emails, start=1):
        log_separator(log, f"E-mail {idx}/{len(emails)}")
        log.info(
            "[%d/%d] De: %s | Assunto: %s",
            idx, len(emails),
            email_data.get("remetente_original", "?"),
            email_data.get("assunto", "?")[:60],
        )

        print(f"\n  ─── [{idx}/{len(emails)}] ──────────────────────────────────")
        print(f"  De: {email_data['remetente_original']}")
        print(f"  Assunto: {email_data['assunto']}")

        body_text = email_data["body"]
        pdf_text = email_data["pdf_text"]

        # 1. Routing Engine (IA ou Mock)
        routing = analyze_email_with_gemini(
            subject=email_data["assunto"],
            sender=email_data["remetente_original"],
            sender_email=email_data.get("sender_email", ""),
            body=body_text,
            pdf_text=pdf_text,
            mock=mock_gemini,
        )

        # ── Aplicar filtro de área (Issue #19) ──
        if area_filter is not None and routing.target_area_id != area_filter:
            skipped_by_filter += 1
            log.debug(
                "E-mail ignorado pelo filtro --area=%d (classificado como área %d): %s",
                area_filter, routing.target_area_id, email_data.get("assunto", "")[:50],
            )
            print(
                f"  ↷ Ignorado pelo filtro --area={area_filter} "
                f"(área detectada: [{routing.target_area_id}] {resolve_target(routing.target_area_id)})"
            )
            continue

        # 2. PII Masking local (camada extra LGPD)
        masked_summary = sanitize_text(routing.executive_summary)
        masked_subject = sanitize_text(routing.assunto_padronizado)

        # 3. Traduzir IDs → strings (via routing_tables.py)
        theme_label = resolve_theme(routing.theme_id)
        target_label = resolve_target(routing.target_area_id)

        # 4. Regras de negócio determinísticas
        requires_board = check_board_approval(
            routing.valor_certame,
            routing.target_area_id,
            routing.theme_id,
        )
        action_link = build_action_link(routing.sender_email, email_data["assunto"])

        # 5. Montar links de anexos
        link_anexo = ""
        if email_data["attachments"]:
            excel_dir = os.path.dirname(excel_path) or "."
            try:
                rel_paths = [os.path.relpath(p, start=excel_dir) for p in email_data["attachments"]]
                link_anexo = ", ".join(rel_paths)
            except ValueError:
                link_anexo = ", ".join(email_data["attachments"])

        # 6. Extrair remetente raiz (nome do órgão via domínio)
        remetente_raiz = _extract_org_name(routing.sender_email)

        # 7. Payload final
        item = {
            "task_id": email_data["task_id"],
            "data_ingestao": email_data["data_ingestao"],
            "remetente_raiz": remetente_raiz,
            "theme_id": routing.theme_id,
            "theme_label": theme_label,
            "target_area_id": routing.target_area_id,
            "target_label": target_label,
            "assunto_padronizado": masked_subject,
            "executive_summary": masked_summary,
            "valor_certame": routing.valor_certame,
            "prazo_fatal": routing.prazo_fatal,
            "status_workflow": "Novo",
            "requires_board_approval": requires_board,
            "action_link": action_link,
            "link_anexo_seguro": link_anexo,
        }

        # ── Console output ──
        board_icon = "🔴 DIRETORIA" if requires_board else "⚪"
        print(f"  Tema: [{routing.theme_id}] {theme_label}")
        print(f"  Destino: [{routing.target_area_id}] {target_label} {board_icon}")
        if routing.valor_certame:
            print(f"  Valor: R$ {routing.valor_certame:,.2f}")
        if routing.prazo_fatal:
            print(f"  Prazo: {routing.prazo_fatal}")
        print(f"  Resumo: {masked_summary[:80]}...")

        log.info(
            "Classificado: tema=%d (%s) | area=%d (%s) | board=%s | valor=%s",
            routing.theme_id, theme_label,
            routing.target_area_id, target_label,
            requires_board,
            f"R$ {routing.valor_certame:,.2f}" if routing.valor_certame else "N/A",
        )

        analyzed_items.append(item)

    if area_filter is not None and skipped_by_filter > 0:
        log.info(
            "Filtro --area=%d aplicado: %d ignorado(s), %d incluído(s).",
            area_filter, skipped_by_filter, len(analyzed_items),
        )
        print(
            f"\n  🔍 Filtro ativo: {skipped_by_filter} e-mail(s) ignorado(s) "
            f"(fora da área {area_filter}), {len(analyzed_items)} incluído(s)."
        )

    return analyzed_items


# ============================================================
#  DRY-RUN: RESUMO NO TERMINAL (sem gravar no Excel)
# ============================================================

def _print_dry_run_summary(items: list, area_filter: int | None) -> None:
    """Exibe um resumo tabular dos itens analisados no modo --dry-run."""
    log = get_logger(__name__)

    area_info = f" | Filtro de Área: {area_filter}" if area_filter else ""
    print(f"\n  {'=' * 60}")
    print(f"  DRY-RUN — RESUMO DO PIPELINE (NENHUM DADO FOI GRAVADO)")
    print(f"  Itens analisados: {len(items)}{area_info}")
    print(f"  {'=' * 60}")

    if not items:
        print("  (Nenhum item — verifique os filtros aplicados)")
        return

    for i, item in enumerate(items, 1):
        board = "🔴 DIRETORIA" if item.get("requires_board_approval") else "⚪"
        valor_str = f"R$ {item['valor_certame']:,.2f}" if item.get("valor_certame") else "N/A"
        print(f"\n  [{i}] {item.get('assunto_padronizado', 'N/A')[:60]}")
        print(f"       Órgão:   {item.get('remetente_raiz', 'N/A')}")
        print(f"       Tema:    [{item['theme_id']}] {item['theme_label']}")
        print(f"       Área:    [{item['target_area_id']}] {item['target_label']} {board}")
        print(f"       Valor:   {valor_str}")
        print(f"       Prazo:   {item.get('prazo_fatal') or 'N/A'}")
        print(f"       Resumo:  {item.get('executive_summary', '')[:80]}")

    board_count = sum(1 for i in items if i.get("requires_board_approval"))
    total_value = sum(i.get("valor_certame") or 0 for i in items)
    print(f"\n  {'─' * 60}")
    print(f"  Total: {len(items)} itens | Diretoria: {board_count} | Volume: R$ {total_value:,.2f}")
    print(f"  {'=' * 60}\n")

    log.info(
        "Dry-run concluído: %d itens | %d Diretoria | Volume total: R$ %.2f",
        len(items), board_count, total_value,
    )


# ============================================================
#  UTILITÁRIOS
# ============================================================

def _extract_org_name(email_addr: str) -> str:
    """Extrai nome do órgão a partir do domínio do e-mail."""
    if not email_addr or "@" not in email_addr:
        return "Órgão Não Identificado"
    domain = email_addr.split("@")[-1].strip()
    org = domain.split(".")[0].upper()
    known_orgs = {
        "SERPRO": "SERPRO",
        "TJES": "TJES - Tribunal de Justiça ES",
        "RECEITA": "Receita Federal",
        "FAZENDA": "Ministério da Fazenda",
        "INTELBRAS": "Intelbras",
        "HIKVISION": "Hikvision",
    }
    return known_orgs.get(org, org)


def _detect_mode(args: argparse.Namespace) -> str:
    """Detecta o modo de execução para o log inicial."""
    if args.test_mode:
        return "test-mode"
    if args.dry_run:
        return "dry-run"
    if args.mock_gemini:
        return "mock-gemini"
    return "production"


# ============================================================
#  MODO DE TESTE (Offline)
# ============================================================

def run_test_mode(
    excel_path: str,
    area_filter: int | None = None,
    dry_run: bool = False,
) -> None:
    """Pipeline completo com 6 e-mails sintéticos cobrindo todos os temas."""
    log = get_logger(__name__)
    test_dir = "secure_attachments_test"
    os.makedirs(test_dir, exist_ok=True)

    mock_emails = [
        {
            "task_id": generate_task_id("test_001_spam"),
            "data_ingestao": datetime.now(timezone.utc).isoformat(),
            "remetente_original": "marketing@eventos-tecnologia.com.br",
            "sender_email": "marketing@eventos-tecnologia.com.br",
            "assunto": "Fwd: Webinar Gratuito: Tendências de IA 2026",
            "body": "Olá parceiro, não perca nosso workshop online gratuito dia 25/06 sobre inovação.",
            "pdf_text": "",
            "attachments": [],
        },
        {
            "task_id": generate_task_id("test_002_cotacao"),
            "data_ingestao": datetime.now(timezone.utc).isoformat(),
            "remetente_original": "compras@serpro.gov.br",
            "sender_email": "compras@serpro.gov.br",
            "assunto": "Solicitação de Cotação de Preços - Pregão Eletrônico nº 140/2026",
            "body": "Prezados,\nSolicitamos cotação para fornecimento de 48 câmeras Intelbras VIP 3230. Valor de referência: R$ 1.250.000,00. Enviar proposta até 22/06/2026. Dados de faturamento: Banco do Brasil Ag 3399-5 CC 55432-1.",
            "pdf_text": "",
            "attachments": [],
        },
        {
            "task_id": generate_task_id("test_003_aditivo"),
            "data_ingestao": datetime.now(timezone.utc).isoformat(),
            "remetente_original": "tjes.contratos@tjes.jus.br",
            "sender_email": "tjes.contratos@tjes.jus.br",
            "assunto": "Re: Liberação de Termo Aditivo - Contrato nº 45/2024",
            "body": "Prezados,\nEnviamos o Termo Aditivo nº 3, aditamento de R$ 845.300,00 no contrato de manutenção de CFTV. Solicitamos assinatura digital do representante legal até 26/06/2026.",
            "pdf_text": "TERMO ADITIVO - Contrato CFTV. Partes: TJES e Mahvla. Valor: R$ 845.300,00. Fiscal: José Carlos, CPF: 111.222.333-44.",
            "attachments": ["minuta_aditivo_45_2024.pdf"],
        },
        {
            "task_id": generate_task_id("test_004_penalidade"),
            "data_ingestao": datetime.now(timezone.utc).isoformat(),
            "remetente_original": "notificacoes@receita.fazenda.gov.br",
            "sender_email": "notificacoes@receita.fazenda.gov.br",
            "assunto": "Notificação de Penalidade Administrativa - Pregão 89/2025",
            "body": "NOTIFICAÇÃO DE MULTA.\nDescumprimento de SLA. Multa de R$ 5.200.000,00 à Mahvla Telecomm. Prazo de defesa: até 23/06/2026. CPF: 444.555.666-77. Agência 1004, Conta 9876-5.",
            "pdf_text": "",
            "attachments": [],
        },
        {
            "task_id": generate_task_id("test_005_tecnico"),
            "data_ingestao": datetime.now(timezone.utc).isoformat(),
            "remetente_original": "comissao.licitacao@tjes.jus.br",
            "sender_email": "comissao.licitacao@tjes.jus.br",
            "assunto": "Fwd: Pedido de Esclarecimento Técnico - Pregão nº 12/2026",
            "body": "Prezada engenharia, encaminhamos esclarecimento sobre CFTV Hikvision. As fibras devem ter dupla abordagem? Responder até 20/06/2026 às 18h.",
            "pdf_text": "Pergunta: As fibras devem ter dupla abordagem física até a central?",
            "attachments": ["esclarecimento_tecnico_12_2026.pdf"],
        },
        {
            "task_id": generate_task_id("test_006_governanca"),
            "data_ingestao": datetime.now(timezone.utc).isoformat(),
            "remetente_original": "dpo@mahvla.com.br",
            "sender_email": "dpo@mahvla.com.br",
            "assunto": "Auditoria Interna LGPD - Relatório Q2 2026",
            "body": "Prezados, segue relatório de auditoria LGPD referente ao Q2/2026. Identificamos 3 pontos de atenção. Prazo para plano de ação corretiva: 30/06/2026.",
            "pdf_text": "",
            "attachments": [],
        },
    ]

    # Simular anexos físicos
    for em in mock_emails:
        saved = []
        for att in em["attachments"]:
            task_dir = os.path.join(test_dir, em["task_id"])
            os.makedirs(task_dir, exist_ok=True)
            fpath = os.path.join(task_dir, att)
            with open(fpath, "w", encoding="utf-8") as f:
                f.write(f"Conteúdo de teste: {att}\n{em['pdf_text']}")
            saved.append(fpath)
        em["attachments"] = saved

    log.info("Modo de teste: %d e-mails sintéticos gerados", len(mock_emails))

    # Processar
    analyzed = process_emails(
        mock_emails,
        excel_path,
        mock_gemini=True,
        area_filter=area_filter,
    )

    # Dry-run ou gravação real
    if dry_run:
        _print_dry_run_summary(analyzed, area_filter)
    else:
        log.info("Gravando %d registros de teste na planilha", len(analyzed))
        print(f"\n  💾 Gravando {len(analyzed)} registros de teste...")
        create_or_update_excel(excel_path, analyzed)
        log.info("Teste concluído. Planilha: %s", excel_path)
        print(f"\n  ✅ Teste concluído! Planilha gerada: {excel_path}")
        print(f"  📁 Anexos de teste: {test_dir}/")


if __name__ == "__main__":
    main()
