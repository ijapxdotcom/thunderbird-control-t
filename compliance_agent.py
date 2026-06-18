"""
compliance_agent.py — Orquestrador Principal v3.0
Pipeline: Thunderbird MBOX → Routing Engine (Gemini) → PII Masking → Excel Executivo
"""

import os
import argparse
import sys
from datetime import datetime, timezone
from dotenv import load_dotenv

# Módulos locais
from mbox_parser import parse_mbox, generate_task_id
from pii_masker import sanitize_text
from gemini_pipeline import analyze_email_with_gemini
from excel_handler import create_or_update_excel, load_all_raw_ids
from routing_tables import (
    resolve_theme, resolve_target, check_board_approval,
    build_action_link, format_br_datetime,
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
║  CONTROL-T | Compliance & Routing Engine v3.0            ║
║  Gestão Inteligente de Licitações                        ║
║                                                          ║
╚══════════════════════════════════════════════════════════╝
"""


def main():
    load_dotenv()

    # CLI
    parser = argparse.ArgumentParser(
        description="Control-T — Compliance Agent & Routing Engine v3.0"
    )
    parser.add_argument("--test-mode", "-t", action="store_true",
                        help="Executa com dados mockados (offline, sem API)")
    parser.add_argument("--mock-gemini", "-mg", action="store_true",
                        help="Usa e-mails reais mas simula a classificação da IA")
    parser.add_argument("--hours", "-hr", type=int,
                        help="Janela de ingestão em horas (padrão: 36)")
    args = parser.parse_args()

    print(BANNER)

    # ─── Configurações ───
    excel_path = os.getenv("EXCEL_PATH", "monitoramento_licitacoes.xlsx")
    secure_dir = os.getenv("SECURE_ATTACHMENTS_DIR", "secure_attachments")

    if args.hours:
        window_hours = args.hours
    else:
        try:
            window_hours = int(os.getenv("INGESTION_WINDOW_HOURS", "36"))
        except ValueError:
            window_hours = 36

    print(f"  ⏰ Janela de Ingestão: Últimas {window_hours} horas")
    print(f"  📊 Planilha: {excel_path}")
    print(f"  📁 Anexos Seguros: {secure_dir}")
    print()

    os.makedirs(secure_dir, exist_ok=True)

    # ─── Modo de Execução ───
    if args.test_mode:
        print("  ▶ MODO DE TESTE OFFLINE (dados simulados)\n")
        run_test_mode(excel_path)
        return

    # ─── Modo Real ───
    zoho_path = os.getenv("ZOHO_INBOX_PATH", "")

    if not zoho_path:
        print("  ❌ Erro: ZOHO_INBOX_PATH não configurado no .env")
        print("  Execute com --test-mode para validar o sistema.")
        sys.exit(1)

    # Carregar IDs já processados
    processed_ids = load_all_raw_ids(excel_path)
    print(f"  📋 Registros já processados: {len(processed_ids)}")

    # Extrair e-mails
    all_emails = []
    if zoho_path and os.path.exists(zoho_path):
        print(f"\n  📬 Processando caixa Zoho...")
        all_emails.extend(parse_mbox(zoho_path, window_hours, secure_dir))
    else:
        print(f"  ⚠ Arquivo INBOX Zoho não encontrado: {zoho_path}")

    # Filtrar duplicados
    new_emails = [e for e in all_emails if e["task_id"] not in processed_ids]
    print(f"\n  📨 E-mails novos na janela: {len(new_emails)}")

    if not new_emails:
        print("  ℹ Nenhum e-mail novo. Atualizando planilha existente...")
        create_or_update_excel(excel_path, [])
        return

    # ─── Processar cada e-mail ───
    analyzed_items = process_emails(new_emails, excel_path, args.mock_gemini)

    # ─── Gravar na planilha ───
    print(f"\n  💾 Gravando {len(analyzed_items)} registros na planilha...")
    create_or_update_excel(excel_path, analyzed_items)
    print("\n  ✅ Orquestração concluída com sucesso!")


def process_emails(emails: list, excel_path: str, mock_gemini: bool = False) -> list:
    """Processa lista de e-mails pelo Routing Engine e retorna payloads formatados."""
    analyzed_items = []

    for idx, email_data in enumerate(emails, start=1):
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

        # Console output
        board_icon = "🔴 DIRETORIA" if requires_board else "⚪"
        print(f"  Tema: [{routing.theme_id}] {theme_label}")
        print(f"  Destino: [{routing.target_area_id}] {target_label} {board_icon}")
        if routing.valor_certame:
            print(f"  Valor: R$ {routing.valor_certame:,.2f}")
        if routing.prazo_fatal:
            print(f"  Prazo: {routing.prazo_fatal}")
        print(f"  Resumo: {masked_summary[:80]}...")

        analyzed_items.append(item)

    return analyzed_items


def _extract_org_name(email_addr: str) -> str:
    """Extrai nome do órgão a partir do domínio do e-mail."""
    if not email_addr or "@" not in email_addr:
        return "Órgão Não Identificado"
    domain = email_addr.split("@")[-1].strip()
    # Remover extensões comuns
    org = domain.split(".")[0].upper()
    # Mapear domínios conhecidos
    known_orgs = {
        "SERPRO": "SERPRO",
        "TJES": "TJES - Tribunal de Justiça ES",
        "RECEITA": "Receita Federal",
        "FAZENDA": "Ministério da Fazenda",
        "INTELBRAS": "Intelbras",
        "HIKVISION": "Hikvision",
    }
    return known_orgs.get(org, org)


# ============================================================
#  MODO DE TESTE (Offline)
# ============================================================

def run_test_mode(excel_path: str):
    """Pipeline completo com 6 e-mails sintéticos cobrindo todos os temas."""
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
            "body": "Prezados,\nSolicitamos cotação para fornecimento de 48 câmeras Intelbras VIP 3230 e switches ópticos. Valor de referência: R$ 1.250.000,00. Enviar proposta até 22/06/2026. Dados de faturamento: Banco do Brasil Ag 3399-5 CC 55432-1.",
            "pdf_text": "",
            "attachments": [],
        },
        {
            "task_id": generate_task_id("test_003_aditivo"),
            "data_ingestao": datetime.now(timezone.utc).isoformat(),
            "remetente_original": "tjes.contratos@tjes.jus.br",
            "sender_email": "tjes.contratos@tjes.jus.br",
            "assunto": "Re: Liberação de Termo Aditivo - Contrato nº 45/2024",
            "body": "Prezados,\nEnviamos o Termo Aditivo nº 3, aditamento financeiro de R$ 845.300,00 no contrato de manutenção de CFTV. Solicitamos assinatura digital do representante legal até 26/06/2026.",
            "pdf_text": "TERMO ADITIVO - Contrato de suporte técnico CFTV. Partes: TJES e Mahvla. Valor: R$ 845.300,00. Fiscal: José Carlos, CPF: 111.222.333-44.",
            "attachments": ["minuta_aditivo_45_2024.pdf"],
        },
        {
            "task_id": generate_task_id("test_004_penalidade"),
            "data_ingestao": datetime.now(timezone.utc).isoformat(),
            "remetente_original": "notificacoes@receita.fazenda.gov.br",
            "sender_email": "notificacoes@receita.fazenda.gov.br",
            "assunto": "Notificação de Penalidade Administrativa - Pregão 89/2025",
            "body": "NOTIFICAÇÃO DE MULTA.\nDescumprimento de SLA contratual. Multa de R$ 5.200.000,00 à Mahvla Telecomm. Prazo de defesa: 5 dias úteis (até 23/06/2026). CPF do servidor: 444.555.666-77. Caixa Econômica, Agência 1004, Conta 9876-5.",
            "pdf_text": "",
            "attachments": [],
        },
        {
            "task_id": generate_task_id("test_005_tecnico"),
            "data_ingestao": datetime.now(timezone.utc).isoformat(),
            "remetente_original": "comissao.licitacao@tjes.jus.br",
            "sender_email": "comissao.licitacao@tjes.jus.br",
            "assunto": "Fwd: Pedido de Esclarecimento Técnico - Pregão nº 12/2026",
            "body": "Prezada engenharia, encaminhamos esclarecimento técnico sobre CFTV Hikvision. As fibras fornecidas devem ter dupla abordagem? Responder até 20/06/2026 às 18h.",
            "pdf_text": "Pergunta: As fibras devem ter dupla abordagem física até a central?",
            "attachments": ["esclarecimento_tecnico_12_2026.pdf"],
        },
        {
            "task_id": generate_task_id("test_006_governanca"),
            "data_ingestao": datetime.now(timezone.utc).isoformat(),
            "remetente_original": "dpo@mahvla.com.br",
            "sender_email": "dpo@mahvla.com.br",
            "assunto": "Auditoria Interna LGPD - Relatório Q2 2026",
            "body": "Prezados, segue relatório de auditoria de conformidade LGPD referente ao Q2/2026. Identificamos 3 pontos de atenção no tratamento de dados de segurança eletrônica. Prazo para plano de ação corretiva: 30/06/2026.",
            "pdf_text": "",
            "attachments": [],
        },
    ]

    # Simular anexos
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

    # Processar
    analyzed = process_emails(mock_emails, excel_path, mock_gemini=True)

    # Gravar
    print(f"\n  💾 Gravando {len(analyzed)} registros de teste...")
    create_or_update_excel(excel_path, analyzed)
    print(f"\n  ✅ Teste concluído! Planilha gerada: {excel_path}")
    print(f"  📁 Anexos de teste: {test_dir}/")


if __name__ == "__main__":
    main()
