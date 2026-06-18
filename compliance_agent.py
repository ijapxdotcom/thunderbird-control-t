import os
import argparse
import sys
from datetime import datetime, timezone
from dotenv import load_dotenv

# Import local modules
from mbox_parser import parse_mbox, generate_task_id
from pii_masker import sanitize_text
from gemini_pipeline import analyze_email_with_gemini
from excel_handler import create_or_update_excel, load_all_raw_ids

# Ensure output is encoded in UTF-8 to prevent Windows terminal display errors
sys.stdout.reconfigure(encoding='utf-8')

def main():
    load_dotenv()
    
    # Parse CLI Arguments
    parser = argparse.ArgumentParser(description="Compliance Agent & Bidding Email Orchestrator")
    parser.add_argument("--test-mode", "-t", action="store_true", help="Runs the pipeline in offline test mode using generated mockup data.")
    parser.add_argument("--mock-gemini", "-mg", action="store_true", help="Uses real local emails but mocks the Gemini API calls.")
    parser.add_argument("--hours", "-hr", type=int, help="Overrides the email ingestion window in hours (default from .env or 36).")
    args = parser.parse_args()

    print("=" * 60)
    print("      COMPLIANCE AGENT & ORQUESTRADOR DE LICITAÇÕES")
    print("=" * 60)

    # 1. Environment Configurations
    excel_path = os.getenv("EXCEL_PATH", "monitoramento_licitacoes.xlsx")
    secure_dir = os.getenv("SECURE_ATTACHMENTS_DIR", "secure_attachments")
    
    # Hour window selection
    if args.hours:
        window_hours = args.hours
    else:
        try:
            window_hours = int(os.getenv("INGESTION_WINDOW_HOURS", "36"))
        except ValueError:
            window_hours = 36
            
    print(f"Janela de Ingestão: Últimas {window_hours} horas")
    print(f"Planilha de Monitoramento: {excel_path}")
    print(f"Diretório de Anexos Seguros: {secure_dir}")
    
    os.makedirs(secure_dir, exist_ok=True)

    # 2. Run Modes
    if args.test_mode:
        print("\n>>> EXECUÇÃO EM MODO DE TESTE OFFLINE (Mockup Data) <<<")
        run_test_mode(excel_path)
        return

    # Real run mode - read Thunderbird files
    zoho_path = os.getenv("ZOHO_INBOX_PATH", "")
    outlook_path = os.getenv("OUTLOOK_INBOX_PATH", "")
    
    if not zoho_path and not outlook_path:
        print("Erro: Nenhum caminho de MBOX configurado em ZOHO_INBOX_PATH ou OUTLOOK_INBOX_PATH no arquivo .env.")
        print("Por favor, configure o arquivo .env ou execute com --test-mode para validar.")
        sys.exit(1)

    # Load already processed task_ids from RAW sheet to avoid re-processing
    processed_ids = load_all_raw_ids(excel_path)
    print(f"Identificadores já processados carregados da planilha: {len(processed_ids)}")

    # Extract emails
    all_emails = []
    
    if zoho_path and os.path.exists(zoho_path):
        print(f"\nProcessando caixa Zoho...")
        all_emails.extend(parse_mbox(zoho_path, window_hours, secure_dir))
    else:
        if zoho_path:
            print(f"Aviso: Arquivo Zoho INBOX não encontrado em: {zoho_path}")
            
    if outlook_path and os.path.exists(outlook_path):
        print(f"\nProcessando caixa Outlook...")
        all_emails.extend(parse_mbox(outlook_path, window_hours, secure_dir))
    else:
        if outlook_path:
            print(f"Aviso: Arquivo Outlook INBOX não encontrado em: {outlook_path}")

    # Remove emails that are already processed (skip duplicate checking)
    new_emails = [e for e in all_emails if e["task_id"] not in processed_ids]
    print(f"\nTotal de e-mails novos encontrados na janela temporal: {len(new_emails)}")

    if not new_emails:
        print("Nenhum e-mail novo para processar nesta execução.")
        # Regenerate Dashboard and operational queue structure to keep it fresh
        create_or_update_excel(excel_path, [])
        return

    # Process each email through Gemini + Local LGPD Data Masking
    analyzed_items = []
    
    for idx, email_data in enumerate(new_emails, start=1):
        print(f"\n[{idx}/{len(new_emails)}] Processando E-mail ID: {email_data['task_id']}")
        print(f"De: {email_data['remetente_original']}")
        print(f"Assunto: {email_data['assunto']}")
        
        # 1. Sanitizar corpo de e-mail localmente (Camada extra de segurança antes ou depois)
        # Passaremos para a IA o texto original, mas a IA fará o mascaramento.
        # Caso ocorra falha na IA ou mock, usamos o sanitizador local.
        body_text = email_data["body"]
        pdf_text = email_data["pdf_text"]
        
        # 2. IA / LLM Pipeline
        analysis = analyze_email_with_gemini(
            subject=email_data["assunto"],
            sender=email_data["remetente_original"],
            body=body_text,
            pdf_text=pdf_text,
            mock=args.mock_gemini
        )
        
        # 3. Post-processing Sanitization (Garante conformidade estrita de dados gravados)
        masked_resumo = sanitize_text(analysis.resumo_executivo)
        masked_assunto = sanitize_text(analysis.assunto_padronizado)
        masked_remetente = sanitize_text(analysis.remetente_raiz)
        
        # Formulate links for Excel
        link_anexo = ""
        if email_data["attachments"]:
            # Excel HYPERLINK format or simple text link pointing to local directory
            # We save absolute paths or relative paths. Relative is safer for transferability.
            rel_paths = [os.path.relpath(p, start=os.path.dirname(excel_path)) for p in email_data["attachments"]]
            link_anexo = ", ".join(rel_paths)

        item_payload = {
            "task_id": email_data["task_id"],
            "data_ingestao": email_data["data_ingestao"],
            "remetente_raiz": masked_remetente,
            "categoria_ia": analysis.categoria_ia,
            "assunto_padronizado": masked_assunto,
            "resumo_executivo": masked_resumo,
            "valor_certame": analysis.valor_certame,
            "prazo_fatal": analysis.prazo_fatal,
            "roteamento_responsavel": analysis.roteamento_responsavel,
            "link_anexo_seguro": link_anexo
        }
        
        print(f"Classificação: {item_payload['categoria_ia']} -> Responsável: {item_payload['roteamento_responsavel']}")
        if item_payload['valor_certame']:
            print(f"Valor Extraído: R$ {item_payload['valor_certame']:,.2f}")
        if item_payload['prazo_fatal']:
            print(f"Prazo Limite: {item_payload['prazo_fatal']}")
            
        analyzed_items.append(item_payload)

    # 3. Update spreadsheet database
    print("\nGravando dados analisados na planilha mestre...")
    create_or_update_excel(excel_path, analyzed_items)
    print("Orquestração concluída com sucesso!")

def run_test_mode(excel_path):
    """
    Simulates a database pipeline ingestion using synthetic bid emails.
    Tests all features without Thunderbird files and without requiring a live Gemini API key.
    """
    import shutil
    
    # Create test directory for attachments
    test_attachments_dir = "secure_attachments_test"
    os.makedirs(test_attachments_dir, exist_ok=True)
    
    # 5 Mock Emails representing different categories and PII risk profiles
    mock_emails = [
        {
            "task_id": generate_task_id("msg_001_spam"),
            "data_ingestao": datetime.now(timezone.utc).isoformat(),
            "remetente_original": "marketing@eventos-tecnologia.com.br",
            "assunto": "Fwd: Participe do Webinar: Tendências de IA e Licitações em 2026",
            "body": "Olá parceiro, não perca nosso workshop online gratuito dia 25/06 sobre inovação na administração pública. Inscreva-se já!",
            "pdf_text": "",
            "attachments": []
        },
        {
            "task_id": generate_task_id("msg_002_cotacao"),
            "data_ingestao": datetime.now(timezone.utc).isoformat(),
            "remetente_original": "compras@serpro.gov.br",
            "assunto": "Solicitação de Cotação de Preços - Pregão Eletrônico nº 140/2026",
            "body": "Prezados,\nSolicitamos cotação de preços para fornecimento de switches ópticos. O valor de referência estimado do certame é de R$ 1.250.000,00. Enviar proposta para o e-mail do pregoeiro joao.silva@serpro.gov.br até 22/06/2026. Dados de faturamento da repartição: Banco do Brasil Ag 3399-5 CC 55432-1.",
            "pdf_text": "",
            "attachments": []
        },
        {
            "task_id": generate_task_id("msg_003_aditivo"),
            "data_ingestao": datetime.now(timezone.utc).isoformat(),
            "remetente_original": "tjes.contratos@tjes.jus.br",
            "assunto": "Re: Liberação de Termo Aditivo Contratual - Contrato Prestação de Serviços nº 45/2024",
            "body": "Prezados,\nEnviamos anexo o Termo Aditivo nº 3 correspondente ao aditamento financeiro de R$ 845.300,00 no contrato de manutenção. Solicitamos assinatura digital do representante legal e envio da via assinada até o dia 26/06/2026.",
            "pdf_text": "CONTEÚDO DO TERMO ADITIVO PDF:\nAditamento ao contrato de suporte técnico. Partes: Tribunal de Justiça do ES e Mahvla Telecomm.\nValor do Aditivo: R$ 845.300,00.\nAssinatura do servidor fiscal: José Carlos de Souza, CPF: 111.222.333-44.",
            "attachments": ["minuta_aditivo_45_2024.pdf"] # we will generate a dummy file
        },
        {
            "task_id": generate_task_id("msg_004_riscodecompliance"),
            "data_ingestao": datetime.now(timezone.utc).isoformat(),
            "remetente_original": "notificacoes@receita.fazenda.gov.br",
            "assunto": "Notificação de Penalidade Administrativa por Atraso - Pregão 89/2025",
            "body": "NOTIFICAÇÃO DE ADVERTÊNCIA E MULTA.\nIdentificamos descumprimento de SLA contratual. Fica aplicada multa administrativa no valor de R$ 5.200.000,00 (cinco milhões e duzentos mil reais) à contratada Mahvla Telecomm. Prazo de defesa administrativa: 5 dias úteis (até 23/06/2026). CPF do servidor responsável pela lavratura: 444.555.666-77. Banco do recolhimento: Caixa Econômica, Agência 1004, Conta 9876-5.",
            "pdf_text": "",
            "attachments": []
        },
        {
            "task_id": generate_task_id("msg_005_esclarecimento"),
            "data_ingestao": datetime.now(timezone.utc).isoformat(),
            "remetente_original": "comissao.licitacao@tjes.jus.br",
            "assunto": "Fwd: Pedido de Esclarecimento Técnico - Pregão nº 12/2026",
            "body": "Prezada engenharia comercial, encaminhamos em anexo a solicitação de esclarecimento técnico enviada pela licitante concorrente sobre as especificações do link dedicado. Favor responder até 20/06/2026 às 18h.",
            "pdf_text": "CONTEÚDO DO PDF:\nPergunta da Licitante concorrente: As fibras fornecidas no certame devem ter dupla abordagem física até a central?",
            "attachments": ["esclarecimento_tecnico_12_2026.pdf"]
        }
    ]

    analyzed_items = []
    
    for idx, email_data in enumerate(mock_emails, start=1):
        print(f"\n[{idx}/5] Teste - Analisando E-mail ID: {email_data['task_id']}")
        print(f"De: {email_data['remetente_original']}")
        print(f"Assunto: {email_data['assunto']}")
        
        # 1. Simulate file attachment creation
        saved_paths = []
        for att in email_data["attachments"]:
            task_dir = os.path.join(test_attachments_dir, email_data["task_id"])
            os.makedirs(task_dir, exist_ok=True)
            dummy_path = os.path.join(task_dir, att)
            with open(dummy_path, "w", encoding="utf-8") as f:
                f.write(f"Conteúdo Dummy de Teste para o anexo {att}.\n" + email_data["pdf_text"])
            saved_paths.append(dummy_path)

        # 2. Run mock Gemini pipeline
        analysis = analyze_email_with_gemini(
            subject=email_data["assunto"],
            sender=email_data["remetente_original"],
            body=email_data["body"],
            pdf_text=email_data["pdf_text"],
            mock=True
        )

        # 3. Local Sanitization for CPFs and bank accounts
        masked_resumo = sanitize_text(analysis.resumo_executivo)
        masked_assunto = sanitize_text(analysis.assunto_padronizado)
        masked_remetente = sanitize_text(analysis.remetente_raiz)

        # Structure links
        link_anexo = ""
        if saved_paths:
            rel_paths = [os.path.relpath(p, start=os.path.dirname(excel_path)) for p in saved_paths]
            link_anexo = ", ".join(rel_paths)

        item_payload = {
            "task_id": email_data["task_id"],
            "data_ingestao": email_data["data_ingestao"],
            "remetente_raiz": masked_remetente,
            "categoria_ia": analysis.categoria_ia,
            "assunto_padronizado": masked_assunto,
            "resumo_executivo": masked_resumo,
            "valor_certame": analysis.valor_certame,
            "prazo_fatal": analysis.prazo_fatal,
            "roteamento_responsavel": analysis.roteamento_responsavel,
            "link_anexo_seguro": link_anexo
        }
        
        print(f"Classificação: {item_payload['categoria_ia']} -> Responsável: {item_payload['roteamento_responsavel']}")
        if item_payload['valor_certame']:
            print(f"Valor Extraído: R$ {item_payload['valor_certame']:,.2f}")
        if item_payload['prazo_fatal']:
            print(f"Prazo Limite: {item_payload['prazo_fatal']}")
            
        analyzed_items.append(item_payload)

    # 3. Update spreadsheet database
    print("\nGravando dados de teste na planilha mestre...")
    create_or_update_excel(excel_path, analyzed_items)
    print(f"\nTeste offline concluído com sucesso!")
    print(f"Verifique o arquivo gerado: {excel_path}")
    print(f"Verifique os anexos gerados na pasta: {test_attachments_dir}/")

if __name__ == "__main__":
    main()
