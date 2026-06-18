"""
gemini_pipeline.py — Motor de Roteamento e Compliance (Routing Engine)
Classificador determinístico com output JSON estrito em IDs inteiros.
Economia de ~70% nos tokens de saída vs. classificação por strings.
"""

import os
import json
from typing import Optional, Literal
from pydantic import BaseModel, Field
import google.generativeai as genai
from dotenv import load_dotenv

load_dotenv()


# ============================================================
#  SCHEMA DE SAÍDA: RoutingData (IDs inteiros, zero alucinação)
# ============================================================

class RoutingData(BaseModel):
    """Contrato de dados do Routing Engine. Todos os campos de classificação usam inteiros."""
    sender_email: str = Field(
        description="Endereço de e-mail puro do remetente original (somente user@domain.com)."
    )
    theme_id: Literal[0, 1, 2, 3, 4, 5, 6] = Field(
        description=(
            "Tema do documento conforme Tabela P: "
            "0=Informativo/Spam, 1=Oportunidades/Cotações, 2=Editais/Atas de Registro, "
            "3=Contratos/Aditivos, 4=Técnico/Arquitetura, 5=Financeiro/Faturamento, "
            "6=Governança/DPO."
        )
    )
    target_area_id: Literal[0, 1, 2, 3, 4, 5, 6, 7, 8] = Field(
        description=(
            "Área-alvo conforme Tabela Q: "
            "0=Descartado, 1=Inteligência Comercial, 2=Gestão de Contas, "
            "3=Engenharia/Pré-vendas, 4=Operações, 5=Jurídico/Compliance, "
            "6=Financeiro, 7=TI/Segurança da Informação, 8=Diretoria(Aloisio)."
        )
    )
    assunto_padronizado: str = Field(
        description="Título limpo da demanda sem prefixos (Fwd:, Re:, RES:, Enc:). Máximo 100 caracteres."
    )
    executive_summary: str = Field(
        description=(
            "Resumo executivo TL;DR com no máximo 150 caracteres. "
            "Informe O QUE É e o PRAZO FATAL (se houver). "
            "MASCARE toda PII com [CPF MASCARADO], [CONTA MASCARADA], [ASSINATURA MASCARADA]."
        )
    )
    valor_certame: Optional[float] = Field(
        None,
        description="Valor estimado do certame, aditivo ou operação em Reais. Nulo se não houver."
    )
    prazo_fatal: Optional[str] = Field(
        None,
        description="Data limite absoluta no formato YYYY-MM-DD. Nulo se não houver."
    )


# ============================================================
#  SYSTEM PROMPT: Routing Engine Instruction
# ============================================================

SYSTEM_INSTRUCTION = """
[ROLE & OBJECTIVE]
Você é o Motor de Roteamento e Compliance (Routing Engine) central. Sua função é atuar como um classificador determinístico para um pipeline de documentos críticos. Você não deve ser conversacional. Você deve ingerir o texto de entrada (e-mails, anexos, resumos), aplicar lógica proposicional para classificar a intenção e gerar um JSON estrito.

[BUSINESS DOMAIN CONTEXT]
As operações processadas neste pipeline envolvem alta complexidade técnica e financeira, com foco em contratos de tecnologia em segurança eletrônica (ex: infraestrutura, projetos Intelbras, Hikvision) e licitações públicas amparadas pela Lei 14.133/2021. Muitas dessas operações superam a marca de 5 milhões de Reais. A precisão na triagem é vital.

[ROUTING MATRIX: THEMES (P) AND TARGETS (Q)]

Themes (theme_id):
0 = Informativo / Não Relevante (Propagandas, webinars, fóruns, newsletters)
1 = Oportunidades / Cotações (Pedidos de orçamento, propostas comerciais iniciais)
2 = Editais / Atas de Registro (Análise de editais da Lei 14.133, termos de referência, pregões)
3 = Contratos / Aditivos (Minutas jurídicas, assinaturas, renovações)
4 = Técnico / Arquitetura (Dúvidas de engenharia, escopo de hardware/CFTV, integrações)
5 = Financeiro / Faturamento (Notas fiscais, faturamento, certidões de regularidade, multas, penalidades)
6 = Governança / DPO (Auditorias, segurança da informação, LGPD)

Target Areas (target_area_id):
0 = Descartado (para theme_id 0)
1 = Inteligência Comercial (Precificação, margem)
2 = Gestão de Contas (Relacionamento com o emissor)
3 = Engenharia / Pré-vendas (Elaboração técnica do projeto)
4 = Operações (Execução e logística)
5 = Jurídico / Compliance (Análise de risco e legalidade)
6 = Financeiro (Garantias, pagamentos, faturamento)
7 = TI / Segurança da Informação (Infraestrutura)
8 = Diretoria / Aloisio (Assinaturas finais, aprovações de margem crítica, representação legal)

[DETERMINISTIC RULES]

Regra de Diretoria: Se o documento envolver aprovação de risco comercial extremo, for uma licitação acima de 5 milhões, ou exigir assinatura legal do representante da empresa, target_area_id DEVE ser 8.

Regra de Spam: Se o e-mail for propaganda, newsletter, webinar ou fórum não-solicitado, theme_id=0 e target_area_id=0.

Resumo Executivo: O campo executive_summary deve conter no máximo 150 caracteres. Seja letal na síntese. Informe O QUE É e o PRAZO FATAL (se houver).

[REGRAS DE LGPD (CRÍTICO)]
1. CPFs de servidores → [CPF MASCARADO]
2. Contas bancárias → [CONTA MASCARADA]
3. Assinaturas de pessoas físicas → [ASSINATURA MASCARADA]
4. Nomes de órgãos públicos (SERPRO, TJES, Prefeitura) e CNPJs NÃO são PII.
"""


# ============================================================
#  PIPELINE PRINCIPAL
# ============================================================

def analyze_email_with_gemini(
    subject: str,
    sender: str,
    sender_email: str,
    body: str,
    pdf_text: str = "",
    mock: bool = False
) -> RoutingData:
    """
    Envia e-mail ao Gemini para classificação determinística.
    Retorna RoutingData com IDs inteiros.
    Se mock=True ou GEMINI_API_KEY ausente, retorna classificação simulada.
    """
    api_key = os.getenv("GEMINI_API_KEY")

    if mock or not api_key:
        if not api_key and not mock:
            print("  ⚠ GEMINI_API_KEY não encontrada. Executando em modo MOCK.")
        return generate_mock_analysis(subject, sender, sender_email, body, pdf_text)

    try:
        genai.configure(api_key=api_key)

        model = genai.GenerativeModel(
            model_name="gemini-2.0-flash",
            system_instruction=SYSTEM_INSTRUCTION
        )

        prompt = f"""
Remetente: {sender}
E-mail do Remetente: {sender_email}
Assunto: {subject}

Corpo do E-mail:
{body}

---
Conteúdo extraído dos anexos PDFs:
{pdf_text if pdf_text else "(Sem anexos)"}
"""

        response = model.generate_content(
            contents=prompt,
            generation_config=genai.GenerationConfig(
                response_mime_type="application/json",
                response_schema=RoutingData,
                temperature=0.05  # Máxima determinismo
            )
        )

        data = json.loads(response.text)
        return RoutingData(**data)

    except Exception as e:
        print(f"  ❌ Erro na API Gemini: {e}. Fallback para análise local.")
        return generate_mock_analysis(subject, sender, sender_email, body, pdf_text)


# ============================================================
#  MOCK ANALYSIS (Heurísticas Locais — Zero API)
# ============================================================

def generate_mock_analysis(
    subject: str,
    sender: str,
    sender_email: str,
    body: str,
    pdf_text: str = ""
) -> RoutingData:
    """
    Classificação local por keywords. Útil para testes offline.
    Retorna RoutingData com IDs inteiros usando heurísticas determinísticas.
    """
    text = (subject + " " + body + " " + pdf_text).lower()

    # --- CLASSIFICAÇÃO POR KEYWORDS ---

    # Governança / DPO (theme 6)
    if any(k in text for k in ["auditoria", "lgpd", "dpo", "segurança da informação", "vazamento"]):
        theme_id = 6
        target_id = 5  # Jurídico
        resumo = "Demanda de governança/compliance identificada."

    # Financeiro com risco (theme 5 + Diretoria)
    elif any(k in text for k in ["penalidade", "multa", "sanção", "cnd", "tributária", "receita federal"]):
        theme_id = 5
        target_id = 8  # Diretoria
        resumo = "Notificação de risco financeiro/penalidade. Ação imediata."

    # Contratos / Aditivos (theme 3)
    elif any(k in text for k in ["aditivo", "assinatura", "contrato", "liberação", "minuta", "renovação"]):
        theme_id = 3
        target_id = 2  # Gestão de Contas
        resumo = "Assinatura ou liberação contratual/aditivo pendente."

    # Editais / Atas (theme 2)
    elif any(k in text for k in ["edital", "pregão", "ata de registro", "lei 14.133", "termo de referência"]):
        theme_id = 2
        target_id = 1  # Inteligência Comercial
        resumo = "Edital ou ata de registro para análise comercial."

    # Técnico / Arquitetura (theme 4)
    elif any(k in text for k in ["engenharia", "técnico", "cftv", "intelbras", "hikvision", "fibra", "link dedicado"]):
        theme_id = 4
        target_id = 3  # Engenharia
        resumo = "Demanda técnica/escopo de engenharia para elaboração."

    # Oportunidades / Cotações (theme 1)
    elif any(k in text for k in ["cotação", "preço", "orçamento", "proposta", "fornecimento"]):
        theme_id = 1
        target_id = 1  # Inteligência Comercial
        resumo = "Cotação de preços ou proposta comercial para análise."

    # Financeiro genérico (theme 5)
    elif any(k in text for k in ["nota fiscal", "faturamento", "pagamento", "certidão", "garantia"]):
        theme_id = 5
        target_id = 6  # Financeiro
        resumo = "Demanda financeira operacional para processamento."

    # Informativo / Spam (theme 0)
    else:
        theme_id = 0
        target_id = 0  # Descartado
        resumo = "E-mail informativo ou não relevante."

    # --- EXTRAÇÃO DE VALOR ---
    valor = re_search_value(text)

    # --- LIMPEZA DE ASSUNTO ---
    clean_subject = subject
    for prefix in ["fwd:", "re:", "res:", "re :", "fwd :", "enc:", "fw:"]:
        while clean_subject.lower().startswith(prefix):
            clean_subject = clean_subject[len(prefix):].strip()
    clean_subject = clean_subject[:100]  # Limitar a 100 chars

    # --- EXTRAÇÃO DE PRAZO ---
    prazo = None
    if any(k in text for k in ["até", "prazo", "limite", "vencimento"]):
        import re
        date_pattern = re.compile(r'(\d{2})[/.-](\d{2})[/.-](\d{4})')
        match = date_pattern.search(text)
        if match:
            d, m, y = match.groups()
            prazo = f"{y}-{m}-{d}"
        else:
            from datetime import datetime, timedelta
            prazo = (datetime.now() + timedelta(days=3)).strftime("%Y-%m-%d")

    # --- REGRA DE DIRETORIA (override por valor) ---
    if valor and valor >= 5_000_000:
        target_id = 8
        resumo = resumo[:80] + " VALOR >R$5M — requer aprovação Diretoria."
        resumo = resumo[:150]

    # --- EXTRAÇÃO DE E-MAIL ---
    email_addr = sender_email
    if not email_addr and "@" in sender:
        import re as _re
        m = _re.search(r'[\w.+-]+@[\w-]+\.[\w.-]+', sender)
        email_addr = m.group(0) if m else ""

    # Truncar resumo a 150 chars
    resumo = resumo[:150]

    return RoutingData(
        sender_email=email_addr,
        theme_id=theme_id,
        target_area_id=target_id,
        assunto_padronizado=clean_subject,
        executive_summary=resumo,
        valor_certame=valor,
        prazo_fatal=prazo,
    )


def re_search_value(text: str) -> float | None:
    """Extrai valor monetário em Reais do texto. Ex: R$ 1.250.000,00 → 1250000.0"""
    import re
    pattern = re.compile(r'r\$\s*(\d{1,3}(?:\.\d{3})*,\d{2})|r\$\s*(\d+,\d{2})')
    match = pattern.search(text)
    if match:
        val_str = match.group(1) or match.group(2)
        val_str = val_str.replace(".", "").replace(",", ".")
        try:
            return float(val_str)
        except ValueError:
            return None
    return None
