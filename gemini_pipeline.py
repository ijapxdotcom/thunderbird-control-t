import os
import json
from typing import Optional, Literal
from pydantic import BaseModel, Field
import google.generativeai as genai
from dotenv import load_dotenv

load_dotenv()

# Define the Pydantic schema for structured output validation
class BiddingAnalysis(BaseModel):
    remetente_raiz: str = Field(description="O órgão público ou cliente original emissor do edital/e-mail (ex: SERPRO, TJES, CREA-SP). Ignore intermediários ou servidores individuais.")
    categoria_ia: Literal["Informativo", "Operacional", "Atenção Alta", "Risco Crítico"] = Field(description="Classificação da matriz de criticidade baseada no teor do conteúdo.")
    assunto_padronizado: str = Field(description="Título limpo da demanda sem prefixos como Fwd:, Re:, RES:, etc.")
    resumo_executivo: str = Field(description="Resumo executivo TL;DR focando em exigências, valores e riscos. MASCARE toda PII (como CPFs de servidores, dados bancários e assinaturas) com placeholders tipo [CPF MASCARADO] ou [CONTA MASCARADA].")
    valor_certame: Optional[float] = Field(None, description="Valor estimado extraído do edital, aditivo ou e-mail. Usar nulo se não houver.")
    prazo_fatal: Optional[str] = Field(None, description="Data limite absoluta para resposta, impugnação ou envio de propostas no formato YYYY-MM-DD. Usar nulo se não houver.")
    roteamento_responsavel: Literal["Diretoria", "Gestão de Contas", "Comercial", "Engenharia", "Descartado"] = Field(description="Setor responsável: Diretoria (para Risco Crítico), Gestão de Contas (para Nível 3), Comercial (para Cotações/Pesquisa de Preços), Engenharia (para Esclarecimentos de edital técnicos), Descartado (para Nível 1).")

SYSTEM_INSTRUCTION = """
Você é um Agente de Compliance Sênior e Engenheiro de Dados.
Sua função é ler e-mails e editais de licitação pública brasileira e extrair dados críticos.
Você deve atuar estritamente como um filtro de mascaramento (Data Masking) em conformidade com a LGPD.

REGRAS DE LGPD (CRÍTICO):
1. Qualquer Informação Pessoalmente Identificável (PII) nos campos gerados, como CPFs de servidores, assinaturas digitais, dados bancários de colaboradores ou contatos pessoais, deve ser substituída por placeholders:
   - CPF -> [CPF MASCARADO]
   - Contas bancárias -> [CONTA MASCARADA] ou [DADOS BANCÁRIOS MASCARADOS]
   - Assinaturas de pessoas físicas -> [ASSINATURA MASCARADA]
2. Nomes de órgãos públicos (ex: SERPRO, TJES, Prefeitura de SP) e CNPJs de órgãos públicos NÃO são PII e devem ser mantidos intactos no remetente_raiz.

REGRAS DE CLASSIFICAÇÃO:
- "Informativo" (Nível 1): Propagandas, webinars, fóruns, newsletters. Roteamento: "Descartado".
- "Operacional" (Nível 2): Cotações, esclarecimentos de edital, pesquisa de preço de rotina. Roteamento: "Comercial" ou "Engenharia".
- "Atenção Alta" (Nível 3): Assinaturas externas de contratos, liberação de aditivos. Roteamento: "Gestão de Contas".
- "Risco Crítico" (Nível 4): Notificações de penalidade, multas, pendências tributárias (CND), impasses internos graves. Roteamento: "Diretoria".
"""

def analyze_email_with_gemini(subject: str, sender: str, body: str, pdf_text: str = "", mock: bool = False) -> BiddingAnalysis:
    """
    Sends email subject, sender, body and pdf attachments text to Gemini API.
    Enforces structured JSON output matching the BiddingAnalysis schema.
    If mock is True or GEMINI_API_KEY is missing, returns a simulated classification.
    """
    api_key = os.getenv("GEMINI_API_KEY")
    
    if mock or not api_key:
        if not api_key and not mock:
            print("Warning: GEMINI_API_KEY not found in .env. Running in MOCK mode.")
        return generate_mock_analysis(subject, sender, body, pdf_text)

    try:
        genai.configure(api_key=api_key)
        
        # We use gemini-1.5-flash as the default lightweight, fast, and structured-output-capable model
        model = genai.GenerativeModel(
            model_name="gemini-1.5-flash",
            system_instruction=SYSTEM_INSTRUCTION
        )
        
        prompt = f"""
Remetente do E-mail: {sender}
Assunto do E-mail: {subject}

Corpo do E-mail:
{body}

---
Conteúdo extraído dos anexos PDFs:
{pdf_text}
"""
        
        response = model.generate_content(
            contents=prompt,
            generation_config=genai.GenerationConfig(
                response_mime_type="application/json",
                response_schema=BiddingAnalysis,
                temperature=0.1
            )
        )
        
        # Parse the JSON response
        data = json.loads(response.text)
        return BiddingAnalysis(**data)
        
    except Exception as e:
        print(f"Error calling Gemini API: {e}. Falling back to local mock analysis.")
        return generate_mock_analysis(subject, sender, body, pdf_text)

def generate_mock_analysis(subject: str, sender: str, body: str, pdf_text: str = "") -> BiddingAnalysis:
    """
    Simulates LLM classification based on simple keywords.
    Useful for offline testing or when the API key is not configured.
    """
    text = (subject + " " + body + " " + pdf_text).lower()
    
    # Simple heuristics to classify
    if any(k in text for k in ["penalidade", "multa", "cnd", "tributária", "receita federal", "sancionatório"]):
        categoria = "Risco Crítico"
        roteamento = "Diretoria"
        resumo = "Notificação de risco de compliance identificada. Requer ação imediata."
    elif any(k in text for k in ["aditivo", "assinatura", "contrato", "liberação"]):
        categoria = "Atenção Alta"
        roteamento = "Gestão de Contas"
        resumo = "Assinatura ou liberação de termo aditivo contratual pendente."
    elif any(k in text for k in ["cotação", "esclarecimento", "edital", "preço"]):
        categoria = "Operacional"
        # Determine routing
        if "engenharia" in text or "técnico" in text:
            roteamento = "Engenharia"
        else:
            roteamento = "Comercial"
        resumo = "Pesquisa de preço ou pedido de esclarecimento operacional de edital."
    else:
        categoria = "Informativo"
        roteamento = "Descartado"
        resumo = "E-mail informativo ou comercial de rotina (propagandas/webinars)."

    # Extract value if present
    valor = None
    value_match = re_search_value(text)
    if value_match:
        valor = value_match
        
    # Standardize subject
    clean_subject = subject
    for prefix in ["fwd:", "re:", "res:", "re :", "fwd :"]:
        if clean_subject.lower().startswith(prefix):
            clean_subject = clean_subject[len(prefix):].strip()

    # Extract original sender
    remetente = "Órgão Não Identificado"
    if "@" in sender:
        domain = sender.split("@")[-1].replace(">", "").strip()
        remetente = domain.split(".")[0].upper()
    
    # Estimate deadline
    prazo = None
    if "prazo" in text or "limite" in text:
        # 3 days from now
        from datetime import datetime, timedelta
        prazo = (datetime.now() + timedelta(days=3)).strftime("%Y-%m-%d")

    return BiddingAnalysis(
        remetente_raiz=remetente,
        categoria_ia=categoria,
        assunto_padronizado=clean_subject,
        resumo_executivo=resumo,
        valor_certame=valor,
        prazo_fatal=prazo,
        roteamento_responsavel=roteamento
    )

def re_search_value(text: str) -> Optional[float]:
    import re
    # Matches patterns like R$ 1.500.000,00 or R$1500000.00
    pattern = re.compile(r'r\$\s*(\d{1,3}(?:\.\d{3})*,\d{2})|r\$\s*(\d+,\d{2})')
    match = pattern.search(text)
    if match:
        val_str = match.group(1) or match.group(2)
        # Convert to float (replace dot with empty, replace comma with dot)
        val_str = val_str.replace(".", "").replace(",", ".")
        try:
            return float(val_str)
        except ValueError:
            return None
    return None
