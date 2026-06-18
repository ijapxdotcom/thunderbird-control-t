import re

def mask_cpf(text):
    """
    Masks formatted CPFs (e.g., 123.456.789-00) and attempts to mask unformatted 11-digit CPFs
    when preceded by terms like 'CPF'.
    """
    if not text:
        return ""
    
    # 1. Formatted CPF: 000.000.000-00
    formatted_cpf_pattern = re.compile(r'\b\d{3}\.\d{3}\.\d{3}-\d{2}\b')
    text = formatted_cpf_pattern.sub('[CPF MASCARADO]', text)
    
    # 2. Unformatted CPF: preceded by 'CPF' label (case-insensitive)
    unformatted_cpf_pattern = re.compile(r'(?i)\b(cpf\s*[:.-]?\s*)(\d{11})\b')
    text = unformatted_cpf_pattern.sub(r'\1[CPF MASCARADO]', text)
    
    return text

def mask_bank_details(text):
    """
    Masks bank account details (agency, account, bank codes) using common Brazilian patterns.
    """
    if not text:
        return ""
    
    # Matches patterns like "Ag: 1234", "C/C: 12345-6", "Agência 1234-5 Conta 123456-7"
    # Case-insensitive matches for agency and account terms
    agency_pattern = re.compile(r'(?i)\b(ag(?:[êe]ncia)?\s*[:.-]?\s*)(\d{3,5}(?:-\d)?)', re.UNICODE)
    account_pattern = re.compile(r'(?i)\b(c[c/]c|conta\s*(?:corrente)?\s*[:.-]?\s*)(\d{4,12}(?:-\d)?)', re.UNICODE)
    
    text = agency_pattern.sub(r'\1[AGENCIA MASCARADA]', text)
    text = account_pattern.sub(r'\1[CONTA MASCARADA]', text)
    
    return text

def mask_digital_signatures(text):
    """
    Masks common signature tags and hashes.
    """
    if not text:
        return ""
    
    # Matches digital signature statements e.g., "Assinado digitalmente por: Nome"
    sig_pattern = re.compile(r'(?i)(assinado\s+digitalmente\s+(?:por\s*)?:?)\s*([^\n,]{3,50})')
    text = sig_pattern.sub(r'\1 [NOME MASCARADO LGPD]', text)
    
    # Matches security hashes or keys often attached to emails
    hash_pattern = re.compile(r'\b[a-f0-9]{32,64}\b')
    # Let's not mask task_id itself (which is 16 chars hex), but longer 32/64 char hashes
    text = hash_pattern.sub('[HASH DE SEGURANÇA MASCARADO]', text)
    
    return text

def sanitize_text(text):
    """
    Main entry point for local text PII sanitization.
    """
    if not text:
        return ""
    text = mask_cpf(text)
    text = mask_bank_details(text)
    text = mask_digital_signatures(text)
    return text
