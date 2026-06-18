# Compliance Agent & Orquestrador de E-mails de Licitação

Este projeto automatiza a leitura, classificação e monitoramento de e-mails de editais de licitação a partir das caixas locais do Thunderbird (Zoho e Outlook). Ele utiliza a API do Gemini para categorizar as demandas por níveis de criticidade e aplicar mascaramento de dados (LGPD), salvando anexos sensíveis em uma pasta segura e gerando um painel interativo no Microsoft Excel.

## 🚀 Como Começar

### 1. Requisitos
- Windows OS
- Python 3.10 ou superior (instalado no sistema)
- Thunderbird com as contas Zoho e Outlook sincronizadas localmente (em formato MBOX)

### 2. Instalação
Inicialize o ambiente virtual e instale as dependências:
```bash
# Criar ambiente virtual (caso não exista)
py -m venv .venv

# Instalar dependências
.venv\Scripts\pip install -r requirements.txt
```

### 3. Configuração (.env)
Copie o arquivo `.env.example` para `.env` e configure suas variáveis:
```env
GEMINI_API_KEY=sua_chave_do_gemini_aqui
EXCEL_PATH=C:\Users\igor.oliveira\Grupo Mahvla\Mahvla Comercial - Documentos\Mahvla Comercial\Demandas-Control-t\monitoramento_licitacoes.xlsx
SECURE_ATTACHMENTS_DIR=C:\Users\igor.oliveira\Grupo Mahvla\Mahvla Comercial - Documentos\Mahvla Comercial\Demandas-Control-t\secure_attachments
```

### 4. Executando o Script

#### Modo de Teste Offline (Dados Sintéticos)
Para testar a geração do arquivo Excel e o mascaramento sem usar a API paga do Gemini e sem ler caixas de e-mails reais:
```bash
.venv\Scripts\python compliance_agent.py --test-mode
```

#### Modo Dry-Run (Ler E-mails Reais com Heurística Local)
Para testar a leitura de e-mails do seu Thunderbird local sem consumir tokens do Gemini:
```bash
.venv\Scripts\python compliance_agent.py --mock-gemini --hours 36
```

#### Modo de Produção (Gemini Real)
Para rodar a orquestração completa em tempo real utilizando a inteligência artificial do Gemini:
```bash
.venv\Scripts\python compliance_agent.py
```

---

## 🏗️ Estrutura do Código

- `compliance_agent.py`: Orquestrador e loop de execução principal.
- `mbox_parser.py`: Leitor incremental de arquivos MBOX do Thunderbird e extrator de textos de PDFs.
- `pii_masker.py`: Algoritmo local de higienização de CPFs, dados bancários e assinaturas (Regras de LGPD).
- `gemini_pipeline.py`: Conexão com o Gemini, definição dos prompts e saída JSON estruturada.
- `excel_handler.py`: Escrita e formatação avançada do Excel (`RAW - Ingestao`, `VIEW - Fila de Operacao` e `DASH - Painel Executivo`).
