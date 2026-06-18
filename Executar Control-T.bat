@echo off
chcp 65001 >nul 2>&1
title Control-T | Compliance Engine v3.0
color 0A

echo.
echo   ╔══════════════════════════════════════════════════════╗
echo   ║                                                      ║
echo   ║   CONTROL-T  ^|  Compliance ^& Routing Engine v3.0    ║
echo   ║   Gestao Inteligente de Licitacoes                   ║
echo   ║                                                      ║
echo   ╠══════════════════════════════════════════════════════╣
echo   ║                                                      ║
echo   ║   Processando e-mails Zoho...                        ║
echo   ║   Classificando com Routing Engine...                ║
echo   ║   Gerando planilha executiva...                      ║
echo   ║                                                      ║
echo   ╚══════════════════════════════════════════════════════╝
echo.

cd /d "c:\PIBIC\thunderbird-control-t"

REM Ativar ambiente virtual
call .venv\Scripts\activate.bat

REM Executar o Compliance Agent
python compliance_agent.py

echo.
echo   ════════════════════════════════════════════════════════
echo.
echo   Pressione qualquer tecla para fechar...
pause >nul
