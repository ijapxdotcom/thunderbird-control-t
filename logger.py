"""
logger.py — Sistema Centralizado de Logging Estruturado
Control-T | Compliance & Routing Engine v3.1

Funcionalidades:
  - RotatingFileHandler: máx. 5 MB por arquivo, 3 backups mantidos
  - Formato estruturado: [DATA HORA] [NÍVEL] [MÓDULO] mensagem
  - Captura global de exceções não tratadas via sys.excepthook
  - Interface simples: get_logger(name) — reutilizável por qualquer módulo
  - Separação de handlers: arquivo (DEBUG+) e console (INFO+ por padrão)
"""

import os
import sys
import logging
from logging.handlers import RotatingFileHandler
from datetime import datetime

# ============================================================
#  CONFIGURAÇÃO GLOBAL
# ============================================================

# Diretório de logs ao lado do projeto
_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_DIR = os.path.join(_BASE_DIR, "logs")
LOG_FILE = os.path.join(LOG_DIR, "control_t.log")

# Limites do arquivo rotativo
MAX_BYTES = 5 * 1024 * 1024   # 5 MB por arquivo
BACKUP_COUNT = 3                # 3 arquivos de backup mantidos

# Formato profissional com módulo de origem
LOG_FORMAT = "[%(asctime)s] [%(levelname)-8s] [%(name)-20s] %(message)s"
DATE_FORMAT = "%Y-%m-%d %H:%M:%S"

# Nível padrão global (pode ser sobrescrito via configure_logging)
DEFAULT_LEVEL = logging.INFO

# Controle para evitar configurar o root logger mais de uma vez
_configured = False


# ============================================================
#  CONFIGURAÇÃO DO SISTEMA DE LOGGING
# ============================================================

def configure_logging(console_level: str = "INFO", file_level: str = "DEBUG") -> None:
    """
    Configura o sistema de logging global do Control-T.

    Deve ser chamado UMA VEZ no início do processo, em compliance_agent.py.
    Configura dois handlers:
      - FileHandler: grava tudo (DEBUG+) em logs/control_t.log com rotação
      - StreamHandler: exibe no console apenas mensagens no nível especificado

    Args:
        console_level: Nível mínimo para exibir no terminal (DEBUG/INFO/WARNING/ERROR)
        file_level: Nível mínimo para gravar no arquivo de log
    """
    global _configured
    if _configured:
        return

    # Garantir diretório de logs
    os.makedirs(LOG_DIR, exist_ok=True)

    # Resolver níveis de string para constantes logging
    numeric_console = _resolve_level(console_level)
    numeric_file = _resolve_level(file_level)
    root_level = min(numeric_console, numeric_file)

    # Root logger
    root = logging.getLogger()
    root.setLevel(root_level)

    # Remover handlers pré-existentes para evitar duplicação
    root.handlers.clear()

    formatter = logging.Formatter(fmt=LOG_FORMAT, datefmt=DATE_FORMAT)

    # ── Handler 1: Arquivo rotativo ──
    try:
        file_handler = RotatingFileHandler(
            filename=LOG_FILE,
            maxBytes=MAX_BYTES,
            backupCount=BACKUP_COUNT,
            encoding="utf-8",
        )
        file_handler.setLevel(numeric_file)
        file_handler.setFormatter(formatter)
        root.addHandler(file_handler)
    except Exception as e:
        # Não travar o sistema se o log em arquivo falhar (ex: permissão)
        print(f"  [WARN] Não foi possível criar handler de arquivo de log: {e}")

    # ── Handler 2: Console ──
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(numeric_console)
    console_handler.setFormatter(_ConsoleFormatter())
    root.addHandler(console_handler)

    # ── Captura global de exceções não tratadas ──
    def _handle_exception(exc_type, exc_value, exc_tb):
        """Captura exceções não capturadas e as grava no log antes de encerrar."""
        if issubclass(exc_type, KeyboardInterrupt):
            # Não logar Ctrl+C como erro
            sys.__excepthook__(exc_type, exc_value, exc_tb)
            return
        crash_logger = logging.getLogger("control_t.crash")
        crash_logger.critical(
            "Exceção não tratada — pipeline encerrado abruptamente.",
            exc_info=(exc_type, exc_value, exc_tb),
        )

    sys.excepthook = _handle_exception

    _configured = True

    # Log de inicialização
    init_logger = logging.getLogger("control_t.logger")
    init_logger.info(
        "Sistema de logging inicializado | arquivo=%s | console=%s | log=%s",
        file_level,
        console_level,
        LOG_FILE,
    )


# ============================================================
#  INTERFACE PÚBLICA
# ============================================================

def get_logger(name: str) -> logging.Logger:
    """
    Retorna um logger nomeado para o módulo especificado.

    Uso em qualquer módulo:
        from logger import get_logger
        log = get_logger(__name__)
        log.info("Mensagem estruturada")
        log.warning("Arquivo não encontrado: %s", path)
        log.error("Falha na API: %s", exc)
        log.debug("Payload enviado: %s", payload[:200])

    Args:
        name: Nome do módulo (use __name__ para automático)

    Returns:
        logging.Logger configurado
    """
    # Se o sistema não foi configurado ainda (ex: uso em testes unitários),
    # garante uma configuração básica para não silenciar mensagens
    if not _configured:
        configure_logging()

    return logging.getLogger(name)


def log_separator(logger: logging.Logger, title: str = "", level: str = "INFO") -> None:
    """
    Grava uma linha separadora no log para marcar início de seção.

    Args:
        logger: Logger a usar
        title: Título da seção (opcional)
        level: Nível do log (INFO por padrão)
    """
    numeric = _resolve_level(level)
    line = f"{'─' * 60}"
    if title:
        line = f"─── {title} {'─' * max(0, 55 - len(title))}"
    logger.log(numeric, line)


def get_log_path() -> str:
    """Retorna o caminho absoluto do arquivo de log atual."""
    return LOG_FILE


# ============================================================
#  FORMATTER CUSTOMIZADO PARA CONSOLE (sem timestamp para não poluir)
# ============================================================

class _ConsoleFormatter(logging.Formatter):
    """
    Formatter limpo para o console. Exibe apenas nível + mensagem,
    sem timestamp (que já aparece no terminal do Windows/bat).
    Para WARNING+ adiciona marcação visual clara.
    """

    _LEVEL_ICONS = {
        logging.DEBUG:    "  [DEBUG]",
        logging.INFO:     "  [INFO] ",
        logging.WARNING:  "  [AVISO]",
        logging.ERROR:    "  [ERRO] ",
        logging.CRITICAL: "  [CRIT] ",
    }

    def format(self, record: logging.LogRecord) -> str:
        icon = self._LEVEL_ICONS.get(record.levelno, "  [LOG]  ")
        message = record.getMessage()

        # Para erros e críticos, adicionar o nome do módulo
        if record.levelno >= logging.WARNING:
            return f"{icon} [{record.name}] {message}"

        return f"{icon} {message}"


# ============================================================
#  UTILITÁRIO INTERNO
# ============================================================

def _resolve_level(level_str: str) -> int:
    """Converte string de nível de log para constante numérica do logging."""
    mapping = {
        "DEBUG": logging.DEBUG,
        "INFO": logging.INFO,
        "WARNING": logging.WARNING,
        "WARN": logging.WARNING,
        "ERROR": logging.ERROR,
        "CRITICAL": logging.CRITICAL,
    }
    return mapping.get(level_str.upper(), logging.INFO)
