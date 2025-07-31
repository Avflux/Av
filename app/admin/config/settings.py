import logging
from .crypto import EnvCrypto

logger = logging.getLogger(__name__)

def load_config():
    """Carrega configurações do arquivo criptografado"""
    try:
        crypto = EnvCrypto()
        return crypto.load_encrypted()
    except Exception as e:
        logger.error(f"Erro ao carregar configurações: {e}")
        return {}

# Carrega configurações
config = load_config()

# Configurações do banco de dados
DB_CONFIG = {
    'host': config.get('MYSQL_HOST', 'localhost'),
    'user': config.get('MYSQL_USER', 'root'),
    'password': config.get('MYSQL_PASSWORD', ''),
    'database': config.get('MYSQL_DATABASE', ''),
    'port': int(config.get('MYSQL_PORT', 3306))
}

# Configurações da aplicação
APP_CONFIG = {
    'title': 'Chronos System',
    'version': '1.0.0',
    'theme': 'dark',
    'geometry': '1024x768',
    'debug': config.get('APP_DEBUG', 'False').lower() == 'true',
}

# Configurações de log
LOG_CONFIG = {
    'log_dir': 'logs',
    'log_level': config.get('LOG_LEVEL', 'INFO'),
    'log_format': '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    'log_file': 'chronos.log'
}