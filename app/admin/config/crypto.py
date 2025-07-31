import base64
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
import os

class EnvCrypto:
    def __init__(self, key_file='crypto.key'):
        self.key_file = key_file
        self.key = self._load_or_generate_key()
        self.cipher_suite = Fernet(self.key)
    
    def _load_or_generate_key(self):
        """Carrega ou gera uma nova chave de criptografia"""
        if os.path.exists(self.key_file):
            with open(self.key_file, 'rb') as key_file:
                return key_file.read()
        else:
            salt = os.urandom(16)
            kdf = PBKDF2HMAC(
                algorithm=hashes.SHA256(),
                length=32,
                salt=salt,
                iterations=100000,
            )
            key = base64.urlsafe_b64encode(kdf.derive(os.urandom(32)))
            
            with open(self.key_file, 'wb') as key_file:
                key_file.write(key)
            return key
    
    def encrypt_content(self, content: str) -> bytes:
        """Criptografa uma string de conteúdo"""
        return self.cipher_suite.encrypt(content.encode())
    
    def decrypt_content(self, encrypted_content: bytes) -> str:
        """Descriptografa conteúdo para string"""
        return self.cipher_suite.decrypt(encrypted_content).decode()
    
    def save_encrypted(self, config_dict: dict, output_path='.env.encrypted'):
        """Salva configurações como arquivo criptografado"""
        content = '\n'.join(f"{k}={v}" for k, v in config_dict.items())
        encrypted = self.encrypt_content(content)
        
        with open(output_path, 'wb') as f:
            f.write(encrypted)
    
    def load_encrypted(self, input_path='.env.encrypted') -> dict:
        """Carrega e descriptografa configurações para um dicionário"""
        try:
            with open(input_path, 'rb') as f:
                encrypted = f.read()
            
            decrypted = self.decrypt_content(encrypted)
            config = {}
            
            for line in decrypted.splitlines():
                if '=' in line:
                    key, value = line.split('=', 1)
                    config[key.strip()] = value.strip()
            
            return config
        except FileNotFoundError:
            return {}
        except Exception as e:
            print(f"Erro ao carregar configurações: {e}")
            return {}