import mysql.connector
from mysql.connector import Error
import logging
import os
import sys

logger = logging.getLogger(__name__)

class DatabaseOperations:
    @staticmethod
    def test_connection(config):
        """Testa a conexão com o banco de dados"""
        try:
            connection = mysql.connector.connect(**config)
            if connection.is_connected():
                return True, None
        except Error as e:
            return False, str(e)
        finally:
            if 'connection' in locals() and connection.is_connected():
                connection.close()
                
    @staticmethod
    def create_database(config, db_name):
        """Cria um novo banco de dados e suas tabelas"""
        try:
            # Remove database from config
            db_config = config.copy()
            db_config.pop('database', None)
            
            connection = mysql.connector.connect(**db_config)
            cursor = connection.cursor()
            
            # Verify if database exists
            cursor.execute("SHOW DATABASES")
            if db_name.lower() in [db[0].lower() for db in cursor.fetchall()]:
                return False, "Banco de dados já existe"
            
            # Create database
            cursor.execute(f"CREATE DATABASE {db_name}")
            connection.database = db_name
            
            # Base path for bundled app (PyInstaller)
            if hasattr(sys, "_MEIPASS"):
                base_path = sys._MEIPASS
            else:
                base_path = os.path.abspath(".")

            # Procura o arquivo SQL
            sql_file_path = os.path.join(base_path, "sql", "create_tables.sql")
            
            if not os.path.exists(sql_file_path):
                raise FileNotFoundError(f"Arquivo SQL não encontrado: {sql_file_path}")
            
            # Create tables
            with open(sql_file_path, 'r', encoding='utf-8') as file:
                sql_script = file.read()
                
            # Divide o script em comandos individuais, preservando os delimitadores
            current_delimiter = ';'
            commands = []
            current_command = ''
            
            for line in sql_script.splitlines():
                line = line.strip()
                if not line or line.startswith('--'):  # Ignora linhas vazias e comentários
                    continue
                    
                if line.upper().startswith('DELIMITER'):
                    if current_command:
                        commands.append((current_command, current_delimiter))
                        current_command = ''
                    current_delimiter = line.split()[1]
                    continue
                
                current_command += ' ' + line
                if line.endswith(current_delimiter):
                    current_command = current_command.strip()
                    if current_delimiter != ';':
                        # Remove o delimitador personalizado do final
                        current_command = current_command[:-len(current_delimiter)]
                    commands.append((current_command, current_delimiter))
                    current_command = ''
            
            # Executa cada comando
            for cmd, delimiter in commands:
                if cmd.strip():
                    try:
                        cursor.execute(cmd)
                        # Consome qualquer resultado pendente
                        try:
                            while cursor.nextset():
                                pass
                        except:
                            pass
                    except Error as e:
                        logger.error(f"Erro ao executar comando SQL: {e}")
                        logger.error(f"Comando que falhou: {cmd}")
                        raise
                
            connection.commit()
            return True, None
            
        except FileNotFoundError as e:
            logger.error(f"Erro ao localizar arquivo SQL: {e}")
            return False, f"Erro ao localizar arquivo SQL: {e}"
        except Error as e:
            logger.error(f"Erro MySQL: {e}")
            return False, str(e)
        finally:
            if 'cursor' in locals():
                cursor.close()
            if 'connection' in locals():
                connection.close()
    
    @staticmethod
    def change_mysql_password(config, current_password, new_password):
        """Altera a senha do usuário MySQL"""
        try:
            connection = mysql.connector.connect(
                **config,
                password=current_password
            )
            cursor = connection.cursor()
            
            cursor.execute(
                f"ALTER USER '{config['user']}'@'localhost' "
                f"IDENTIFIED BY '{new_password}'"
            )
            connection.commit()
            return True, None
            
        except Error as e:
            return False, str(e)
        finally:
            if 'cursor' in locals():
                cursor.close()
            if 'connection' in locals():
                connection.close()