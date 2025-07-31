import customtkinter as ctk
import logging, os, sys
from tkinter import messagebox
from app.admin.config.crypto import EnvCrypto
from app.admin.database.operations import DatabaseOperations
from .dialogs.change_password import ChangePasswordDialog
from PIL import Image

logger = logging.getLogger(__name__)

class DatabaseConfigDialog(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Sistema Chronos - Administrativo")
                
        self.setup_window()
        self.load_current_config()
        self.setup_ui()
    
    def setup_window(self):
        """Configura a janela"""
        self.geometry("400x650")
        self.resizable(False, False)
        self.center_window()
    
    def center_window(self):
        """Centraliza a janela"""
        x = (self.winfo_screenwidth() - 400) // 2
        y = (self.winfo_screenheight() - 650) // 2
        self.geometry(f"450x650+{x}+{y}")
    
    def load_current_config(self):
        """Carrega configuração atual"""
        crypto = EnvCrypto()
        config = crypto.load_encrypted()
        self.current_config = {
            'host': config.get('MYSQL_HOST', 'localhost'),
            'port': config.get('MYSQL_PORT', '3306'),
            'user': config.get('MYSQL_USER', 'root'),
            'password': config.get('MYSQL_PASSWORD', ''),
            'database': config.get('MYSQL_DATABASE', '')
        }
    
    def setup_ui(self):
        """Setup dialog interface"""
        # Base path for bundled app (PyInstaller)
        if hasattr(sys, "_MEIPASS"):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.abspath(".")

        # Load and display logo images
        light_image_path = os.path.join(base_path, "app", "admin", "icon", "logo_light.png")
        dark_image_path = os.path.join(base_path, "app", "admin", "icon", "logo_dark.png")

        if not os.path.exists(light_image_path) or not os.path.exists(dark_image_path):
            raise FileNotFoundError(f"Image files not found: {light_image_path}, {dark_image_path}")

        logo_image = ctk.CTkImage(
            light_image=Image.open(light_image_path),
            dark_image=Image.open(dark_image_path),
            size=(223, 50)
        )

        logo_label = ctk.CTkLabel(
            self,
            image=logo_image,
            text=""  # Empty text as we only want to display the image
        )
        logo_label.pack(pady=(20, 10))

        # Title
        self.title_label = ctk.CTkLabel(
            self,
            text="Configurações do MySQL",
            font=("Roboto", 20, "bold")
        )
        self.title_label.pack(pady=(10, 15))
        
        # Form frame
        self.form_frame = ctk.CTkFrame(self)
        self.form_frame.pack(padx=20, fill="x")
        
        # Configuration fields
        self.entries = {}
        fields = [
            ('host', 'Host:', 'Digite o host'),
            ('port', 'Porta:', 'Digite a porta'),
            ('user', 'Usuário:', 'Digite o usuário'),
            ('password', 'Senha:', 'Digite a senha'),
            ('database', 'Banco de Dados:', 'Nome do banco de dados')
        ]
        
        for key, label_text, placeholder in fields:
            field_frame = ctk.CTkFrame(self.form_frame, fg_color="transparent")
            field_frame.pack(fill="x", pady=10)
            
            label = ctk.CTkLabel(
                field_frame,
                text=label_text,
                font=("Roboto", 14)
            )
            label.pack(anchor="w")
            
            entry = ctk.CTkEntry(
                field_frame,
                width=340,
                height=35,
                placeholder_text=placeholder
            )
            entry.pack(fill="x", pady=(5, 0))
            
            # Inserir o valor inicial ou deixar o placeholder visível
            initial_value = self.current_config.get(key, '')
            if initial_value:
                entry.insert(0, initial_value)
            
            if key == 'password':
                entry.configure(show="*")
            
            self.entries[key] = entry
        
        # Buttons frame
        self.button_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.button_frame.pack(fill="x", padx=20, pady=(20, 10))
        
        # Test connection button
        self.test_btn = ctk.CTkButton(
            self.button_frame,
            text="Testar Conexão",
            fg_color="#FF5722",
            hover_color="#CE461B",
            command=self.test_connection,
            width=120,
            font=("Roboto", 14)
        )
        self.test_btn.pack(side="left", padx=5)
        
        # Change MySQL password button
        self.change_pass_btn = ctk.CTkButton(
            self.button_frame,
            text="Alterar Senha MySQL",
            fg_color="#FF5722",
            hover_color="#CE461B",
            command=self.show_change_password_dialog,
            width=120,
            font=("Roboto", 14)
        )
        self.change_pass_btn.pack(side="left", padx=5)
        
        # Create database button
        self.create_db_btn = ctk.CTkButton(
            self.button_frame,
            text="Criar Banco",
            fg_color="#FF5722",
            hover_color="#CE461B",
            command=self.create_database,
            width=120,
            font=("Roboto", 14)
        )
        self.create_db_btn.pack(side="right", padx=5)

    def show_change_password_dialog(self):
        """Show dialog to change MySQL password"""
        dialog = ChangePasswordDialog(self)
        self.wait_window(dialog)
        
        # Atualizar o campo de senha se a senha foi alterada com sucesso
        if hasattr(dialog, 'changed_password') and dialog.changed_password is not None:
            self.entries['password'].delete(0, 'end')
            self.entries['password'].insert(0, dialog.changed_password)
            self.update_config()

    def update_config(self):
        """Update encrypted configuration file"""
        try:
            config = {
                f"MYSQL_{key.upper()}": entry.get().strip() 
                for key, entry in self.entries.items()
            }
            
            crypto = EnvCrypto()
            existing_config = crypto.load_encrypted()
            existing_config.update(config)
            
            crypto.save_encrypted(existing_config)
            logger.info("Database configuration updated and encrypted")
            
        except Exception as e:
            logger.error(f"Erro ao atualizar configurações: {e}")
            messagebox.showerror("Erro", f"Erro ao atualizar configurações: {e}")

    def test_connection(self):
        """Test MySQL connection using current settings"""
        config = {key: entry.get().strip() 
                for key, entry in self.entries.items()}
        
        if not all([config['host'], config['port'], config['user']]):
            messagebox.showerror("Erro", "Preencha os campos obrigatórios!")
            return

        success, error = DatabaseOperations.test_connection(config)
        if success:
            if not config['database']:
                messagebox.showinfo("Sucesso", 
                    "Conexão estabelecida com sucesso! Informe o nome do banco de dados para criação dos arquivos de chave para o Sistema Chronos")
            else:
                messagebox.showinfo("Sucesso", "Conexão estabelecida com sucesso!")
                self.create_crypto_files(config)
        else:
            messagebox.showerror("Erro", f"Erro na conexão: {error}")
    
    def create_crypto_files(self, config):
        """Create encrypted configuration files"""
        try:
            crypto = EnvCrypto()
            encrypted_config = {
                'MYSQL_HOST': config['host'],
                'MYSQL_PORT': config['port'],
                'MYSQL_USER': config['user'],
                'MYSQL_PASSWORD': config['password'],
                'MYSQL_DATABASE': config['database']
            }
            crypto.save_encrypted(encrypted_config)
            logger.info("Crypto files created successfully")
        except Exception as e:
            logger.error(f"Error creating crypto files: {e}")
            messagebox.showerror("Erro", f"Erro ao criar arquivos de configuração: {e}")
    
    def create_database(self):
        """Create database and default tables"""
        config = {key: entry.get().strip() 
                for key, entry in self.entries.items()}
        db_name = config.get('database')
        
        if not db_name:
            messagebox.showerror("Erro", "Informe o nome do banco de dados!")
            return
            
        success, error = DatabaseOperations.create_database(config, db_name)
        if success:
            messagebox.showinfo("Sucesso", 
                            f"Banco de dados '{db_name}' criado com sucesso!")
            self.create_crypto_files(config)
        else:
            messagebox.showerror("Erro", f"Erro ao criar banco: {error}")