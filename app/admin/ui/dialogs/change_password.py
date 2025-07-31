import customtkinter as ctk
from tkinter import messagebox
import logging
from app.admin.database.operations import DatabaseOperations

logger = logging.getLogger(__name__)

class ChangePasswordDialog(ctk.CTkToplevel):
    def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent
            self.title("Alterar Senha MySQL")
            self.setup_window()
            self.setup_ui()
            self.changed_password = None  # Novo atributo para armazenar a senha alterada
    
    def setup_window(self):
        """Configura a janela"""
        self.geometry("300x400")
        self.resizable(False, False)
        self.transient(self.parent)
        self.grab_set()
        self.center_window()
    
    def center_window(self):
        """Centraliza a janela"""
        x = (self.winfo_screenwidth() - 300) // 2
        y = (self.winfo_screenheight() - 400) // 2
        self.geometry(f"300x400+{x}+{y}")
    
    def setup_ui(self):
        # Title
        title_label = ctk.CTkLabel(
            self,
            text="Alterar Senha MySQL",
            font=("Roboto", 20, "bold")
        )
        title_label.pack(pady=(20, 30))
        
        # Current password
        current_frame = ctk.CTkFrame(self, fg_color="transparent")
        current_frame.pack(fill="x", padx=20, pady=5)
        
        current_label = ctk.CTkLabel(
            current_frame,
            text="Senha Atual:",
            font=("Roboto", 12)
        )
        current_label.pack(anchor="w")
        
        self.current_password = ctk.CTkEntry(
            current_frame,
            width=260,
            show="*"
        )
        self.current_password.pack(fill="x", pady=(5, 0))
        
        # New password
        new_frame = ctk.CTkFrame(self, fg_color="transparent")
        new_frame.pack(fill="x", padx=20, pady=5)
        
        new_label = ctk.CTkLabel(
            new_frame,
            text="Nova Senha:",
            font=("Roboto", 12)
        )
        new_label.pack(anchor="w")
        
        self.new_password = ctk.CTkEntry(
            new_frame,
            width=260,
            show="*"
        )
        self.new_password.pack(fill="x", pady=(5, 0))
        
        # Confirm password
        confirm_frame = ctk.CTkFrame(self, fg_color="transparent")
        confirm_frame.pack(fill="x", padx=20, pady=5)
        
        confirm_label = ctk.CTkLabel(
            confirm_frame,
            text="Confirmar Nova Senha:",
            font=("Roboto", 12)
        )
        confirm_label.pack(anchor="w")
        
        self.confirm_password = ctk.CTkEntry(
            confirm_frame,
            width=260,
            show="*"
        )
        self.confirm_password.pack(fill="x", pady=(5, 0))
        
        # Change button
        self.change_btn = ctk.CTkButton(
            self,
            text="Alterar Senha",
            fg_color="#FF5722",
            hover_color="#CE461B",
            command=self.change_password,
            width=200,
            font=("Roboto", 14)
        )
        self.change_btn.pack(pady=30)
        
    def change_password(self):
        """Change MySQL user password"""
        current = self.current_password.get()
        new = self.new_password.get()
        confirm = self.confirm_password.get()
        
        if not all([current, new, confirm]):
            messagebox.showerror("Erro", "Preencha todos os campos!")
            return
            
        if new != confirm:
            messagebox.showerror("Erro", "As senhas não coincidem!")
            return
            
        config = {key: self.parent.entries[key].get().strip() 
                for key in ['host', 'user', 'port']}
        
        success, error = DatabaseOperations.change_mysql_password(
            config, current, new
        )
        
        if success:
            messagebox.showinfo("Sucesso", "Senha alterada com sucesso!")
            self.changed_password = new  # Armazena a nova senha no atributo
            self.destroy()
        else:
            messagebox.showerror("Erro", f"Erro ao alterar senha: {error}")