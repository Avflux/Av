import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import stat
from datetime import datetime
import platform

# Importações condicionais para sistemas Unix
try:
    import pwd
    import grp
    UNIX_SYSTEM = True
except ImportError:
    UNIX_SYSTEM = False

# Configuração do tema
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class ExcelFileBrowser:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("Navegador de Arquivos Excel")
        self.root.geometry("800x600")
        
        # Variável para armazenar o diretório atual
        self.current_directory = ""
        
        self.setup_ui()
        
    def setup_ui(self):
        # Frame principal
        main_frame = ctk.CTkFrame(self.root)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Frame superior para seleção de diretório
        top_frame = ctk.CTkFrame(main_frame)
        top_frame.pack(fill="x", padx=10, pady=10)
        
        # Label e botão para seleção de diretório
        directory_label = ctk.CTkLabel(top_frame, text="Diretório:")
        directory_label.pack(side="left", padx=(10, 5))
        
        self.directory_var = ctk.StringVar(value="Nenhum diretório selecionado")
        self.directory_display = ctk.CTkLabel(top_frame, textvariable=self.directory_var, 
                                            width=400, anchor="w")
        self.directory_display.pack(side="left", padx=(0, 10), fill="x", expand=True)
        
        select_button = ctk.CTkButton(top_frame, text="Selecionar Diretório", 
                                    command=self.select_directory)
        select_button.pack(side="right", padx=10)
        
        # Frame para a lista de arquivos
        list_frame = ctk.CTkFrame(main_frame)
        list_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # Headers
        header_frame = ctk.CTkFrame(list_frame)
        header_frame.pack(fill="x", padx=10, pady=(10, 5))
        
        ctk.CTkLabel(header_frame, text="Nome do Arquivo", font=("Arial", 12, "bold"), 
                    width=250).pack(side="left", padx=5)
        ctk.CTkLabel(header_frame, text="Modificado por", font=("Arial", 12, "bold"), 
                    width=150).pack(side="left", padx=5)
        ctk.CTkLabel(header_frame, text="Data/Hora", font=("Arial", 12, "bold"), 
                    width=200).pack(side="left", padx=5)
        
        # Scrollable frame para a lista
        self.scrollable_frame = ctk.CTkScrollableFrame(list_frame)
        self.scrollable_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # Frame inferior para informações
        info_frame = ctk.CTkFrame(main_frame)
        info_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        self.info_label = ctk.CTkLabel(info_frame, text="Selecione um diretório para começar")
        self.info_label.pack(pady=10)
        
    def select_directory(self):
        directory = filedialog.askdirectory(title="Selecione o diretório")
        if directory:
            self.current_directory = directory
            self.directory_var.set(directory)
            self.load_excel_files()
    
    def get_file_owner(self, file_path):
        """Obtém o proprietário do arquivo"""
        try:
            if platform.system() == "Windows":
                # No Windows, obtém o proprietário usando WMI ou retorna usuário atual
                import getpass
                return getpass.getuser()
            elif UNIX_SYSTEM:
                # Unix/Linux/macOS
                file_stat = os.stat(file_path)
                try:
                    owner = pwd.getpwuid(file_stat.st_uid).pw_name
                    return owner
                except KeyError:
                    return f"UID: {file_stat.st_uid}"
            else:
                import getpass
                return getpass.getuser()
        except Exception as e:
            return "Desconhecido"
    
    def get_modification_time(self, file_path):
        """Obtém a data/hora de modificação do arquivo"""
        try:
            mtime = os.path.getmtime(file_path)
            return datetime.fromtimestamp(mtime).strftime("%d/%m/%Y %H:%M:%S")
        except Exception as e:
            return "Desconhecido"
    
    def clear_file_list(self):
        """Limpa a lista atual de arquivos"""
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
    
    def load_excel_files(self):
        """Carrega e exibe os arquivos .xlsx do diretório selecionado"""
        self.clear_file_list()
        
        try:
            if not os.path.exists(self.current_directory):
                messagebox.showerror("Erro", "Diretório não encontrado!")
                return
            
            excel_files = []
            
            # Busca arquivos .xlsx
            for filename in os.listdir(self.current_directory):
                if filename.lower().endswith('.xlsx'):
                    file_path = os.path.join(self.current_directory, filename)
                    
                    # Remove a extensão .xlsx
                    name_without_ext = filename[:-5]
                    
                    # Obtém informações do arquivo
                    owner = self.get_file_owner(file_path)
                    mod_time = self.get_modification_time(file_path)
                    
                    excel_files.append({
                        'name': name_without_ext,
                        'owner': owner,
                        'mod_time': mod_time,
                        'full_path': file_path
                    })
            
            # Ordena por nome
            excel_files.sort(key=lambda x: x['name'].lower())
            
            # Exibe os arquivos
            if excel_files:
                for i, file_info in enumerate(excel_files):
                    self.create_file_row(file_info, i)
                
                self.info_label.configure(text=f"Encontrados {len(excel_files)} arquivo(s) Excel")
            else:
                no_files_label = ctk.CTkLabel(self.scrollable_frame, 
                                            text="Nenhum arquivo .xlsx encontrado neste diretório",
                                            font=("Arial", 14))
                no_files_label.pack(pady=20)
                self.info_label.configure(text="Nenhum arquivo Excel encontrado")
                
        except PermissionError:
            messagebox.showerror("Erro", "Sem permissão para acessar o diretório!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar arquivos: {str(e)}")
    
    def create_file_row(self, file_info, index):
        """Cria uma linha para cada arquivo"""
        # Cores alternadas para melhor visualização
        bg_color = ("gray90", "gray20") if index % 2 == 0 else ("gray95", "gray25")
        
        row_frame = ctk.CTkFrame(self.scrollable_frame, fg_color=bg_color)
        row_frame.pack(fill="x", padx=5, pady=2)
        
        # Nome do arquivo
        name_label = ctk.CTkLabel(row_frame, text=file_info['name'], 
                                 width=250, anchor="w")
        name_label.pack(side="left", padx=10, pady=5)
        
        # Proprietário
        owner_label = ctk.CTkLabel(row_frame, text=file_info['owner'], 
                                  width=150, anchor="w")
        owner_label.pack(side="left", padx=5, pady=5)
        
        # Data/hora de modificação
        time_label = ctk.CTkLabel(row_frame, text=file_info['mod_time'], 
                                 width=200, anchor="w")
        time_label.pack(side="left", padx=5, pady=5)
        
        # Botão para abrir arquivo (opcional)
        open_button = ctk.CTkButton(row_frame, text="Abrir", width=60,
                                   command=lambda path=file_info['full_path']: self.open_file(path))
        open_button.pack(side="right", padx=10, pady=5)
    
    def open_file(self, file_path):
        """Abre o arquivo com o programa padrão do sistema"""
        try:
            if platform.system() == "Windows":
                os.startfile(file_path)
            elif platform.system() == "Darwin":  # macOS
                os.system(f"open '{file_path}'")
            else:  # Linux
                os.system(f"xdg-open '{file_path}'")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir o arquivo: {str(e)}")
    
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = ExcelFileBrowser()
    app.run()