from customtkinter import CTkFrame
import customtkinter as ctk
from tkinter import filedialog
from tkinter import ttk, messagebox
from datetime import datetime
from pathlib import Path
import logging
from PIL import Image
import os
from ....database.connection import DatabaseConnection
import pathlib

logger = logging.getLogger(__name__)

class SheetsTab(CTkFrame):
    def __init__(self, parent, user_data):
        super().__init__(parent)
        self.user_data = user_data
        self.db = DatabaseConnection()
        self.setup_sheets()
        self.pack(expand=True, fill="both")

    def setup_sheets(self):
        """Configura a aba de gerenciamento de planilhas"""
        # Título da aba
        title_frame = ctk.CTkFrame(self, fg_color="transparent")
        title_frame.pack(fill="x", padx=10, pady=(5, 15))
        
        title_label = ctk.CTkLabel(
            title_frame,
            text="Gerenciamento de Planilhas",
            font=("Roboto", 18, "bold"),
            text_color=("#ff5722", "#ff5722")
        )
        title_label.pack(side="left")
        
        # Contador de planilhas
        self.sheets_count_label = ctk.CTkLabel(
            title_frame,
            text="Total: 0 planilhas",
            font=("Roboto", 12),
            text_color=("#666666", "#999999")
        )
        self.sheets_count_label.pack(side="right")
        
        # Frame principal
        sheets_frame = ctk.CTkFrame(self, fg_color="transparent")
        sheets_frame.pack(expand=True, fill="both", padx=2, pady=2)
        
        # Configurar divisão 70/30
        sheets_frame.grid_columnconfigure(0, weight=7)
        sheets_frame.grid_columnconfigure(1, weight=3)
        sheets_frame.grid_rowconfigure(0, weight=1)
        
        # Frame esquerdo
        left_frame = ctk.CTkFrame(sheets_frame, fg_color=("#ffffff", "#2b2b2b"))
        left_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        left_frame.grid_columnconfigure(0, weight=1)
        left_frame.grid_rowconfigure(1, weight=1)
        
        # Frame de busca
        search_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        search_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=(5, 5))
        
        self.sheets_search_entry = ctk.CTkEntry(
            search_frame,
            placeholder_text="🔍 Buscar planilhas..."
        )
        self.sheets_search_entry.pack(side="left", fill="x", expand=True)
        self.sheets_search_entry.bind('<Return>', lambda e: self.search_sheets())
        
        search_button = ctk.CTkButton(
            search_frame,
            text="Buscar",
            width=80,
            height=32,
            command=self.search_sheets,
            fg_color="#ff5722",
            hover_color="#ce461b",
            font=("Roboto", 12)
        )
        search_button.pack(side="right", padx=(5, 0))
        
        # Frame da Treeview
        tree_frame = ctk.CTkFrame(left_frame)
        tree_frame.grid(row=1, column=0, sticky="nsew")
        tree_frame.grid_columnconfigure(0, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)
        
        # Treeview com estilo customizado
        self.sheets_tree = ttk.Treeview(
            tree_frame,
            columns=("nome", "salvo_por", "modificado_em"),
            show="headings",
            style="Custom.Treeview"
        )
        
        # Configurar colunas com proporções adequadas
        columns = {
            "nome": (235, "Nome", "w"),
            "salvo_por": (235, "Salvo por", "w"),
            "modificado_em": (80, "Modificado em", "center")
        }
        
        # Configurar colunas
        self.sheets_tree["columns"] = list(columns.keys())
        self.sheets_tree["show"] = "headings"
        
        for col, (width, heading, anchor) in columns.items():
            self.sheets_tree.heading(col, text=heading)
            self.sheets_tree.column(col, width=width, minwidth=width, anchor=anchor)
            
        # Configurar tags de cores para as linhas
        self.sheets_tree.tag_configure("today", background="#4CAF50", foreground="black")
        self.sheets_tree.tag_configure("recent", background="#FFC107", foreground="black")
        self.sheets_tree.tag_configure("old", background="#F44336", foreground="black")
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.sheets_tree.yview)
        self.sheets_tree.configure(yscrollcommand=scrollbar.set)
        
        # Grid para lista e scrollbar
        self.sheets_tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        # Frame direito (formulário)
        right_frame = ctk.CTkFrame(sheets_frame, fg_color="transparent")
        right_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0), pady=0)
        right_frame.grid_propagate(False)  # Impede que o frame encolha
        right_frame.configure(width=300)  # Largura fixa para o frame direito
        
        # Frame do formulário com borda e cor de fundo
        form_frame = ctk.CTkFrame(
            right_frame,
            width=300,
            height=400,
            fg_color=("#CFCFCF", "#333333"),
            border_width=1,
            border_color=("#e0e0e0", "#404040")
        )
        form_frame.pack(fill="both", padx=10, pady=5)
        form_frame.pack_propagate(False)
        
        # Label e entrada para o diretório
        dir_label = ctk.CTkLabel(
            form_frame, 
            text="Controle de Planilhas",
            font=("Roboto", 16, "bold"),
            text_color=("#ff5722", "#ff5722")
        )
        dir_label.pack(fill="x", padx=10, pady=(10, 0))
        
        # Frame para entrada e botão de seleção
        dir_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        dir_frame.pack(fill="x", padx=10, pady=(5, 0))
        
        self.dir_entry = ctk.CTkEntry(dir_frame, 
            placeholder_text="Selecione o diretório")
        self.dir_entry.pack(side="left", fill="x", expand=True)
        
        # Caminho absoluto do diretório do projeto
        project_root = pathlib.Path(__file__).resolve().parents[4]
        excel_icon_path = project_root / 'icons' / 'excel.png'
        excel_image = ctk.CTkImage(
            light_image=Image.open(excel_icon_path),
            dark_image=Image.open(excel_icon_path),
            size=(20, 20)
        )
        
        select_dir_button = ctk.CTkButton(
            dir_frame,
            text="",
            image=excel_image,
            width=30,
            command=self.select_sheets_directory,
            fg_color=("#ff5722", "#ff5722"),
            hover_color=("#ce461b", "#ce461b")
        )
        select_dir_button.pack(side="left", padx=(5, 0))
        
        # Botões de ação
        button_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        button_frame.pack(fill="x", padx=10, pady=10)
        
        # Frame para os botões lado a lado
        buttons_row = ctk.CTkFrame(button_frame, fg_color="transparent")
        buttons_row.pack(fill="x")
        buttons_row.grid_columnconfigure(0, weight=1)
        buttons_row.grid_columnconfigure(1, weight=1)
        
        refresh_button = ctk.CTkButton(
            buttons_row,
            text="Atualizar Lista",
            command=self.load_sheets_list,
            fg_color=("#ff5722", "#ff5722"),
            hover_color=("#ce461b", "#ce461b")
        )
        refresh_button.grid(row=0, column=0, sticky="ew", padx=(0, 2))
        
        action_button = ctk.CTkButton(
            buttons_row,
            text="Editar",
            fg_color=("#ff5722", "#ff5722"),
            hover_color=("#ce461b", "#ce461b"),
            command=self.abrir_excel_viewer
        )
        action_button.grid(row=0, column=1, sticky="ew", padx=(2, 0))
        
        # Legenda de cores
        legend_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        legend_frame.pack(fill="x", padx=10, pady=(10, 5))
        
        legend_title = ctk.CTkLabel(legend_frame, text="Legenda de cores:", anchor="w")
        legend_title.pack(fill="x", pady=(0, 5))
        
        # Cores para a legenda
        legend_items = [
            ("#4CAF50", "Hoje"),
            ("#FFC107", "Até 2 dias atrás"),
            ("#F44336", "3 ou mais dias atrás")
        ]
        
        for color, text in legend_items:
            item_frame = ctk.CTkFrame(legend_frame, fg_color="transparent")
            item_frame.pack(fill="x", pady=2)
            
            color_box = ctk.CTkFrame(item_frame, width=15, height=15, fg_color=color)
            color_box.pack(side="left", padx=(5, 10))
            
            label = ctk.CTkLabel(item_frame, text=text, anchor="w")
            label.pack(side="left", fill="x")
        
        # Carregar diretório padrão se existir
        self.load_default_sheets_directory()

    def abrir_excel_viewer(self):
        """Abre o visualizador de Excel com o arquivo selecionado"""
        try:
            # Verificar se há um item selecionado na Treeview
            selected_item = self.sheets_tree.focus()
            if not selected_item:
                messagebox.showinfo("Aviso", "Selecione uma planilha para editar.")
                return
                
            # Obter o nome do arquivo selecionado
            values = self.sheets_tree.item(selected_item, "values")
            if not values or len(values) < 1:
                messagebox.showinfo("Aviso", "Selecione uma planilha válida.")
                return
                
            file_name = values[0]  # Nome do arquivo sem extensão
            directory = self.dir_entry.get()
            
            if not directory or not os.path.exists(directory):
                messagebox.showerror("Erro", "Diretório inválido.")
                return
                
            # Caminho completo do arquivo
            file_path = os.path.join(directory, f"{file_name}.xlsx")
            
            if not os.path.exists(file_path):
                messagebox.showerror("Erro", f"Arquivo não encontrado: {file_path}")
                return
                
            # Importar a classe ExcelViewer
            from ..Excel_Viewer import ExcelViewer
            
            # Criar e mostrar a janela do visualizador
            excel_viewer = ExcelViewer(parent=self.winfo_toplevel())
            
            # Configurar o visualizador para abrir o arquivo diretamente
            excel_viewer.after(100, lambda: self.carregar_arquivo_excel(excel_viewer, file_path))
            
            # Mostrar a janela
            excel_viewer.mainloop()
            
        except Exception as e:
            logger.error(f"Erro ao abrir o visualizador de Excel: {e}")
            messagebox.showerror("Erro", f"Erro ao abrir o visualizador: {e}")
    
    def carregar_arquivo_excel(self, viewer, file_path):
        """Carrega o arquivo Excel diretamente no visualizador"""
        try:
            # Chamar o método abrir_excel do visualizador com o caminho do arquivo
            # Primeiro, vamos modificar temporariamente a função filedialog.askopenfilename
            # para retornar nosso caminho
            original_askopenfilename = filedialog.askopenfilename
            filedialog.askopenfilename = lambda **kwargs: file_path
            
            # Chamar o método abrir_excel
            viewer.abrir_excel()
            
            # Restaurar a função original
            filedialog.askopenfilename = original_askopenfilename
            
        except Exception as e:
            logger.error(f"Erro ao carregar arquivo Excel: {e}")

    def load_sheets_list(self, search_term=None):
        """Carrega a lista de planilhas do diretório selecionado"""
        try:
            directory = self.dir_entry.get()
            if not directory or not os.path.exists(directory):
                messagebox.showerror("Erro", "Selecione um diretório válido!")
                return

            if not hasattr(self, 'cached_files') or not search_term:
                # Limpar a Treeview e cache se não houver busca
                self.cached_files = []
                for item in self.sheets_tree.get_children():
                    self.sheets_tree.delete(item)

                # Listar arquivos .xlsx
                for file in Path(directory).glob("*.xlsx"):
                    try:
                        # Obter informações do arquivo
                        stats = file.stat()
                        modified_time = datetime.fromtimestamp(stats.st_mtime)
                        
                        # Obter o proprietário do arquivo
                        author = self.get_file_owner_advanced(str(file))

                        # Armazenar no cache
                        self.cached_files.append({
                            'name': file.stem,
                            'author': author,
                            'modified': modified_time.strftime("%d/%m/%Y %H:%M")
                        })
                    except Exception as e:
                        logger.error(f"Erro ao processar arquivo {file}: {e}")
                        continue

            # Limpar a Treeview para nova exibição
            for item in self.sheets_tree.get_children():
                self.sheets_tree.delete(item)

            # Filtrar e exibir arquivos
            count = 0
            for file_info in self.cached_files:
                if not search_term or search_term.lower() in file_info['name'].lower():
                    # Determinar a tag de cor com base na data de modificação
                    tag = self.get_date_tag(file_info['modified'])
                    
                    # Inserir item com a tag de cor apropriada
                    self.sheets_tree.insert(
                        "",
                        "end",
                        values=(
                            file_info['name'],
                            file_info['author'],
                            file_info['modified']
                        ),
                        tags=(tag,)
                    )
                    count += 1

            # Atualizar contador
            self.sheets_count_label.configure(text=f"Total: {count} planilha{'s' if count != 1 else ''}")

            # Salvar o diretório no config.txt
            self.save_sheets_directory(directory)

        except Exception as e:
            logger.error(f"Erro ao carregar lista de planilhas: {e}")
            messagebox.showerror("Erro", f"Erro ao carregar lista de planilhas: {e}")

    def select_sheets_directory(self):
        """Abre diálogo para selecionar diretório das planilhas"""
        try:
            directory = filedialog.askdirectory()
            if directory:
                self.dir_entry.delete(0, "end")
                self.dir_entry.insert(0, directory)
                # Carrega a lista automaticamente após selecionar o diretório
                self.load_sheets_list()
        except Exception as e:
            logger.error(f"Erro ao selecionar diretório: {e}")
            messagebox.showerror("Erro", f"Erro ao selecionar diretório: {e}")

    def get_file_owner_advanced(self, file_path):
        """Obtém quem modificou o arquivo por último usando múltiplos métodos"""
        import platform
        
        try:
            if platform.system() == "Windows":
                # Método 1: Tentar ler metadados do Excel - FOCO NO ÚLTIMO MODIFICADOR
                try:
                    import openpyxl
                    
                    workbook = openpyxl.load_workbook(file_path, read_only=True)
                    # Priorizar lastModifiedBy (quem modificou por último)
                    if hasattr(workbook, 'properties') and workbook.properties:
                        if workbook.properties.lastModifiedBy:
                            last_modified = workbook.properties.lastModifiedBy
                            workbook.close()
                            return last_modified
                        # Só usar creator se lastModifiedBy não estiver disponível
                        if workbook.properties.creator:
                            creator = workbook.properties.creator
                            workbook.close()
                            return f"{creator} (criador)"
                    workbook.close()
                except ImportError:
                    pass
                except Exception as e:
                    logger.debug(f"Erro ao ler metadados Excel: {e}")
                    pass
                
                # Método 2: Tentar obter através do histórico de modificações (PowerShell avançado)
                try:
                    import subprocess
                    
                    # Escapar o caminho
                    escaped_path = file_path.replace("'", "''").replace('"', '""')
                    
                    # Comando PowerShell para obter informações detalhadas do arquivo
                    cmd = [
                        "powershell",
                        "-NoProfile",
                        "-NonInteractive",
                        "-Command",
                        f"""
                        $file = Get-Item '{escaped_path}'
                        $shell = New-Object -ComObject Shell.Application
                        $folder = $shell.Namespace($file.DirectoryName)
                        $item = $folder.ParseName($file.Name)
                        
                        # Tentar obter diferentes propriedades (Author, Last Author, etc.)
                        for ($i = 0; $i -lt 300; $i++) {{
                            $prop = $folder.GetDetailsOf($item, $i)
                            $header = $folder.GetDetailsOf($folder.Items, $i)
                            if ($header -like "*Author*" -or $header -like "*Modified*" -or $header -like "*Last*") {{
                                if ($prop -and $prop.Trim() -ne "") {{
                                    Write-Output "$header`: $prop"
                                }}
                            }}
                        }}
                        """
                    ]
                    
                    result = subprocess.run(
                        cmd, 
                        capture_output=True, 
                        text=True, 
                        timeout=15,
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                    
                    if result.returncode == 0 and result.stdout.strip():
                        lines = result.stdout.strip().split('\n')
                        last_author = None
                        author = None
                        
                        for line in lines:
                            if ':' in line:
                                header, value = line.split(':', 1)
                                header = header.strip().lower()
                                value = value.strip()
                                
                                if value and value != "":
                                    # Priorizar "Last Author" ou similares
                                    if 'last' in header and 'author' in header:
                                        last_author = value
                                    elif 'author' in header and not author:
                                        author = value
                        
                        # Retornar o último autor se disponível, senão o autor
                        if last_author:
                            return last_author
                        elif author:
                            return f"{author} (autor)"
                            
                except Exception as e:
                    logger.debug(f"Erro com PowerShell avançado: {e}")
                    pass
                
                # Método 2: Usar PowerShell para obter proprietário
                try:
                    import subprocess
                    
                    # Escapar o caminho para uso no PowerShell
                    escaped_path = file_path.replace("'", "''").replace('"', '""')
                    
                    # Comando PowerShell para obter proprietário
                    cmd = [
                        "powershell",
                        "-NoProfile",
                        "-NonInteractive",
                        "-Command",
                        f"(Get-Acl -Path '{escaped_path}').Owner"
                    ]
                    
                    result = subprocess.run(
                        cmd, 
                        capture_output=True, 
                        text=True, 
                        timeout=10,
                        creationflags=subprocess.CREATE_NO_WINDOW  # Não mostrar janela do PowerShell
                    )
                    
                    if result.returncode == 0 and result.stdout.strip():
                        owner = result.stdout.strip()
                        # Simplificar o nome (remover domínio se for local)
                        if '\\' in owner:
                            domain, user = owner.split('\\', 1)
                            # Se for domínio local ou BUILTIN, mostrar só o usuário
                            if domain.upper() in ['BUILTIN', 'NT AUTHORITY'] or domain == os.environ.get('COMPUTERNAME', ''):
                                return user
                            return owner
                        return owner
                        
                except Exception as e:
                    logger.debug(f"Erro com PowerShell: {e}")
                    pass
                
                # Método 3: Tentar usar WMI
                try:
                    import subprocess
                    
                    # Escapar o caminho
                    escaped_path = file_path.replace("\\", "\\\\").replace("'", "\\'")
                    
                    cmd = [
                        "wmic",
                        "datafile",
                        f"where name='{escaped_path}'",
                        "get",
                        "Owner",
                        "/format:value"
                    ]
                    
                    result = subprocess.run(
                        cmd,
                        capture_output=True,
                        text=True,
                        timeout=15,
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                    
                    if result.returncode == 0 and result.stdout:
                        for line in result.stdout.split('\n'):
                            if line.startswith('Owner=') and len(line) > 6:
                                owner = line[6:].strip()
                                if owner and owner != 'Owner=':
                                    return owner
                                
                except Exception as e:
                    logger.debug(f"Erro com WMI: {e}")
                    pass
                
                # Método 4: Fallback para usuário atual
                try:
                    import getpass
                    return getpass.getuser()
                except:
                    return "Sistema"
                    
            else:
                # Unix/Linux/macOS
                try:
                    import pwd
                    file_stat = os.stat(file_path)
                    owner = pwd.getpwuid(file_stat.st_uid).pw_name
                    return owner
                except (ImportError, KeyError):
                    try:
                        import getpass
                        return getpass.getuser()
                    except:
                        return "Desconhecido"
                        
        except Exception as e:
            logger.error(f"Erro ao obter proprietário do arquivo {file_path}: {e}")
            return "Desconhecido"

    def get_date_tag(self, date_str):
        """Determina a tag de cor com base na data de modificação"""
        try:
            # Converter a string de data para objeto datetime
            mod_date = datetime.strptime(date_str, "%d/%m/%Y %H:%M")
            
            # Obter a data atual
            today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
            
            # Calcular a diferença em dias
            days_diff = (today - mod_date.replace(hour=0, minute=0, second=0, microsecond=0)).days
            
            # Determinar a tag com base na diferença de dias
            if days_diff == 0:  # Hoje
                return "today"
            elif days_diff <= 2:  # Até 2 dias atrás
                return "recent"
            else:  # 3 ou mais dias atrás
                return "old"
        except Exception as e:
            logger.error(f"Erro ao determinar tag de data: {e}")
            return ""  # Sem tag em caso de erro
    
    def save_sheets_directory(self, directory):
        """Salva o diretório das planilhas no config.txt"""
        try:
            config_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__)))), 'config.txt')
            lines = []
            if os.path.exists(config_path):
                with open(config_path, 'r') as f:
                    lines = f.readlines()
                    while len(lines) < 3:
                        lines.append('\n')
            else:
                lines = ['\n'] * 3

            lines[2] = f'SHEETS_DIR={directory}\n'

            with open(config_path, 'w') as f:
                f.writelines(lines)
        except Exception as e:
            logger.error(f"Erro ao salvar diretório no config: {e}")

    def search_sheets(self):
        """Realiza a busca de planilhas com base no termo digitado"""
        search_term = self.sheets_search_entry.get().strip()
        
        # Limpar a Treeview
        for item in self.sheets_tree.get_children():
            self.sheets_tree.delete(item)

        # Filtrar e exibir arquivos do cache
        count = 0
        for file_info in self.cached_files:
            if not search_term or search_term.lower() in file_info['name'].lower():
                self.sheets_tree.insert(
                    "",
                    "end",
                    values=(
                        file_info['name'],
                        file_info['author'],
                        file_info['modified']
                    )
                )
                count += 1

        # Atualizar contador
        self.sheets_count_label.configure(text=f"Total: {count} planilha{'s' if count != 1 else ''}")

    def load_default_sheets_directory(self):
        """Carrega o diretório padrão das planilhas do config.txt"""
        try:
            config_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__)))), 'config.txt')
            if os.path.exists(config_path):
                with open(config_path, 'r') as f:
                    lines = f.readlines()
                    if len(lines) >= 3:
                        if lines[2].strip().startswith('SHEETS_DIR='):
                            directory = lines[2].strip().split('=')[1]
                            if os.path.exists(directory):
                                self.dir_entry.insert(0, directory)
                                return

            # Se não encontrar no config, usar diretório padrão
            default_dir = os.path.join(os.path.expanduser("~"), "Documents", "Chronos", "Planilhas")
            if os.path.exists(default_dir):
                self.dir_entry.insert(0, default_dir)
        except Exception as e:
            logger.error(f"Erro ao carregar diretório padrão: {e}")
        