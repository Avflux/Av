def setup_sheets_tab(self):
        """Configura a aba de gerenciamento de planilhas"""
        # Título da aba
        title_frame = ctk.CTkFrame(self.tab_sheets, fg_color="transparent")
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
        sheets_frame = ctk.CTkFrame(self.tab_sheets, fg_color="transparent")
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
            command=self.search_sheets,
            fg_color=("#ff5722", "#ff5722"),
            hover_color=("#ce461b", "#ce461b")
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
            "nome": (200, "Nome"),
            "salvo_por": (200, "Salvo por"),
            "modificado_em": (150, "Modificado em")
        }
        
        for col, (width, text) in columns.items():
            self.sheets_tree.heading(col, text=text)
            self.sheets_tree.column(col, width=width, minwidth=width)
        
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
        
        self.dir_entry = ctk.CTkEntry(form_frame, 
            placeholder_text="Selecione o diretório")
        self.dir_entry.pack(fill="x", padx=10, pady=(5, 0))
        
        # Botões de ação
        button_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        button_frame.pack(fill="x", padx=10, pady=10)
        
        select_dir_button = ctk.CTkButton(
            button_frame,
            text="Selecionar",
            command=self.select_sheets_directory,
            fg_color=("#ff5722", "#ff5722"),
            hover_color=("#ce461b", "#ce461b")
        )
        select_dir_button.pack(fill="x", pady=(0, 5))
        
        refresh_button = ctk.CTkButton(
            button_frame,
            text="Atualizar Lista",
            command=self.load_sheets_list,
            fg_color=("#ff5722", "#ff5722"),
            hover_color=("#ce461b", "#ce461b")
        )
        refresh_button.pack(fill="x", pady=(0, 5))
        
        action_button = ctk.CTkButton(
            button_frame,
            text="Editar",
            fg_color=("#ff5722", "#ff5722"),
            hover_color=("#ce461b", "#ce461b")
        )
        action_button.pack(fill="x")
        
        # Carregar diretório padrão se existir
        self.load_default_sheets_directory()

    def select_sheets_directory(self):
        """Abre diálogo para selecionar diretório das planilhas"""
        try:
            directory = filedialog.askdirectory()
            if directory:
                self.dir_entry.delete(0, "end")
                self.dir_entry.insert(0, directory)
                # Não carrega automaticamente, usuário deve clicar em Atualizar
        except Exception as e:
            logger.error(f"Erro ao selecionar diretório: {e}")
            messagebox.showerror("Erro", f"Erro ao selecionar diretório: {e}")

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
                        
                        # Tentar obter o autor do arquivo usando o Shell do Windows
                        author = "Sistema"
                        try:
                            # Inicializar COM para thread atual
                            pythoncom.CoInitialize()
                            
                            # Obter o PIDL do arquivo
                            pidl = shell.SHParseDisplayName(str(file))[0]
                            
                            # Obter a interface do Shell para o arquivo
                            shell_folder = shell.SHGetDesktopFolder()
                            shell_item = shell_folder.BindToObject(
                                pidl, None, shell.IID_IShellItem2
                            )
                            
                            # Tentar obter o autor das propriedades do sistema
                            author = shell_item.GetString(shell.PKEY_FileOwner)
                            if not author:
                                author = "Sistema"
                        except Exception as e:
                            logger.error(f"Erro ao ler propriedades do arquivo {file}: {e}")
                        finally:
                            pythoncom.CoUninitialize()

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

            # Salvar o diretório no config.txt
            self.save_sheets_directory(directory)

        except Exception as e:
            logger.error(f"Erro ao carregar lista de planilhas: {e}")
            messagebox.showerror("Erro", f"Erro ao carregar lista de planilhas: {e}")

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
