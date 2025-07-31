from customtkinter import CTkFrame
import customtkinter as ctk
from tkinter import ttk, messagebox
import logging
import bcrypt
from datetime import datetime
from ....database.connection import DatabaseConnection

logger = logging.getLogger(__name__)

class UsersTab(CTkFrame):
    def __init__(self, parent, user_data, manager=None):
        super().__init__(parent)
        self.user_data = user_data
        self.manager = manager if manager is not None else None
        self.db = DatabaseConnection()
        self.entry_widgets = {}  # Inicializa o dicionário de widgets
        self.setup_users()
        self.pack(expand=True, fill="both")

    def setup_users(self):
        """Configura a aba de usuários"""
        # Título da aba com estilo melhorado
        title_frame = ctk.CTkFrame(self, fg_color="transparent")
        title_frame.pack(fill="x", padx=10, pady=(5, 15))
        
        title_label = ctk.CTkLabel(
            title_frame,
            text="Gerenciamento de Usuários",
            font=("Roboto", 18, "bold"),
            text_color=("#ff5722", "#ff5722")
        )
        title_label.pack(side="left")
        
        # Adicionar contador de usuários
        self.user_count_label = ctk.CTkLabel(
            title_frame,
            text="Total: 0 usuários",
            font=("Roboto", 12),
            text_color=("#666666", "#999999")
        )
        self.user_count_label.pack(side="right")

        users_frame = ctk.CTkFrame(self, fg_color="transparent")
        users_frame.pack(expand=True, fill="both", padx=2, pady=2)  # Ajustado para 2px como nas outras abas
        
        # Ajustar pesos do grid para divisão 70/30
        users_frame.grid_columnconfigure(0, weight=70)  # 70% para lista
        users_frame.grid_columnconfigure(1, weight=30)  # 30% para formulário
        users_frame.grid_rowconfigure(0, weight=1)
        
        # Frame esquerdo (lista de usuários)
        left_frame = ctk.CTkFrame(users_frame, fg_color="transparent")
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5), pady=0)
        left_frame.grid_propagate(False)  # Impede que o frame encolha
        left_frame.configure(width=600)  # Largura fixa para o frame esquerdo
        
        # Barra de pesquisa com ícone e estilo melhorado
        search_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        search_frame.pack(fill="x", padx=2, pady=(0, 2))
        
        self.search_entry = ctk.CTkEntry(
            search_frame,
            placeholder_text="🔍 Buscar usuário...",
            height=32,
            font=("Roboto", 12)
        )
        self.search_entry.pack(side="left", fill="x", expand=True, padx=(0, 2))
        self.search_entry.bind("<Return>", lambda event: self.search_users())
        
        search_btn = ctk.CTkButton(
            search_frame,
            text="Buscar",
            width=80,
            height=32,
            command=self.search_users,
            fg_color="#ff5722",
            hover_color="#ce461b",
            font=("Roboto", 12)
        )
        search_btn.pack(side="right")
        
        # Lista de usuários
        tree_frame = ctk.CTkFrame(left_frame)
        tree_frame.pack(fill="both", expand=True)
        tree_frame.grid_columnconfigure(0, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)
        
        # Criar Treeview
        self.tree = ttk.Treeview(
            tree_frame,
            columns=("ID", "Nome", "Email", "Equipe", "Tipo"),
            show="headings",
            style="Custom.Treeview"
        )
        
        # Configurar colunas com proporções adequadas
        columns = {
            "ID": 50,
            "Nome": 200,
            "Email": 150,
            "Equipe": 100,
            "Tipo": 80
        }
        
        for col, width in columns.items():
            self.tree.heading(col, text=col)
            self.tree.column(col, width=width, minwidth=width)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # Grid para lista e scrollbar
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        # Binding para seleção
        self.tree.bind("<<TreeviewSelect>>", self.on_user_select)
        
        # Frame direito (formulário)
        right_frame = ctk.CTkFrame(users_frame, fg_color="transparent")
        right_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0), pady=0)
        right_frame.grid_propagate(False)  # Impede que o frame encolha
        right_frame.configure(width=300)  # Largura fixa para o frame direito

        # Formulário de usuário
        self.setup_user_frame(right_frame)

        # Carregar usuários
        self.load_users()
        self.update_counter()
        # Removido: self.update_equipes_combobox() aqui

    def setup_user_frame(self, parent):
        """Configura o formulário de usuário com novo layout"""
        # Frame principal com borda e cor de fundo
        form_frame = ctk.CTkFrame(
            parent,
            width=300,
            height=400,
            fg_color=("#CFCFCF", "#333333"),
            border_width=1,
            border_color=("#e0e0e0", "#404040")
        )
        form_frame.pack(fill="both", padx=10, pady=5)
        form_frame.pack_propagate(False)  # Impede que o frame encolha
        
        # Título do formulário
        title = ctk.CTkLabel(
            form_frame,
            text="Dados do Usuário",
            font=("Roboto", 16, "bold"),
            text_color=("#ff5722", "#ff5722")
        )
        title.pack(pady=(10, 15))

        # Frame para os campos com padding consistente
        fields_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        fields_frame.pack(fill="both", expand=True, padx=15, pady=0)
        
        # Configuração dos campos
        fields = [
            ("nome", "Nome"),
            ("email", "Email"),
            ("name_id", "ID do Usuário"),
            ("senha", "Senha"),
            ("equipe", "Equipe"),
            ("tipo_usuario", "Tipo"),
            ("data_entrada", "Data Entrada")
        ]
        
        for field, label in fields:
            field_frame = ctk.CTkFrame(fields_frame, fg_color="transparent")
            field_frame.pack(fill="x", pady=3)
            
            if field == "equipe":
                if self.user_data['equipe_id'] == 1:
                    # Buscar equipes diretamente do banco
                    query = "SELECT nome FROM equipes ORDER BY nome"
                    result = self.db.execute_query(query)
                    equipes = [row['nome'] for row in result] if result else []
                    if not equipes:
                        equipes = ["Nenhuma equipe"]
                    widget = ctk.CTkComboBox(
                        field_frame,
                        values=equipes,
                        height=28
                    )
                    widget.set(equipes[0])
                else:
                    query = "SELECT nome FROM equipes WHERE id = %s"
                    result = self.db.execute_query(query, (self.user_data['equipe_id'],))
                    equipe_nome = result[0]['nome'] if result else ""
                    widget = ctk.CTkComboBox(
                        field_frame,
                        values=[equipe_nome] if equipe_nome else ["Nenhuma equipe"],
                        height=28
                    )
                    widget.set(equipe_nome)
                widget.pack(fill="x")
                self.entry_widgets[field] = widget
            elif field == "tipo_usuario":
                widget = ctk.CTkComboBox(
                    field_frame,
                    values=["admin", "master", "comum"],
                    height=28
                )
                widget.set("comum")
                widget.pack(fill="x")
                self.entry_widgets[field] = widget
            else:
                widget = ctk.CTkEntry(
                    field_frame,
                    placeholder_text=label,
                    height=28
                )
                if field == "senha":
                    widget.configure(show="*")
                elif field == "data_entrada":
                    widget.insert(0, datetime.now().strftime('%d/%m/%Y'))
                    widget.configure(state="readonly")
                widget.pack(fill="x")
                self.entry_widgets[field] = widget
        
        # Frame para botões com padding ajustado
        btn_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        btn_frame.pack(fill="x", padx=15, pady=(10, 5))
        
        # Configurar grid para 2 colunas
        btn_frame.grid_columnconfigure(0, weight=1)
        btn_frame.grid_columnconfigure(1, weight=1)
        
        # Botões na primeira linha
        save_btn = ctk.CTkButton(
            btn_frame,
            text="Salvar",
            height=28,
            command=self.save_user,
            fg_color="#ff5722",
            hover_color="#ce461b"
        )
        save_btn.grid(row=0, column=0, padx=2, pady=2, sticky="ew")
        
        clear_btn = ctk.CTkButton(
            btn_frame,
            text="Limpar",
            height=28,
            command=self.clear_form,
            fg_color=("#9E9E9E", "#404040"),
            hover_color=("#8E8E8E", "#505050"),
            text_color=("#ffffff", "#ffffff")
        )
        clear_btn.grid(row=0, column=1, padx=2, pady=2, sticky="ew")
        
        # Botão excluir
        delete_btn = ctk.CTkButton(
            form_frame,
            text="Excluir",
            height=28,
            command=self.delete_user,
            fg_color="#dc2626",
            hover_color="#991b1b"
        )
        delete_btn.pack(fill="x", padx=15, pady=(5, 10))

    def load_users(self):
        """Carrega todos os usuários na treeview"""
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # Query base
        base_query = """
            SELECT u.id, u.nome, u.email, e.nome as equipe, u.tipo_usuario
            FROM usuarios u
            LEFT JOIN equipes e ON u.equipe_id = e.id
        """
        
        # Se o usuário for da equipe 1, mostra todos os usuários
        # Caso contrário, mostra apenas os usuários da mesma equipe
        if self.user_data['equipe_id'] == 1:
            query = base_query
            params = ()
        else:
            query = base_query + " WHERE u.equipe_id = %s"
            params = (self.user_data['equipe_id'],)
            
        try:
            users = self.db.execute_query(query, params)
            for user in users:
                self.tree.insert("", "end", values=(
                    user['id'],
                    user['nome'],
                    user['email'],
                    user['equipe'],
                    user['tipo_usuario']
                ))
        except Exception as e:
            logger.error(f"Erro ao carregar usuários: {e}")
            messagebox.showerror("Erro", "Erro ao carregar lista de usuários")
    
    def search_users(self):
        """Pesquisa usuários com base no termo de busca"""
        search_term = self.search_entry.get()
        
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        if not search_term:
            self.load_users()
            return
            
        query = """
            SELECT u.id, u.nome, u.email, e.nome as equipe, u.tipo_usuario
            FROM usuarios u
            LEFT JOIN equipes e ON u.equipe_id = e.id
            WHERE u.equipe_id = %s
            AND u.nome LIKE %s
        """
        
        try:
            users = self.db.execute_query(query, (self.user_data['equipe_id'], f"%{search_term}%"))
            for user in users:
                self.tree.insert("", "end", values=(
                    user['id'],
                    user['nome'],
                    user['email'],
                    user['equipe'],
                    user['tipo_usuario']
                ))
        except Exception as e:
            logger.error(f"Erro ao pesquisar usuários: {e}")
            messagebox.showerror("Erro", "Erro ao pesquisar usuários")
            
    def on_user_select(self, event):
        """Carrega os dados do usuário selecionado no formulário"""
        selected_item = self.tree.selection()
        if not selected_item:
            return
            
        user_id = self.tree.item(selected_item)['values'][0]
        
        query = """
            SELECT u.id, u.nome, u.email, u.name_id, u.tipo_usuario, u.data_entrada, e.nome as equipe_nome
            FROM usuarios u
            LEFT JOIN equipes e ON u.equipe_id = e.id
            WHERE u.id = %s
        """
        
        try:
            result = self.db.execute_query(query, (user_id,))
            if result:
                user = result[0]
                self.current_user_id = user_id
                
                # Preencher formulário
                self.entry_widgets['nome'].delete(0, 'end')
                self.entry_widgets['nome'].insert(0, user['nome'])
                
                self.entry_widgets['email'].delete(0, 'end')
                self.entry_widgets['email'].insert(0, user['email'])
                
                self.entry_widgets['name_id'].delete(0, 'end')
                self.entry_widgets['name_id'].insert(0, user['name_id'])
                
                # Limpar campo de senha e preencher com asteriscos
                self.entry_widgets['senha'].delete(0, 'end')
                self.entry_widgets['senha'].insert(0, '********')
                
                if user['equipe_nome']:
                    self.entry_widgets['equipe'].set(user['equipe_nome'])
                
                self.entry_widgets['tipo_usuario'].set(user['tipo_usuario'])
                
                self.entry_widgets['data_entrada'].delete(0, 'end')
                if user['data_entrada']:
                    formatted_date = user['data_entrada'].strftime('%d/%m/%Y')
                    self.entry_widgets['data_entrada'].insert(0, formatted_date)
                
        except Exception as e:
            logger.error(f"Erro ao carregar dados do usuário: {e}")
            messagebox.showerror("Erro", "Erro ao carregar dados do usuário")
    
    def save_user(self):
        """Salva ou atualiza um usuário"""
        try:
            data = {field: widget.get() for field, widget in self.entry_widgets.items()}
            
            # Validações
            required_fields = ['nome', 'email', 'name_id']
            if not all(data[field] for field in required_fields):
                messagebox.showerror("Erro", "Todos os campos são obrigatórios!")
                return
            
            # Verificar se já existe name_id igual (apenas para novo usuário)
            if not self.current_user_id:
                check_query = "SELECT id FROM usuarios WHERE name_id = %s"
                check_result = self.db.execute_query(check_query, (data['name_id'],))
                if check_result:
                    messagebox.showerror("Erro", "Já existe um usuário com este ID (ID do usuário). Escolha outro.")
                    return
            
            # Tratamento da data
            try:
                # Converter data do formato dd/mm/aaaa para aaaa-mm-dd
                date_parts = data['data_entrada'].split('/')
                if len(date_parts) == 3:
                    data['data_entrada'] = f"{date_parts[2]}-{date_parts[1]}-{date_parts[0]}"
                else:
                    raise ValueError("Formato de data inválido")
            except ValueError as e:
                messagebox.showerror("Erro", "Data inválida! Use o formato dd/mm/aaaa")
                return
            
            # Buscar ID da equipe
            equipe_query = "SELECT id FROM equipes WHERE nome = %s"
            equipe_result = self.db.execute_query(equipe_query, (data['equipe'],))
            if not equipe_result:
                messagebox.showerror("Erro", "Equipe não encontrada!")
                return
            
            equipe_id = equipe_result[0]['id']
            
            # Tratamento especial para senha
            if self.current_user_id:  # Atualização
                if data['senha'] == '********':
                    # Senha não foi alterada, não incluir no UPDATE
                    query = """
                        UPDATE usuarios SET
                            nome = %s,
                            email = %s,
                            name_id = %s,
                            equipe_id = %s,
                            tipo_usuario = %s,
                            data_entrada = %s
                        WHERE id = %s
                    """
                    params = (
                        data['nome'],
                        data['email'],
                        data['name_id'],
                        equipe_id,
                        data['tipo_usuario'],
                        data['data_entrada'],
                        self.current_user_id
                    )
                else:
                    # Senha foi alterada, incluir no UPDATE
                    senha_hash = bcrypt.hashpw(data['senha'].encode('utf-8'), bcrypt.gensalt())
                    query = """
                        UPDATE usuarios SET
                            nome = %s,
                            email = %s,
                            name_id = %s,
                            senha = %s,
                            equipe_id = %s,
                            tipo_usuario = %s,
                            data_entrada = %s
                        WHERE id = %s
                    """
                    params = (
                        data['nome'],
                        data['email'],
                        data['name_id'],
                        senha_hash,
                        equipe_id,
                        data['tipo_usuario'],
                        data['data_entrada'],
                        self.current_user_id
                    )
                
                # Executar a query de atualização
                cursor = self.db.connection.cursor(dictionary=True)
                cursor.execute(query, params)
                self.db.connection.commit()
                cursor.close()
                
                message = "Usuário atualizado com sucesso!"
            else:  # Novo usuário
                query = """
                    INSERT INTO usuarios (
                        nome, email, name_id, senha, equipe_id,
                        tipo_usuario, data_entrada, status
                    ) VALUES (%s, %s, %s, %s, %s, %s, %s, TRUE)
                """
                params = (
                    data['nome'],
                    data['email'],
                    data['name_id'],
                    bcrypt.hashpw(data['senha'].encode('utf-8'), bcrypt.gensalt()),
                    equipe_id,
                    data['tipo_usuario'],
                    data['data_entrada']
                )
                message = "Usuário cadastrado com sucesso!"
                
                # Execute a inserção do usuário e obtenha o ID gerado
                cursor = self.db.connection.cursor(dictionary=True)
                cursor.execute(query, params)
                new_user_id = cursor.lastrowid
                
                # O trigger after_usuario_insert irá criar automaticamente o registro na tabela user_lock_unlock
                self.db.connection.commit()
                cursor.close()
            
            messagebox.showinfo("Sucesso", message)
            
            # Atualizar todas as listas e indicadores
            self.clear_form()
            self.load_users()
            self.load_blocks()  # Atualizar lista de bloqueios
            self.update_status_indicators()  # Atualizar indicadores
            
        except Exception as e:
            logger.error(f"Erro ao salvar usuário: {e}")
            messagebox.showerror("Erro", f"Erro ao salvar usuário: {e}")
    
    def delete_user(self):
        """Remove um usuário do sistema"""
        if not self.current_user_id:
            messagebox.showwarning("Aviso", "Selecione um usuário para deletar!")
            return
        
        if messagebox.askyesno("Confirmar", "Deseja realmente deletar este usuário?"):
            try:
                # Deletar usuário em vez de marcar como inativo
                query = "DELETE FROM usuarios WHERE id = %s"
                self.db.execute_query(query, (self.current_user_id,))
                
                messagebox.showinfo("Sucesso", "Usuário deletado com sucesso!")
                
                # Limpar formulário e atualizar todas as listas
                self.clear_form()
                self.load_users()
                self.load_blocks()
                self.update_status_indicators()
                
            except Exception as e:
                logger.error(f"Erro ao deletar usuário: {e}")
                messagebox.showerror("Erro", f"Erro ao deletar usuário: {e}")

    def clear_form(self):
        """Limpa o formulário e prepara para novo cadastro"""
        self.current_user_id = None
        # Limpar campos de texto mantendo placeholders
        for field in ['nome', 'email', 'name_id', 'senha']:
            if self.entry_widgets[field].get():  # Só limpa se tiver conteúdo
                self.entry_widgets[field].delete(0, 'end')
        # Resetar comboboxes
        self.update_equipes_combobox()
        self.entry_widgets['tipo_usuario'].set('comum')
        # Resetar data para atual
        self.entry_widgets['data_entrada'].configure(state="normal")
        self.entry_widgets['data_entrada'].delete(0, 'end')
        self.entry_widgets['data_entrada'].insert(0, datetime.now().strftime('%d/%m/%Y'))
        self.entry_widgets['data_entrada'].configure(state="readonly")

    def update_counter(self):
        """Atualiza o contador de usuários na aba."""
        try:
            if hasattr(self, 'tree') and hasattr(self, 'user_count_label'):
                user_count = len(self.tree.get_children())
                self.user_count_label.configure(text=f"Total: {user_count} usuário{'s' if user_count != 1 else ''}")
        except Exception as e:
            logger.error(f"Erro ao atualizar contador de usuários: {e}")

    def update_equipes_combobox(self):
        if 'equipe' in self.entry_widgets:
            equipes_raw = self.manager.get_equipes() if self.manager else []
            equipes = [e[0] if isinstance(e, (tuple, list)) else str(e) for e in equipes_raw]
            if not equipes:
                equipes = ["Nenhuma equipe"]
            self.entry_widgets['equipe'].configure(values=equipes)
            # Se o ComboBox estiver vazio, seleciona a primeira equipe
            if not self.entry_widgets['equipe'].get() and equipes:
                self.entry_widgets['equipe'].set(equipes[0])