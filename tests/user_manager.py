import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, messagebox
import logging
import os
import sys
from pathlib import Path
from datetime import datetime, timedelta
from tkinter import filedialog
import bcrypt
from PIL import Image
from ...database.connection import DatabaseConnection
from ...utils.excel_selector import ExcelSelector

logger = logging.getLogger(__name__)

class UserManagementFrame(ctk.CTkFrame):
    def __init__(self, parent, user_data):
        super().__init__(parent)
        
        # Armazenar dados do usuário logado
        self.user_data = user_data
        self.parent = parent
        
        # Configurar pesos do grid para responsividade
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        self.db = DatabaseConnection()
        
        # Dicionários para armazenar widgets
        self.entry_widgets = {}
        self.current_user_id = None
        
        self.setup_ui()
        
    def setup_ui(self):
        """Configura a interface principal com abas"""
        # Frame principal com padding reduzido e cor de fundo transparente
        self.main_frame = ctk.CTkFrame(self)
        
        # Tabs container com estilo melhorado
        self.tab_view = ctk.CTkTabview(
            self.main_frame,
            fg_color=("#DBDBDB", "#2b2b2b"),
            segmented_button_fg_color=("#ff5722", "#ff5722"),
            segmented_button_selected_color=("#ff5722", "#ff5722"),
            segmented_button_selected_hover_color=("#ce461b", "#ce461b"),
            segmented_button_unselected_color=("#8E8E8E", "#2b2b2b")
        )
        self.tab_view.pack(expand=True, fill="both")
        
        # Configurar pesos do grid no main_frame
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(0, weight=1)
        
        # Criar abas com nomes mais descritivos
        self.tab_users = self.tab_view.add("Gerenciar Usuários")
        self.tab_teams = self.tab_view.add("Gerenciar Equipes") 
        self.tab_blocks = self.tab_view.add("Controle de Acesso")
        self.tab_activities = self.tab_view.add("Gerenciar Atividades")
        self.tab_sheets = self.tab_view.add("Gerenciar Planilhas")
        
        # Handler para mudança de abas
        def on_tab_change():
            # Remove o foco de qualquer widget selecionado
            if self.focus_get():
                self.focus_set()
        
        # Configurar callback para mudança de aba
        self.tab_view.configure(command=on_tab_change)
        
        # Configurar pesos do grid em cada aba
        for tab in [self.tab_users, self.tab_teams, self.tab_blocks, self.tab_activities, self.tab_sheets]:
            tab.grid_columnconfigure(0, weight=1)
            tab.grid_rowconfigure(0, weight=1)
        
        # Configurar cada aba
        self.setup_users_tab()
        self.setup_teams_tab()
        self.setup_blocks_tab()
        self.setup_activities_tab()
        self.setup_sheets_tab()
        
        # Empacotar o main_frame por último
        self.main_frame.pack(expand=True, fill="both", padx=15, pady=15)
        
    def setup_users_tab(self):
        """Configura a aba de usuários"""
        # Título da aba com estilo melhorado
        title_frame = ctk.CTkFrame(self.tab_users, fg_color="transparent")
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

        users_frame = ctk.CTkFrame(self.tab_users, fg_color="transparent")
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
        self.setup_user_form(right_frame)
        
        # Carregar usuários
        self.load_users()
        self.update_counters()

    def setup_user_form(self, parent):
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
            
            if field in ["equipe", "tipo_usuario"]:
                if field == "equipe":
                    # Se o usuário não for da equipe 1, mostrar apenas a equipe dele
                    if self.user_data['equipe_id'] == 1:
                        widget = ctk.CTkComboBox(
                            field_frame,
                            values=self.get_equipes(),
                            height=28
                        )
                    else:
                        # Buscar o nome da equipe do usuário
                        query = "SELECT nome FROM equipes WHERE id = %s"
                        result = self.db.execute_query(query, (self.user_data['equipe_id'],))
                        equipe_nome = result[0]['nome'] if result else ""
                        
                        widget = ctk.CTkComboBox(
                            field_frame,
                            values=[equipe_nome],
                            height=28
                        )
                        widget.set(equipe_nome)  # Define o valor inicial
                else:
                    widget = ctk.CTkComboBox(
                        field_frame,
                        values=["admin", "master", "comum"],
                        height=28
                    )
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
        
    def setup_teams_tab(self):
        """Configura a aba de equipes"""
        # Título com estatísticas
        title_frame = ctk.CTkFrame(self.tab_teams, fg_color="transparent")
        title_frame.pack(fill="x", padx=10, pady=(5, 15))
        
        title_label = ctk.CTkLabel(
            title_frame,
            text="Gerenciamento de Equipes",
            font=("Roboto", 18, "bold"),
            text_color=("#ff5722", "#ff5722")
        )
        title_label.pack(side="left")
        
        self.team_count_label = ctk.CTkLabel(
            title_frame,
            text="Total: 0 equipes",
            font=("Roboto", 12),
            text_color=("#666666", "#999999")
        )
        self.team_count_label.pack(side="right")

        teams_frame = ctk.CTkFrame(self.tab_teams, fg_color="transparent")
        teams_frame.pack(expand=True, fill="both", padx=2, pady=2)
        
        # Ajustar pesos do grid para divisão 70/30
        teams_frame.grid_columnconfigure(0, weight=70)
        teams_frame.grid_columnconfigure(1, weight=30)
        teams_frame.grid_rowconfigure(0, weight=1)
        
        # Frame esquerdo (lista de equipes)
        left_frame = ctk.CTkFrame(teams_frame, fg_color="transparent")
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5), pady=0)
        left_frame.grid_propagate(False)  # Impede que o frame encolha
        left_frame.configure(width=600)  # Largura fixa para o frame esquerdo
        
        # Configurar peso da linha no left_frame
        left_frame.grid_columnconfigure(0, weight=1)
        left_frame.grid_rowconfigure(0, weight=0)  # Search frame
        left_frame.grid_rowconfigure(1, weight=1)  # Tree frame
        
        # Barra de pesquisa
        search_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        search_frame.grid(row=0, column=0, sticky="ew", padx=2, pady=(0, 2))
        
        self.teams_search_entry = ctk.CTkEntry(
            search_frame,
            placeholder_text="🔍 Buscar equipe...",
            height=32
        )
        self.teams_search_entry.pack(side="left", fill="x", expand=True, padx=(0, 2))
        self.teams_search_entry.bind("<Return>", lambda event: self.search_teams())
        
        search_btn = ctk.CTkButton(
            search_frame,
            text="Buscar",
            width=80,
            height=32,
            command=self.search_teams,
            fg_color="#ff5722",
            hover_color="#ce461b"
        )
        search_btn.pack(side="right")
        
        # Lista de equipes
        tree_frame = ctk.CTkFrame(left_frame)
        tree_frame.grid(row=1, column=0, sticky="nsew")
        tree_frame.grid_columnconfigure(0, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)
        
        # Criar Treeview para equipes
        self.teams_tree = ttk.Treeview(
            tree_frame,
            columns=("ID", "Nome", "Membros", "Data Criação"),
            show="headings",
            style="Custom.Treeview"
        )
        
        # Configurar colunas
        columns = {
            "ID": 50,
            "Nome": 200,
            "Membros": 150,
            "Data Criação": 180
        }
        
        for col, width in columns.items():
            self.teams_tree.heading(col, text=col)
            self.teams_tree.column(col, width=width, minwidth=width)
        
        # Adicionar scrollbar
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.teams_tree.yview)
        self.teams_tree.configure(yscrollcommand=scrollbar.set)
        
        # Grid para lista e scrollbar
        self.teams_tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        # Frame direito (ações)
        right_frame = ctk.CTkFrame(teams_frame, fg_color="transparent")
        right_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0), pady=0)
        right_frame.grid_propagate(False)  # Impede que o frame encolha
        right_frame.configure(width=300)  # Largura fixa para o frame direito
        
        # Frame para nova equipe com borda e cor de fundo
        new_team_frame = ctk.CTkFrame(
            right_frame,
            width=300,
            height=400,
            fg_color=("#CFCFCF", "#333333"),  # Cor de fundo atualizada para modo claro
            border_width=1,
            border_color=("#e0e0e0", "#404040")
        )
        new_team_frame.pack(fill="both", padx=10, pady=5)
        new_team_frame.pack_propagate(False)
        
        # Título dentro do frame
        title = ctk.CTkLabel(
            new_team_frame,
            text="Nova Equipe",
            font=("Roboto", 16, "bold"),
            text_color=("#ff5722", "#ff5722")
        )
        title.pack(pady=(10, 15))
        
        # Frame para campos
        fields_frame = ctk.CTkFrame(new_team_frame, fg_color="transparent")
        fields_frame.pack(fill="both", expand=True, padx=15, pady=0)
        
        # Label para o campo
        ctk.CTkLabel(
            fields_frame,
            text="Nome da Equipe",
            font=("Roboto", 11),
            anchor="w"
        ).pack(fill="x", padx=2, pady=(0, 2))
        
        # Entrada para nova equipe
        self.team_entry = ctk.CTkEntry(
            fields_frame,
            placeholder_text="Nome da nova equipe",
            height=28
        )
        self.team_entry.pack(fill="x", pady=3)
        self.team_entry.bind("<Return>", lambda event: self.add_team())
        
        # Frame para botões
        btn_frame = ctk.CTkFrame(new_team_frame, fg_color="transparent")
        btn_frame.pack(fill="x", padx=15, pady=(10, 5))
        
        # Configurar grid para 2 colunas
        btn_frame.grid_columnconfigure(0, weight=1)
        btn_frame.grid_columnconfigure(1, weight=1)
        
        # Botões na primeira linha
        add_team_btn = ctk.CTkButton(
            btn_frame,
            text="Adicionar",
            height=28,
            command=self.add_team,
            fg_color="#ff5722",
            hover_color="#ce461b"
        )
        add_team_btn.grid(row=0, column=0, padx=2, pady=2, sticky="ew")
        
        clear_btn = ctk.CTkButton(
            btn_frame,
            text="Limpar",
            height=28,
            command=lambda: self.team_entry.delete(0, 'end'),
            fg_color=("#9E9E9E", "#404040"),
            hover_color=("#8E8E8E", "#505050"),
            text_color=("#ffffff", "#ffffff")
        )
        clear_btn.grid(row=0, column=1, padx=2, pady=2, sticky="ew")
        
        # Botão excluir equipe
        delete_team_btn = ctk.CTkButton(
            new_team_frame,
            text="Excluir Equipe",
            height=28,
            command=self.delete_selected_team,
            fg_color="#dc2626",
            hover_color="#991b1b"
        )
        delete_team_btn.pack(fill="x", padx=15, pady=(5, 10))
        
        # Manter menu de contexto
        self.teams_tree.bind("<Button-3>", self.show_team_context_menu)
        
        # Carregar dados
        self.load_teams()
        self.update_counters()

    def setup_blocks_tab(self):
        """Configura a aba de bloqueios"""
        # Título com estatísticas
        title_frame = ctk.CTkFrame(self.tab_blocks, fg_color="transparent")
        title_frame.pack(fill="x", padx=10, pady=(5, 15))
        
        title_label = ctk.CTkLabel(
            title_frame,
            text="Controle de Acesso",
            font=("Roboto", 18, "bold"),
            text_color=("#ff5722", "#ff5722")
        )
        title_label.pack(side="left")
        
        self.block_stats_label = ctk.CTkLabel(
            title_frame,
            text="Bloqueados: 0 | Liberados: 0",
            font=("Roboto", 12),
            text_color=("#666666", "#999999")
        )
        self.block_stats_label.pack(side="right")

        blocks_frame = ctk.CTkFrame(self.tab_blocks, fg_color="transparent")
        blocks_frame.pack(expand=True, fill="both", padx=2, pady=2)
        
        # Ajustar pesos do grid para divisão 70/30
        blocks_frame.grid_columnconfigure(0, weight=70)  # 70% para lista
        blocks_frame.grid_columnconfigure(1, weight=30)  # 30% para ações
        blocks_frame.grid_rowconfigure(0, weight=1)
        
        # Frame esquerdo (lista de bloqueios)
        left_frame = ctk.CTkFrame(blocks_frame, fg_color="transparent")
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5), pady=0)
        left_frame.grid_propagate(False)  # Impede que o frame encolha
        left_frame.configure(width=600)  # Largura fixa para o frame esquerdo
        
        # Configurar peso da linha no left_frame
        left_frame.grid_columnconfigure(0, weight=1)
        left_frame.grid_rowconfigure(0, weight=0)  # Search frame
        left_frame.grid_rowconfigure(1, weight=1)  # Tree frame
        
        # Barra de pesquisa
        search_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        search_frame.grid(row=0, column=0, sticky="ew", padx=2, pady=(0, 2))
        
        self.blocks_search_entry = ctk.CTkEntry(
            search_frame,
            placeholder_text="🔍 Buscar usuário bloqueado...",
            height=32
        )
        self.blocks_search_entry.pack(side="left", fill="x", expand=True, padx=(0, 2))
        self.blocks_search_entry.bind("<Return>", lambda event: self.search_blocks())
        
        search_btn = ctk.CTkButton(
            search_frame,
            text="Buscar",
            width=80,
            height=32,
            command=self.search_blocks,
            fg_color="#ff5722",
            hover_color="#ce461b"
        )
        search_btn.pack(side="right")
        
        # Lista de bloqueios
        tree_frame = ctk.CTkFrame(left_frame)
        tree_frame.grid(row=1, column=0, sticky="nsew")
        tree_frame.grid_columnconfigure(0, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)
        
        # Criar Treeview para bloqueios
        self.blocks_tree = ttk.Treeview(
            tree_frame,
            columns=("ID", "Nome", "Equipe", "Status", "Controle"),
            show="headings",
            style="Custom.Treeview"
        )
        
        # Configurar colunas
        columns = {
            "ID": 50,
            "Nome": 200,
            "Equipe": 150,
            "Status": 100,
            "Controle": 80
        }
        
        for col, width in columns.items():
            self.blocks_tree.heading(col, text=col)
            self.blocks_tree.column(col, width=width, minwidth=width)
        
        # Adicionar scrollbar
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.blocks_tree.yview)
        self.blocks_tree.configure(yscrollcommand=scrollbar.set)
        
        # Grid para lista e scrollbar
        self.blocks_tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        # Frame direito (ações)
        right_frame = ctk.CTkFrame(blocks_frame, fg_color="transparent")
        right_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0), pady=0)
        right_frame.grid_propagate(False)  # Impede que o frame encolha
        right_frame.configure(width=300)  # Largura fixa para o frame direito
        
        # Frame para status do sistema com borda e cor de fundo
        status_frame = ctk.CTkFrame(
            right_frame,
            width=300,
            height=400,
            fg_color=("#CFCFCF", "#333333"),  # Cor de fundo atualizada para modo claro
            border_width=1,
            border_color=("#e0e0e0", "#404040")
        )
        status_frame.pack(fill="both", padx=10, pady=5)
        status_frame.pack_propagate(False)
        
        # Título do frame
        status_title = ctk.CTkLabel(
            status_frame,
            text="Status do Sistema",
            font=("Roboto", 16, "bold"),
            text_color=("#ff5722", "#ff5722")
        )
        status_title.pack(pady=(10, 15))
        
        # Frame para indicadores com padding consistente
        indicators_frame = ctk.CTkFrame(status_frame, fg_color="transparent")
        indicators_frame.pack(fill="both", expand=True, padx=15, pady=0)
        
        # Indicadores de status com espaçamento consistente
        self.status_indicators = {}
        for status in ["Usuários Ativos", "Usuários Bloqueados", "Tentativas de Login"]:
            indicator_frame = ctk.CTkFrame(indicators_frame, fg_color="transparent")
            indicator_frame.pack(fill="x", pady=3)
            
            # Label do indicador
            ctk.CTkLabel(
                indicator_frame,
                text=status,
                font=("Roboto", 11),
                anchor="w"
            ).pack(side="left", padx=2)
            
            # Valor do indicador
            self.status_indicators[status] = ctk.CTkLabel(
                indicator_frame,
                text="0",
                font=("Roboto", 12, "bold"),
                text_color=("#ff5722", "#ff5722")
            )
            self.status_indicators[status].pack(side="right", padx=2)
        
        # Frame para botões com padding consistente
        btn_frame = ctk.CTkFrame(status_frame, fg_color="transparent")
        btn_frame.pack(fill="x", padx=15, pady=(10, 10))
        
        # Botão de alternar bloqueio
        self.block_btn = ctk.CTkButton(
            btn_frame,
            text="Bloquear/Desbloquear",
            command=self.toggle_user_lock,
            height=28,
            fg_color="#dc2626",
            hover_color="#991b1b"
        )
        self.block_btn.pack(fill="x")
        
        # Binding para menu de contexto
        self.blocks_tree.bind("<Button-3>", self.show_block_context_menu)
        
        # Carregar bloqueios
        self.load_blocks()
        self.update_counters()

    def setup_activities_tab(self):
        """Configura a aba de atividades"""
        # Título da aba com estatísticas
        title_frame = ctk.CTkFrame(self.tab_activities, fg_color="transparent")
        title_frame.pack(fill="x", padx=10, pady=(5, 15))
        
        title_label = ctk.CTkLabel(
            title_frame,
            text="Gerenciamento de Atividades",
            font=("Roboto", 18, "bold"),
            text_color=("#ff5722", "#ff5722")
        )
        title_label.pack(side="left")
        
        self.activities_count_label = ctk.CTkLabel(
            title_frame,
            text="Total: 0 atividades",
            font=("Roboto", 12),
            text_color=("#666666", "#999999")
        )
        self.activities_count_label.pack(side="right")

        activities_frame = ctk.CTkFrame(self.tab_activities, fg_color="transparent")
        activities_frame.pack(expand=True, fill="both", padx=2, pady=2)
        
        # Ajustar pesos do grid para divisão 70/30
        activities_frame.grid_columnconfigure(0, weight=70)  # 70% para lista
        activities_frame.grid_columnconfigure(1, weight=30)  # 30% para formulário
        activities_frame.grid_rowconfigure(0, weight=1)
        
        # Frame esquerdo (lista de atividades)
        left_frame = ctk.CTkFrame(activities_frame, fg_color="transparent")
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5), pady=0)
        left_frame.grid_propagate(False)
        left_frame.configure(width=600)
        
        # Configurar peso da linha no left_frame
        left_frame.grid_columnconfigure(0, weight=1)
        left_frame.grid_rowconfigure(0, weight=0)  # Search frame
        left_frame.grid_rowconfigure(1, weight=1)  # Tree frame
        
        # Barra de pesquisa
        search_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        search_frame.grid(row=0, column=0, sticky="ew", padx=2, pady=(0, 2))
        search_frame.grid_columnconfigure(0, weight=1)  # Faz o entry expandir
        search_frame.grid_columnconfigure(1, weight=0)  # Botão mantém tamanho fixo
        
        self.activities_search_entry = ctk.CTkEntry(
            search_frame,
            placeholder_text="🔍 Buscar atividade...",
            height=32,
            font=("Roboto", 12)
        )
        self.activities_search_entry.grid(row=0, column=0, sticky="ew", padx=(0, 2))
        self.activities_search_entry.bind("<Return>", lambda event: self.search_activities())
        
        search_btn = ctk.CTkButton(
            search_frame,
            text="Buscar",
            width=80,
            height=32,
            command=self.search_activities,  # Updated command
            fg_color="#ff5722",
            hover_color="#ce461b",
            font=("Roboto", 12)
        )
        search_btn.grid(row=0, column=1, sticky="e")
        
        # Lista de atividades
        tree_frame = ctk.CTkFrame(left_frame)
        tree_frame.grid(row=1, column=0, sticky="nsew")
        tree_frame.grid_columnconfigure(0, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)
        
        # Criar Treeview com estilo moderno
        self.activities_tree = ttk.Treeview(
            tree_frame,
            columns=("Nome", "Descrição", "Atividade", "Tempo Total", "Status"),
            show="headings",
            style="Custom.Treeview"
        )
        
        # Configurar colunas
        columns = {
            "Nome": 100,
            "Descrição": 170,
            "Atividade": 170,
            "Tempo Total": 80,
            "Status": 80
        }
        
        for col, width in columns.items():
            self.activities_tree.heading(col, text=col)
            self.activities_tree.column(col, width=width, minwidth=width)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.activities_tree.yview)
        self.activities_tree.configure(yscrollcommand=scrollbar.set)
        
        # Grid para lista e scrollbar
        self.activities_tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        # Adicionar binding para duplo clique
        self.activities_tree.bind("<Double-1>", self.on_activity_double_click)
        
        # Carregar atividades
        self.load_activities()
        
        # Frame direito (formulário)
        right_frame = ctk.CTkFrame(activities_frame, fg_color="transparent")
        right_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0), pady=0)
        right_frame.grid_propagate(False)
        right_frame.configure(width=300)
        
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
        
        # Título do formulário
        title = ctk.CTkLabel(
            form_frame,
            text="Dados da Atividade",
            font=("Roboto", 16, "bold"),
            text_color=("#ff5722", "#ff5722")
        )
        title.pack(pady=(10, 0))

        # Frame do botão de reset, acima dos campos, alinhado à direita
        reset_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        reset_frame.pack(fill="x", padx=1, pady=(0, 1), anchor="e")
        reset_frame.grid_columnconfigure(0, weight=1)
        reset_frame.grid_columnconfigure(1, weight=0)
        try:
            python_icon_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__)))), 'icons', 'reset.png')
            python_image = Image.open(python_icon_path)
            python_icon = ctk.CTkImage(light_image=python_image, dark_image=python_image, size=(30, 30))
        except Exception as e:
            logger.error(f"Erro ao carregar ícone do Python: {e}")
            python_icon = None
        reset_btn = ctk.CTkButton(
            reset_frame,
            text="",
            image=python_icon,
            width=30,
            height=30,
            fg_color="#f5f5f5",
            hover_color="#e0e0e0",
            command=self.clear_activity_form
        )
        reset_btn.grid(row=0, column=1, sticky="e", pady=(0, 2))

        # Frame para os campos com padding consistente
        fields_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        fields_frame.pack(fill="both", expand=True, padx=15, pady=0)
        
        # Configuração dos campos
        fields = [
            ("nome", "Nome"),
            ("equipe", "Equipe"),
            ("descricao", "Descrição"),
            ("atividade", "Atividade"),
            ("inicio", "Início"),
            ("fim", "Fim"),
            ("tempo_total", "Tempo Total"),
            ("status", "Status")
        ]
        
        self.activity_entries = {}
        
        for field, label in fields:
            field_frame = ctk.CTkFrame(fields_frame, fg_color="transparent")
            field_frame.pack(fill="x", pady=3)
            
            if field == "status":
                widget = ctk.CTkComboBox(
                    field_frame,
                    values=["Concluído", "Pausado", "Ativo"],
                    height=28
                )
            elif field == "equipe":
                # Se o usuário não for da equipe 1, mostrar apenas a equipe dele
                if self.user_data['equipe_id'] == 1:
                    widget = ctk.CTkComboBox(
                        field_frame,
                        values=self.get_equipes(),
                        height=28
                    )
                else:
                    # Buscar o nome da equipe do usuário
                    query = "SELECT nome FROM equipes WHERE id = %s"
                    result = self.db.execute_query(query, (self.user_data['equipe_id'],))
                    equipe_nome = result[0]['nome'] if result else ""
                    
                    widget = ctk.CTkComboBox(
                        field_frame,
                        values=[equipe_nome],
                        height=28
                    )
                    widget.set(equipe_nome)  # Define o valor inicial
            elif field == "inicio" or field == "fim":
                # Criar entry com data/hora atual e editável
                current_time = datetime.now().strftime('%d/%m/%Y %H:%M')
                widget = ctk.CTkEntry(
                    field_frame,
                    placeholder_text=label,
                    height=28
                )
                widget.insert(0, current_time)
                
                # Função para validar e calcular o tempo total
                def calculate_total_time(*args):
                    try:
                        inicio_str = self.activity_entries['inicio'].get()
                        fim_str = self.activity_entries['fim'].get()
                        
                        # Validar formato das datas
                        inicio_time = datetime.strptime(inicio_str, '%d/%m/%Y %H:%M')
                        fim_time = datetime.strptime(fim_str, '%d/%m/%Y %H:%M')
                        
                        if fim_time <= inicio_time:
                            return
                        
                        # Calcular diferença em dias e horas
                        diff = fim_time - inicio_time
                        total_days = diff.days
                        total_hours = (diff.seconds / 3600)  # Converter segundos em horas
                        
                        # Calcular horas totais (8.8h por dia útil)
                        total_work_hours = (total_days * 8.8) + total_hours
                        
                        # Converter para horas e minutos
                        hours = int(total_work_hours)
                        minutes = int((total_work_hours - hours) * 60)
                        
                        # Formatar como hhh:mm
                        time_str = f"{hours:03d}:{minutes:02d}"
                        
                        # Atualizar campo tempo_total
                        self.activity_entries['tempo_total'].delete(0, 'end')
                        self.activity_entries['tempo_total'].insert(0, time_str)
                        
                    except (ValueError, TypeError) as e:
                        logger.error(f"Erro ao calcular tempo total: {e}")
                
                # Adicionar evento para calcular tempo quando perder o foco
                widget.bind("<FocusOut>", calculate_total_time)
                
            elif field == "tempo_total":
                # Frame especial para tempo total com checkbox
                time_frame = ctk.CTkFrame(field_frame, fg_color="transparent")
                time_frame.pack(fill="x")
                time_frame.grid_columnconfigure(0, weight=0)  # Dia
                time_frame.grid_columnconfigure(1, weight=1)  # HHH:mm
                time_frame.grid_columnconfigure(2, weight=0)  # Checkbox

                # Entry para dias
                day_entry = ctk.CTkEntry(
                    time_frame,
                    placeholder_text="Dia(s)",
                    width=50,
                    height=28
                )
                day_entry.grid(row=0, column=0, sticky="ew", padx=(0, 5))
                self.activity_entries['dias'] = day_entry

                # Entry para tempo com validação
                time_entry = ctk.CTkEntry(
                    time_frame,
                    placeholder_text="HHH:mm",
                    height=28
                )
                time_entry.grid(row=0, column=1, sticky="ew", padx=(0, 5))

                # Função para validar entrada de tempo
                def validate_time_input(value):
                    if not value:
                        return True
                    if not all(c.isdigit() or c == ':' for c in value):
                        return False
                    if len(value) > 6:
                        return False
                    return True

                # Função para formatar o tempo enquanto digita
                def format_time_entry(event=None):
                    time_str = time_entry.get().replace(":", "")
                    if not time_str:
                        return
                    if len(time_str) > 5:
                        time_str = time_str[:5]
                    if len(time_str) >= 3:
                        time_str = time_str[:3] + ":" + time_str[3:]
                    time_entry.delete(0, "end")
                    time_entry.insert(0, time_str)

                # Função para validar formato final do tempo
                def validate_time(time_str):
                    if not time_str:
                        return True
                    try:
                        if ":" not in time_str:
                            return False
                        hour, minute = map(int, time_str.split(':'))
                        return 0 <= hour <= 999 and 0 <= minute <= 59
                    except (ValueError, TypeError):
                        return False

                # Função para atualizar tempo quando perder o foco
                def on_focus_out(event=None):
                    current = time_entry.get()
                    if not current:
                        return
                    if not validate_time(current):
                        time_entry.delete(0, 'end')
                        time_entry.insert(0, "000:00")
                    update_time()

                # Função para atualizar o campo de tempo total ao mudar o campo de dias
                def on_day_entry_change(event=None):
                    try:
                        dias_str = day_entry.get()
                        dias = int(dias_str) if dias_str else 0
                        # Calcula o total de minutos apenas dos dias
                        total_minutes = dias * 528  # 8h48m = 528 min
                        # Atualiza o campo HHH:mm
                        new_hours = total_minutes // 60
                        new_minutes = total_minutes % 60
                        time_entry.delete(0, 'end')
                        time_entry.insert(0, f"{new_hours:03d}:{new_minutes:02d}")
                        update_time()  # Atualiza datas
                    except Exception as e:
                        logger.error(f"Erro ao atualizar tempo pelo campo de dias: {e}")

                # Função para atualizar o campo de dias ao mudar o campo de tempo
                def on_time_entry_change(event=None):
                    try:
                        time_str = time_entry.get()
                        if not time_str or ":" not in time_str:
                            day_entry.delete(0, 'end')
                            return
                        hours, minutes = map(int, time_str.split(":"))
                        total_minutes = hours * 60 + minutes
                        dias = total_minutes // 528
                        day_entry.delete(0, 'end')
                        if dias > 0:
                            day_entry.insert(0, str(dias))
                        else:
                            day_entry.insert(0, "")
                        update_time()  # Atualiza datas
                    except Exception as e:
                        logger.error(f"Erro ao atualizar dias pelo campo de tempo: {e}")

                # Registrar validação
                validate_cmd = (self.register(validate_time_input), '%P')
                time_entry.configure(validate="key", validatecommand=validate_cmd)

                # Adicionar eventos
                time_entry.bind("<KeyRelease>", format_time_entry)
                time_entry.bind("<FocusOut>", on_focus_out)
                time_entry.bind("<Return>", on_time_entry_change)
                day_entry.bind("<Return>", on_day_entry_change)

                # Checkbox para adicionar/subtrair tempo
                self.add_time_var = tk.BooleanVar(value=True)
                add_time_cb = ctk.CTkCheckBox(
                    time_frame,
                    text="Adicionar",
                    variable=self.add_time_var,
                    height=28,
                    width=20
                )
                add_time_cb.grid(row=0, column=2, sticky="e")

                # Função para atualizar tempo
                def update_time(*args):
                    try:
                        time_str = time_entry.get()
                        if not time_str or time_str == "000:00":
                            return
                        if not validate_time(time_str):
                            return
                        hours, minutes = map(int, time_str.split(':'))
                        time_delta = timedelta(hours=hours, minutes=minutes)
                        inicio_str = self.activity_entries['inicio'].get()
                        fim_str = self.activity_entries['fim'].get()
                        inicio_time = datetime.strptime(inicio_str, '%d/%m/%Y %H:%M')
                        fim_time = datetime.strptime(fim_str, '%d/%m/%Y %H:%M')
                        if self.add_time_var.get():
                            new_fim = inicio_time + time_delta
                            self.activity_entries['fim'].delete(0, 'end')
                            self.activity_entries['fim'].insert(0, new_fim.strftime('%d/%m/%Y %H:%M'))
                        else:
                            new_inicio = fim_time - time_delta
                            self.activity_entries['inicio'].delete(0, 'end')
                            self.activity_entries['inicio'].insert(0, new_inicio.strftime('%d/%m/%Y %H:%M'))
                    except Exception as e:
                        logger.error(f"Erro ao atualizar tempo: {e}")

                def on_checkbox_change():
                    self.after(100, update_time)
                add_time_cb.configure(command=on_checkbox_change)

                widget = time_entry
            else:
                widget = ctk.CTkEntry(
                    field_frame,
                    placeholder_text=label,
                    height=28
                )
            
            if field != "tempo_total":  # Para tempo_total já empacotamos acima
                widget.pack(fill="x")
            self.activity_entries[field] = widget
        
        # Frame para botões
        btn_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        btn_frame.pack(fill="x", padx=15, pady=(10, 5))
        
        # Configurar grid para 3 colunas
        btn_frame.grid_columnconfigure(0, weight=1)
        btn_frame.grid_columnconfigure(1, weight=1)
        btn_frame.grid_columnconfigure(2, weight=1)
        
        # Carregar ícone do Excel
        if hasattr(sys, "_MEIPASS"):
            icons_dir = os.path.join(sys._MEIPASS, 'icons', 'excel.png')
        else:
            icons_dir = os.path.join(os.path.abspath("."), 'icons', 'excel.png')
        
        try:
            excel_image = Image.open(icons_dir)
            self.excel_icon = ctk.CTkImage(light_image=excel_image, dark_image=excel_image, size=(20, 20))
        except Exception as e:
            logger.error(f"Erro ao carregar ícone do Excel: {e}")
            self.excel_icon = None
        
        # Botão do Excel
        excel_btn = ctk.CTkButton(
            btn_frame,
            text="",
            image=self.excel_icon,
            width=32,
            height=28,
            command=self.open_excel_selector,
            fg_color="#FF5722",
            hover_color="#CE461B"
        )
        excel_btn.grid(row=0, column=0, padx=2, pady=2, sticky="ew")
        
        # Botões
        save_btn = ctk.CTkButton(
            btn_frame,
            text="Salvar",
            height=28,
            command=self.save_activity,  # Atualizado para chamar o novo método
            fg_color="#ff5722",
            hover_color="#ce461b"
        )
        save_btn.grid(row=0, column=1, padx=2, pady=2, sticky="ew")
        
        save_all_btn = ctk.CTkButton(
            btn_frame,
            text="Salvar a Todos",
            height=28,
            command=self.save_activity_to_all,  # Atualizado para chamar o novo método
            fg_color="#ff5722",
            hover_color="#ce461b"
        )
        save_all_btn.grid(row=0, column=2, padx=2, pady=2, sticky="ew")

        # Frame para botões
        btn_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        btn_frame.pack(fill="x", padx=15, pady=(10, 5))

        # Adicionar botão de reset acima dos botões, à direita
        reset_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        reset_frame.pack(fill="x", padx=15, pady=(0, 0))
        reset_frame.grid_columnconfigure(0, weight=1)
        reset_frame.grid_columnconfigure(1, weight=0)

        # Carregar ícone do Python (usar python.png na pasta icons, senão fallback)
        try:
            python_icon_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__)))), 'icons', 'python.png')
            if not os.path.exists(python_icon_path):
                python_icon_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__)))), 'icons', 'logo_login_light.png')
            python_image = Image.open(python_icon_path)
            python_icon = ctk.CTkImage(light_image=python_image, dark_image=python_image, size=(20, 20))
        except Exception as e:
            logger.error(f"Erro ao carregar ícone do Python: {e}")
            python_icon = None

        reset_btn = ctk.CTkButton(
            reset_frame,
            text="",
            image=python_icon,
            width=28,
            height=28,
            fg_color="#f5f5f5",
            hover_color="#e0e0e0",
            command=self.clear_activity_form
        )
        reset_btn.grid(row=0, column=1, sticky="e", pady=(0, 2))

    # Métodos auxiliares para cada aba
    def add_team(self):
        """Adiciona uma nova equipe e atualiza o ComboBox"""
        team_name = self.team_entry.get().strip()
        if not team_name:
            messagebox.showwarning("Aviso", "Digite um nome para a equipe!")
            return
                
        try:
            query = "INSERT INTO equipes (nome) VALUES (%s)"
            self.db.execute_query(query, (team_name,))
            messagebox.showinfo("Sucesso", "Equipe adicionada com sucesso!")
            self.team_entry.delete(0, 'end')
            self.load_teams()
            
            # Atualiza o ComboBox na aba de usuários
            self.update_equipes_combobox()
            
        except Exception as e:
            logger.error(f"Erro ao adicionar equipe: {e}")
            messagebox.showerror("Erro", "Erro ao adicionar equipe")

    def load_teams(self):
        """Carrega lista de equipes"""
        for item in self.teams_tree.get_children():
            self.teams_tree.delete(item)
            
        try:
            query = """
                SELECT e.id, e.nome, COUNT(u.id) as membros, e.created_at
                FROM equipes e
                LEFT JOIN usuarios u ON e.id = u.equipe_id AND u.status = TRUE
                GROUP BY e.id, e.nome, e.created_at
                ORDER BY e.nome
            """
            teams = self.db.execute_query(query)
            
            # Verificar se teams não é None e tem resultados
            if teams:
                for team in teams:
                    self.teams_tree.insert("", "end", values=(
                        team['id'],
                        team['nome'],
                        team['membros'],
                        team['created_at'].strftime('%d/%m/%Y %H:%M')
                    ))
            else:
                logger.info("Nenhuma equipe encontrada no banco de dados")
            
        except Exception as e:
            logger.error(f"Erro ao carregar equipes: {e}")
            messagebox.showerror("Erro", "Erro ao carregar equipes")

    def load_blocks(self):
        """Carrega lista de bloqueios sem filtrar pela equipe do usuário logado"""
        for item in self.blocks_tree.get_children():
            self.blocks_tree.delete(item)
            
        try:
            # Query agora não filtra pela equipe do usuário logado
            query = """
                SELECT ul.id, u.nome, e.nome as equipe, ul.lock_status, ul.unlock_control
                FROM user_lock_unlock ul
                JOIN usuarios u ON ul.user_id = u.id
                JOIN equipes e ON u.equipe_id = e.id
                WHERE u.status = TRUE
                ORDER BY u.nome
            """
            blocks = self.db.execute_query(query)
            
            if blocks:
                for block in blocks:
                    status = "Bloqueado" if block['lock_status'] else "Desbloqueado"
                    controle = "Liberado" if block['unlock_control'] else "Bloqueado"
                    
                    self.blocks_tree.insert("", "end", values=(
                        block['id'],
                        block['nome'],
                        block['equipe'],
                        status,
                        controle
                    ))
                    
                # Atualizar indicadores para todos os usuários
                self.update_status_indicators()
                
        except Exception as e:
            logger.error(f"Erro ao carregar bloqueios: {e}")
            messagebox.showerror("Erro", "Erro ao carregar lista de bloqueios")

    def update_status_indicators(self):
        """Atualiza os indicadores de status apenas para a equipe atual"""
        try:
            # Usuários ativos da equipe atual
            query_active = """
                SELECT COUNT(*) as count 
                FROM usuarios 
                WHERE is_logged_in = TRUE
                AND equipe_id = %s
            """
            active_users = self.db.execute_query(query_active, (self.user_data['equipe_id'],))[0]['count']
            self.status_indicators["Usuários Ativos"].configure(text=str(active_users))
            
            # Usuários bloqueados da equipe atual
            query_blocked = """
                SELECT COUNT(*) as count 
                FROM user_lock_unlock ul
                JOIN usuarios u ON ul.user_id = u.id
                WHERE ul.unlock_control = FALSE
                AND u.equipe_id = %s
            """
            blocked_users = self.db.execute_query(query_blocked, (self.user_data['equipe_id'],))[0]['count']
            self.status_indicators["Usuários Bloqueados"].configure(text=str(blocked_users))
            
            # Tentativas de login da equipe atual
            self.status_indicators["Tentativas de Login"].configure(text="0")
            
        except Exception as e:
            logger.error(f"Erro ao atualizar indicadores: {e}")
            messagebox.showerror("Erro", "Erro ao atualizar indicadores")

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
        if self.user_data['equipe_id'] == 1:
            self.entry_widgets['equipe'].set(self.get_equipes()[0] if self.get_equipes() else '')
        else:
            # Se não for da equipe 1, manter a equipe atual
            query = "SELECT nome FROM equipes WHERE id = %s"
            result = self.db.execute_query(query, (self.user_data['equipe_id'],))
            equipe_nome = result[0]['nome'] if result else ''
            self.entry_widgets['equipe'].set(equipe_nome)
            
        self.entry_widgets['tipo_usuario'].set('comum')
        
        # Resetar data para atual
        self.entry_widgets['data_entrada'].configure(state="normal")
        self.entry_widgets['data_entrada'].delete(0, 'end')
        self.entry_widgets['data_entrada'].insert(0, datetime.now().strftime('%d/%m/%Y'))
        self.entry_widgets['data_entrada'].configure(state="readonly")
        
    def get_equipes(self):
        """Retorna lista de todas as equipes do banco"""
        try:
            # Query simplificada para trazer todas as equipes
            query = """
                SELECT nome 
                FROM equipes 
                ORDER BY nome
            """
            result = self.db.execute_query(query)
            
            # Debug para verificar o resultado
            logger.debug(f"Equipes encontradas: {result}")
            
            return [row['nome'] for row in result] if result else []
        except Exception as e:
            logger.error(f"Erro ao buscar equipes: {e}")
            return []

    def update_equipes_combobox(self):
        """Atualiza o ComboBox de equipes com dados do banco"""
        try:
            equipes = self.get_equipes()
            if 'equipe' in self.entry_widgets:
                self.entry_widgets['equipe'].configure(values=equipes)
                # Se o ComboBox estiver vazio, seleciona a primeira equipe
                if not self.entry_widgets['equipe'].get() and equipes:
                    self.entry_widgets['equipe'].set(equipes[0])
        except Exception as e:
            logger.error(f"Erro ao atualizar ComboBox de equipes: {e}")

    def show_team_context_menu(self, event):
        """Exibe menu de contexto para equipes"""
        try:
            # Identificar item clicado
            item = self.teams_tree.identify_row(event.y)
            if not item:
                return
            
            # Selecionar o item
            self.teams_tree.selection_set(item)
            
            # Criar menu
            menu = tk.Menu(self, tearoff=0)
            menu.add_command(label="Excluir Equipe", 
                            command=lambda: self.delete_team(self.teams_tree.item(item)['values'][0]))
            
            # Exibir menu
            menu.post(event.x_root, event.y_root)
            
        except Exception as e:
            logger.error(f"Erro ao exibir menu de contexto: {e}")

    def delete_team(self, team_id):
        """Exclui uma equipe"""
        try:
            # Verificar se há usuários ativos na equipe
            check_query = """
                SELECT COUNT(*) as count 
                FROM usuarios 
                WHERE equipe_id = %s AND status = TRUE
            """
            result = self.db.execute_query(check_query, (team_id,))
            
            if result[0]['count'] > 0:
                messagebox.showerror(
                    "Erro",
                    "Não é possível excluir esta equipe pois existem usuários ativos vinculados a ela."
                )
                return
            
            # Confirmar exclusão
            if messagebox.askyesno("Confirmar", "Deseja realmente excluir esta equipe?"):
                delete_query = "DELETE FROM equipes WHERE id = %s"
                cursor = self.db.connection.cursor(dictionary=True)
                cursor.execute(delete_query, (team_id,))
                self.db.connection.commit()
                cursor.close()
                
                messagebox.showinfo("Sucesso", "Equipe excluída com sucesso!")
                self.load_teams()
                self.update_counters()  # Atualizar contadores
                
        except Exception as e:
            logger.error(f"Erro ao excluir equipe: {e}")
            messagebox.showerror("Erro", "Erro ao excluir equipe")

    def delete_selected_team(self):
        """Exclui a equipe selecionada"""
        selected_items = self.teams_tree.selection()
        if not selected_items:
            messagebox.showwarning("Aviso", "Selecione uma equipe para excluir!")
            return
        
        team_id = self.teams_tree.item(selected_items[0])['values'][0]
        self.delete_team(team_id)

    def toggle_user_lock(self):
        """Altera o status de bloqueio do usuário selecionado"""
        selected_items = self.blocks_tree.selection()
        if not selected_items:
            messagebox.showwarning("Aviso", "Selecione um usuário para alterar o status!")
            return
        
        try:
            block_id = self.blocks_tree.item(selected_items[0])['values'][0]
            
            # Buscar status atual
            status_query = "SELECT unlock_control FROM user_lock_unlock WHERE id = %s"
            result = self.db.execute_query(status_query, (block_id,))
            
            if result:
                current_status = result[0]['unlock_control']
                new_status = not current_status
                
                # Atualizar status
                update_query = """
                    UPDATE user_lock_unlock 
                    SET unlock_control = %s,
                        updated_at = CURRENT_TIMESTAMP
                    WHERE id = %s
                """
                self.db.execute_query(update_query, (new_status, block_id))
                
                status_text = "desbloqueado" if new_status else "bloqueado"
                messagebox.showinfo("Sucesso", f"Usuário {status_text} com sucesso!")
                
                # Recarregar lista
                self.load_blocks()
                self.update_counters()
                
        except Exception as e:
            logger.error(f"Erro ao alterar status de bloqueio: {e}")
            messagebox.showerror("Erro", "Erro ao alterar status de bloqueio")

    def search_blocks(self):
        """Pesquisa usuários bloqueados com base no texto inserido"""
        search_text = self.blocks_search_entry.get().strip()
        
        # Limpar a árvore
        for item in self.blocks_tree.get_children():
            self.blocks_tree.delete(item)
        
        if not search_text:
            self.load_blocks()
            return
            
        try:
            # Query de busca agora não filtra pela equipe do usuário logado
            query = """
                SELECT ul.id, u.nome, e.nome as equipe, ul.lock_status, ul.unlock_control
                FROM user_lock_unlock ul
                JOIN usuarios u ON ul.user_id = u.id
                JOIN equipes e ON u.equipe_id = e.id
                WHERE u.status = TRUE
                AND u.nome LIKE %s
                ORDER BY u.nome
            """
            blocks = self.db.execute_query(query, (f"%{search_text}%",))
            
            if blocks:
                for block in blocks:
                    status = "Bloqueado" if block['lock_status'] else "Desbloqueado"
                    controle = "Liberado" if block['unlock_control'] else "Bloqueado"
                    
                    self.blocks_tree.insert("", "end", values=(
                        block['id'],
                        block['nome'],
                        block['equipe'],
                        status,
                        controle
                    ))
                    
        except Exception as e:
            logger.error(f"Erro ao buscar bloqueios: {str(e)}")
            messagebox.showerror(
                "Erro",
                "Ocorreu um erro ao buscar os bloqueios. Por favor, tente novamente."
            )

    def search_teams(self):
        """Pesquisa equipes com base no nome"""
        search_text = self.teams_search_entry.get().strip()
        
        # Limpar a árvore
        for item in self.teams_tree.get_children():
            self.teams_tree.delete(item)
        
        if not search_text:
            self.load_teams()
            return
        
        try:
            query = """
                SELECT e.id, e.nome, COUNT(u.id) as membros, e.created_at
                FROM equipes e
                LEFT JOIN usuarios u ON e.id = u.equipe_id AND u.status = TRUE
                WHERE LOWER(e.nome) LIKE LOWER(%s)
                GROUP BY e.id, e.nome, e.created_at
                ORDER BY e.nome
            """
            
            search_pattern = f"%{search_text}%"
            teams = self.db.execute_query(query, (search_pattern,))
            
            if teams:
                for team in teams:
                    self.teams_tree.insert(
                        "",
                        "end",
                        values=(
                            team['id'],
                            team['nome'],
                            team['membros'],
                            team['created_at'].strftime('%d/%m/%Y %H:%M')
                        )
                    )
                    
        except Exception as e:
            logger.error(f"Erro ao buscar equipes: {str(e)}")
            messagebox.showerror(
                "Erro",
                "Ocorreu um erro ao buscar as equipes. Por favor, tente novamente."
            )

    def show_block_context_menu(self, event):
        """Exibe menu de contexto para bloqueios"""
        try:
            # Identificar item clicado
            item = self.blocks_tree.identify_row(event.y)
            if not item:
                return
            
            # Selecionar o item
            self.blocks_tree.selection_set(item)
            
            # Criar menu
            menu = tk.Menu(self, tearoff=0)
            menu.add_command(label="Alterar Status", 
                            command=lambda: self.toggle_user_lock())
            
            # Exibir menu
            menu.post(event.x_root, event.y_root)
            
        except Exception as e:
            logger.error(f"Erro ao exibir menu de contexto: {e}")

    def update_counters(self):
        """Atualiza os contadores de usuários, equipes e bloqueios (baseado nos itens carregados no Treeview)"""
        try:
            # Verificar se os widgets existem antes de atualizar
            if not (hasattr(self, 'tree') and hasattr(self, 'user_count_label') and
                    hasattr(self, 'teams_tree') and hasattr(self, 'team_count_label') and
                    hasattr(self, 'blocks_tree') and hasattr(self, 'block_stats_label')):
                return

            # Contar usuários carregados na Treeview
            user_count = len(self.tree.get_children())
            self.user_count_label.configure(text=f"Total: {user_count} usuário{'s' if user_count != 1 else ''}")
            
            # Contar equipes carregadas na Treeview
            team_count = len(self.teams_tree.get_children())
            self.team_count_label.configure(text=f"Total: {team_count} equipe{'s' if team_count != 1 else ''}")
            
            # Contar bloqueios carregados na Treeview
            blocked = 0
            unblocked = 0
            for item in self.blocks_tree.get_children():
                values = self.blocks_tree.item(item, 'values')
                if len(values) >= 5:
                    if values[4] == 'Bloqueado':
                        blocked += 1
                    elif values[4] == 'Liberado':
                        unblocked += 1
            self.block_stats_label.configure(
                text=f"Bloqueados: {blocked} | Liberados: {unblocked}"
            )
        except Exception as e:
            logger.error(f"Erro ao atualizar contadores: {e}")
            messagebox.showerror("Erro", "Erro ao atualizar contadores")

    def open_excel_selector(self):
        """Abre o seletor de Excel para escolher a descrição"""
        def on_select(description):
            if description and "descricao" in self.activity_entries:
                self.activity_entries["descricao"].delete(0, "end")
                self.activity_entries["descricao"].insert(0, description)
        
        ExcelSelector(self, on_select)

    def load_activities(self):
        """Carrega todas as atividades na treeview"""
        # Limpar a árvore
        for item in self.activities_tree.get_children():
            self.activities_tree.delete(item)
            
        try:
            # Query base que junta atividades com usuários (sem filtro de equipe)
            query = """
                SELECT a.*, u.nome as user_name
                FROM atividades a
                JOIN usuarios u ON a.user_id = u.id
                ORDER BY a.id DESC
            """
            
            activities = self.db.execute_query(query)
            
            if activities:
                for activity in activities:
                    # Determinar o status baseado nas colunas booleanas
                    status = "Concluído" if activity['concluido'] else \
                            "Pausado" if activity['pausado'] else \
                            "Ativo" if activity['ativo'] else "Indefinido"
                    
                    self.activities_tree.insert("", "end", values=(
                        activity['user_name'],
                        activity['description'],
                        activity['atividade'],
                        activity['total_time'],
                        status
                    ))
                
                # Atualizar contador de atividades
                self.activities_count_label.configure(
                    text=f"Total: {len(activities)} atividades"
                )
            else:
                self.activities_count_label.configure(text="Total: 0 atividades")
                
        except Exception as e:
            logger.error(f"Erro ao carregar atividades: {e}")
            messagebox.showerror("Erro", "Erro ao carregar lista de atividades")
    
    def search_activities(self):
        """Pesquisa atividades com base no termo de busca"""
        search_term = self.activities_search_entry.get().strip()
        
        # Limpar a árvore
        for item in self.activities_tree.get_children():
            self.activities_tree.delete(item)
        
        if not search_term:
            self.load_activities()
            return
        
        try:
            # Query que busca por correspondência em vários campos (sem filtro de equipe)
            query = """
                SELECT a.*, u.nome as user_name
                FROM atividades a
                JOIN usuarios u ON a.user_id = u.id
                WHERE (
                    u.nome LIKE %s OR
                    a.description LIKE %s OR
                    a.atividade LIKE %s
                )
                ORDER BY a.id DESC
            """
            
            search_pattern = f"%{search_term}%"
            activities = self.db.execute_query(
                query, 
                (search_pattern, search_pattern, search_pattern)
            )
            
            if activities:
                for activity in activities:
                    # Determinar o status baseado nas colunas booleanas
                    status = "Concluído" if activity['concluido'] else \
                            "Pausado" if activity['pausado'] else \
                            "Ativo" if activity['ativo'] else "Indefinido"
                    
                    self.activities_tree.insert("", "end", values=(
                        activity['user_name'],
                        activity['description'],
                        activity['atividade'],
                        activity['total_time'],
                        status
                    ))
                
                # Atualizar contador com resultados da busca
                self.activities_count_label.configure(
                    text=f"Encontrado(s): {len(activities)} atividade(s)"
                )
            else:
                self.activities_count_label.configure(text="Nenhuma atividade encontrada")
                
        except Exception as e:
            logger.error(f"Erro ao pesquisar atividades: {e}")
            messagebox.showerror("Erro", "Erro ao pesquisar atividades")

    def on_activity_double_click(self, event):
        """Manipula o evento de duplo clique na TreeView de atividades"""
        # Obter o item selecionado
        selected_item = self.activities_tree.selection()
        if not selected_item:
            return
            
        try:
            # Obter os valores do item selecionado
            values = self.activities_tree.item(selected_item)['values']
            if values:
                # O nome está na primeira coluna (índice 0)
                nome = values[0]
                
                # Atualizar o campo 'nome' no formulário
                if 'nome' in self.activity_entries:
                    self.activity_entries['nome'].delete(0, 'end')
                    self.activity_entries['nome'].insert(0, nome)
                    
        except Exception as e:
            logger.error(f"Erro ao processar duplo clique: {e}")
            messagebox.showerror("Erro", "Erro ao selecionar usuário")

    def save_activity(self):
        """Salva uma atividade no banco de dados"""
        try:
            # Coletar dados do formulário
            data = {field: widget.get() for field, widget in self.activity_entries.items()}
            
            # Validar campos obrigatórios
            required_fields = ['nome', 'equipe', 'descricao', 'atividade', 'inicio', 'fim', 'tempo_total', 'status']
            empty_fields = [field for field in required_fields if not data[field].strip()]
            
            if empty_fields:
                messagebox.showerror(
                    "Erro",
                    f"Os seguintes campos são obrigatórios: {', '.join(empty_fields)}"
                )
                return
            
            # Verificar se o usuário existe
            user_query = """
                SELECT id, equipe_id 
                FROM usuarios 
                WHERE nome = %s AND status = TRUE
            """
            user_result = self.db.execute_query(user_query, (data['nome'],))
            
            if not user_result:
                messagebox.showerror(
                    "Erro",
                    "Usuário não encontrado ou inativo. Verifique o nome do usuário."
                )
                return
            
            user_id = user_result[0]['id']
            user_equipe_id = user_result[0]['equipe_id']
            
            # Verificar se o usuário pertence à equipe selecionada
            equipe_query = "SELECT id FROM equipes WHERE nome = %s"
            equipe_result = self.db.execute_query(equipe_query, (data['equipe'],))
            
            if not equipe_result or equipe_result[0]['id'] != user_equipe_id:
                messagebox.showerror(
                    "Erro",
                    "O usuário não pertence à equipe selecionada."
                )
                return
            
            # Converter datas para o formato do banco
            try:
                start_time = datetime.strptime(data['inicio'], '%d/%m/%Y %H:%M')
                end_time = datetime.strptime(data['fim'], '%d/%m/%Y %H:%M')
                
                if end_time <= start_time:
                    messagebox.showerror(
                        "Erro",
                        "A data/hora de fim deve ser posterior à data/hora de início."
                    )
                    return
                
            except ValueError:
                messagebox.showerror(
                    "Erro",
                    "Formato de data/hora inválido. Use dd/mm/aaaa HH:MM"
                )
                return
            
            # Determinar status da atividade
            status = data['status'].lower()
            ativo = status == "ativo"
            pausado = status == "pausado"
            concluido = status == "concluído"
            
            # Preparar query de inserção
            insert_query = """
                INSERT INTO atividades (
                    user_id, description, atividade, start_time,
                    end_time, total_time, ativo, pausado, concluido
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
            """
            
            params = (
                user_id,
                data['descricao'],
                data['atividade'],
                start_time,
                end_time,
                data['tempo_total'],
                ativo,
                pausado,
                concluido
            )
            
            # Executar inserção
            cursor = self.db.connection.cursor(dictionary=True)
            cursor.execute(insert_query, params)
            self.db.connection.commit()
            cursor.close()
            
            messagebox.showinfo("Sucesso", "Atividade salva com sucesso!")
            
            # Recarregar lista de atividades
            self.load_activities()
            
            # Limpar formulário
            self.clear_activity_form()
            
        except Exception as e:
            logger.error(f"Erro ao salvar atividade: {e}")
            messagebox.showerror("Erro", f"Erro ao salvar atividade: {e}")
    
    def clear_activity_form(self):
        """Limpa e restaura o formulário de atividade para o estado inicial"""
        # Limpar todos os campos
        for field in self.activity_entries:
            widget = self.activity_entries[field]
            if hasattr(widget, 'set'):
                widget.set("")
            else:
                widget.delete(0, 'end')

        # Restaurar valores iniciais
        # Nome: vazio (placeholder visível)
        self.activity_entries['nome'].delete(0, 'end')
        # Equipe: restaurar equipe padrão do usuário
        if self.user_data['equipe_id'] == 1:
            equipes = self.get_equipes()
            if equipes:
                self.activity_entries['equipe'].set(equipes[0])
            else:
                self.activity_entries['equipe'].set("")
        else:
            query = "SELECT nome FROM equipes WHERE id = %s"
            result = self.db.execute_query(query, (self.user_data['equipe_id'],))
            equipe_nome = result[0]['nome'] if result else ''
            self.activity_entries['equipe'].set(equipe_nome)
        # Descrição: vazio (placeholder visível)
        self.activity_entries['descricao'].delete(0, 'end')
        # Atividade: vazio (placeholder visível)
        self.activity_entries['atividade'].delete(0, 'end')
        # Início e Fim: data/hora atual
        current_time = datetime.now().strftime('%d/%m/%Y %H:%M')
        self.activity_entries['inicio'].delete(0, 'end')
        self.activity_entries['inicio'].insert(0, current_time)
        self.activity_entries['fim'].delete(0, 'end')
        self.activity_entries['fim'].insert(0, current_time)
        # Tempo total: vazio (placeholder visível)
        self.activity_entries['tempo_total'].delete(0, 'end')
        # Dias: vazio (placeholder visível)
        if 'dias' in self.activity_entries:
            self.activity_entries['dias'].delete(0, 'end')
        # Status: Concluído (ou Ativo, conforme desejado)
        self.activity_entries['status'].set('Concluído')

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
        
        # Carregar ícone do Excel
        excel_image = ctk.CTkImage(
            light_image=Image.open(os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__)))), 'icons', 'excel.png')),
            dark_image=Image.open(os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__)))), 'icons', 'excel.png')),
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
            from .Excel_Viewer import ExcelViewer
            
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

    def save_activity_to_all(self):
        """Salva uma atividade para todos os usuários da equipe selecionada"""
        try:
            # Coletar dados do formulário
            data = {field: widget.get() for field, widget in self.activity_entries.items()}
            
            # Validar campos obrigatórios (exceto 'nome')
            required_fields = ['equipe', 'descricao', 'atividade', 'inicio', 'fim', 'tempo_total', 'status']
            empty_fields = [field for field in required_fields if not data[field].strip()]
            
            if empty_fields:
                messagebox.showerror(
                    "Erro",
                    f"Os seguintes campos são obrigatórios: {', '.join(empty_fields)}"
                )
                return
            
            # Buscar ID da equipe
            equipe_query = "SELECT id FROM equipes WHERE nome = %s"
            equipe_result = self.db.execute_query(equipe_query, (data['equipe'],))
            
            if not equipe_result:
                messagebox.showerror("Erro", "Equipe não encontrada!")
                return
            
            equipe_id = equipe_result[0]['id']
            
            # Buscar todos os usuários ativos da equipe
            users_query = """
                SELECT id 
                FROM usuarios 
                WHERE equipe_id = %s AND status = TRUE
            """
            users = self.db.execute_query(users_query, (equipe_id,))
            
            if not users:
                messagebox.showerror(
                    "Erro",
                    "Não foram encontrados usuários ativos nesta equipe."
                )
                return
            
            # Converter datas para o formato do banco
            try:
                start_time = datetime.strptime(data['inicio'], '%d/%m/%Y %H:%M')
                end_time = datetime.strptime(data['fim'], '%d/%m/%Y %H:%M')
                
                if end_time <= start_time:
                    messagebox.showerror(
                        "Erro",
                        "A data/hora de fim deve ser posterior à data/hora de início."
                    )
                    return
                
            except ValueError:
                messagebox.showerror(
                    "Erro",
                    "Formato de data/hora inválido. Use dd/mm/aaaa HH:MM"
                )
                return
            
            # Determinar status da atividade
            status = data['status'].lower()
            ativo = status == "ativo"
            pausado = status == "pausado"
            concluido = status == "concluído"
            
            # Preparar query de inserção
            insert_query = """
                INSERT INTO atividades (
                    user_id, description, atividade, start_time,
                    end_time, total_time, ativo, pausado, concluido
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
            """
            
            # Inserir atividade para cada usuário
            cursor = self.db.connection.cursor(dictionary=True)
            inserted_count = 0
            
            for user in users:
                try:
                    params = (
                        user['id'],
                        data['descricao'],
                        data['atividade'],
                        start_time,
                        end_time,
                        data['tempo_total'],
                        ativo,
                        pausado,
                        concluido
                    )
                    cursor.execute(insert_query, params)
                    inserted_count += 1
                except Exception as e:
                    logger.error(f"Erro ao inserir atividade para usuário {user['id']}: {e}")
                    continue
            
            self.db.connection.commit()
            cursor.close()
            
            if inserted_count > 0:
                messagebox.showinfo(
                    "Sucesso",
                    f"Atividade salva com sucesso para {inserted_count} usuário(s)!"
                )
                
                # Recarregar lista de atividades
                self.load_activities()
                
                # Limpar formulário
                self.clear_activity_form()
            else:
                messagebox.showerror(
                    "Erro",
                    "Não foi possível salvar a atividade para nenhum usuário."
                )
            
        except Exception as e:
            logger.error(f"Erro ao salvar atividade para todos: {e}")
            messagebox.showerror("Erro", f"Erro ao salvar atividade: {e}")

    def add_treeview_sorting(self, tree):
        """Adiciona ordenação ao clicar nos cabeçalhos das colunas da Treeview"""
        def treeview_sort_column(tv, col, reverse):
            try:
                l = [(tv.set(k, col), k) for k in tv.get_children('')]
                # Tentar converter para número se possível
                try:
                    l.sort(key=lambda t: float(t[0].replace(',', '.')), reverse=reverse)
                except ValueError:
                    l.sort(key=lambda t: t[0].lower() if isinstance(t[0], str) else t[0], reverse=reverse)
                for index, (val, k) in enumerate(l):
                    tv.move(k, '', index)
                # Alternar ordem na próxima vez
                tv.heading(col, command=lambda: treeview_sort_column(tv, col, not reverse))
            except Exception as e:
                logger.error(f"Erro ao ordenar coluna '{col}': {e}")

        for col in tree["columns"]:
            tree.heading(col, command=lambda _col=col: treeview_sort_column(tree, _col, False))

        # Após criar cada Treeview, adicionar sorting:
        self.add_treeview_sorting(self.tree)
        self.add_treeview_sorting(self.teams_tree)
        self.add_treeview_sorting(self.blocks_tree)
        self.add_treeview_sorting(self.activities_tree)
        self.add_treeview_sorting(self.sheets_tree)

