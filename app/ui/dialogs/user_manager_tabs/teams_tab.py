from customtkinter import CTkFrame
import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import logging
from ....database.connection import DatabaseConnection

logger = logging.getLogger(__name__)

class TeamsTab(CTkFrame):
    def __init__(self, parent, user_data, manager=None):
        super().__init__(parent)
        self.user_data = user_data
        self.manager = manager
        self.db = DatabaseConnection()
        self.setup_teams()
        self.pack(expand=True, fill="both")

    def setup_teams(self):
        """Configura a aba de equipes"""
        # Título com estatísticas
        title_frame = ctk.CTkFrame(self, fg_color="transparent")
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

        teams_frame = ctk.CTkFrame(self, fg_color="transparent")
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
        self.update_counter()

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
                self.update_counter()  # Atualizar contadores
                
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

    def update_equipes_combobox(self):
        """Atualiza o ComboBox de equipes com dados do banco"""
        if 'equipe' in self.entry_widgets:
            equipes = self.manager.get_equipes() if self.manager else []
            self.entry_widgets['equipe'].configure(values=equipes)
            # Se o ComboBox estiver vazio, seleciona a primeira equipe
            if not self.entry_widgets['equipe'].get() and equipes:
                self.entry_widgets['equipe'].set(equipes[0])

    def update_counter(self):
        """Atualiza o contador de equipes na aba."""
        try:
            if hasattr(self, 'teams_tree') and hasattr(self, 'team_count_label'):
                team_count = len(self.teams_tree.get_children())
                self.team_count_label.configure(text=f"Total: {team_count} equipe{'s' if team_count != 1 else ''}")
        except Exception as e:
            logger.error(f"Erro ao atualizar contador de equipes: {e}")