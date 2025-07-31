from customtkinter import CTkFrame
import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, messagebox
import logging
from ....database.connection import DatabaseConnection

logger = logging.getLogger(__name__)

class BlocksTab(CTkFrame):
    def __init__(self, parent, user_data):
        super().__init__(parent)
        self.user_data = user_data
        self.db = DatabaseConnection()
        self.setup_blocks()
        self.pack(expand=True, fill="both")

    def setup_blocks(self):
        """Configura a aba de bloqueios"""
        # Título com estatísticas
        title_frame = ctk.CTkFrame(self, fg_color="transparent")
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

        blocks_frame = ctk.CTkFrame(self, fg_color="transparent")
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
        self.update_counter()

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
                self.update_counter()
                
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

    def update_counter(self):
        """Atualiza o contador de bloqueios na aba."""
        try:
            if hasattr(self, 'blocks_tree') and hasattr(self, 'block_stats_label'):
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
            logger.error(f"Erro ao atualizar contador de bloqueios: {e}")