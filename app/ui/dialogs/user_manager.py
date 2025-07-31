import customtkinter as ctk
from tkinter import messagebox
import logging
from ...database.connection import DatabaseConnection
from ..dialogs.user_manager_tabs.users_tab import UsersTab
from ..dialogs.user_manager_tabs.teams_tab import TeamsTab
from ..dialogs.user_manager_tabs.activities_tab import ActivitiesTab
from ..dialogs.user_manager_tabs.blocks_tab import BlocksTab
from ..dialogs.user_manager_tabs.sheets_tab import SheetsTab

logger = logging.getLogger(__name__)

class UserManager(ctk.CTkFrame):
    def __init__(self, parent, user_data):
        super().__init__(parent)
        
        # Armazenar dados do usuário logado
        self.user_data = user_data
        self.parent = parent

        self.db = DatabaseConnection()
        
        # Configurar pesos do grid para responsividade
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Dicionários para armazenar widgets
        self.entry_widgets = {}
        self.current_user_id = None
        
        self.setup_ui()  # Cria self.tab_view e as abas
        
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

        # Criar abas com nomes mais descritivos
        self.tab_users = self.tab_view.add("Gerenciar Usuários")
        self.tab_teams = self.tab_view.add("Gerenciar Equipes")
        self.tab_blocks = self.tab_view.add("Controle de Acesso")
        self.tab_activities = self.tab_view.add("Gerenciar Atividades")
        self.tab_sheets = self.tab_view.add("Gerenciar Planilhas")

        # Agora sim, crie as abas, passando o tab_view e user_data
        self.users_tab = UsersTab(self.tab_users, self.user_data)
        self.teams_tab = TeamsTab(self.tab_teams, self.user_data)
        self.blocks_tab = BlocksTab(self.tab_blocks, self.user_data)
        self.activities_tab = ActivitiesTab(self.tab_activities, self.user_data, manager=self)
        self.sheets_tab = SheetsTab(self.tab_sheets, self.user_data)
        
        # Configurar pesos do grid no main_frame
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(0, weight=1)
        
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
        
        # Empacotar o main_frame por último
        self.main_frame.pack(expand=True, fill="both", padx=15, pady=15)

    def update_counters(self):
        """Atualiza os contadores de usuários, equipes e bloqueios chamando métodos das abas."""
        try:
            if hasattr(self.users_tab, 'update_counter'):
                self.users_tab.update_counter()
            if hasattr(self.teams_tab, 'update_counter'):
                self.teams_tab.update_counter()
            if hasattr(self.blocks_tab, 'update_counter'):
                self.blocks_tab.update_counter()
        except Exception as e:
            logger.error(f"Erro ao atualizar contadores centralizados: {e}")
            messagebox.showerror("Erro", "Erro ao atualizar contadores")

    def get_equipes(self):
        """Retorna lista de todas as equipes do banco"""
        try:
            query = """
                SELECT nome 
                FROM equipes 
                ORDER BY nome
            """
            result = self.db.execute_query(query)
            return [row['nome'] for row in result] if result else []
        except Exception as e:
            logger.error(f"Erro ao buscar equipes: {e}")
            return []
