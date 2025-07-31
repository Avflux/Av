import customtkinter as ctk
import logging, os, time
from tkinter import messagebox, filedialog
from ..notifications.notification_manager import NotificationManager
from ...database.connection import DatabaseConnection
from ..dialogs.search_dialog import SearchFrame
from ..dialogs.activities_printer_dialog import ActivitiesPrinterDialog
from ..dialogs.dashboard_daily import DashboardDaily
from ..dialogs.perfil_dialog import PerfilFrame
from ..dialogs.user_manager import UserManager
from ..dialogs.activity_topframe import ActivityTopFrame
from ...utils.system_tray_icon import SystemTrayIcon
from ...utils.helpers import BaseWindow
from ...core.printer.templates.activities_printer import ActivitiesPrinter
from ...core.printer.query.query_activities import QueryActivities
from ...core.printer.observer.base_value_observer import BaseValueObserver
from ..components.activities.activity_controls import ActivityControls
from ..components.logic.activity_controls_logic import ActivityControlsLogic

logger = logging.getLogger(__name__)

class MainWindow(BaseWindow):
    def __init__(self, master, user_data):
        super().__init__(master)
        self.window_manager = master.window_manager if hasattr(master, 'window_manager') else None
        self.withdraw()  # Esconde temporariamente
        self.after(100, self._show_window)  # Mostra após 100ms
        
        # Configurar janela
        self.title(f"Sistema Chronos - Bem-vindo, {user_data['nome']}!")
        self.minsize(1400, 600)
        if self.window_manager:
            self.window_manager.position_window(self, parent=master)
        # Sempre usar o WindowManager para posicionamento e tamanho; não usar cálculo manual.
            
        # Inicialização dos componentes
        self.db = DatabaseConnection()
        self.user_data = user_data
        self.notification_manager = NotificationManager()
        self.notification_manager.initialize(self, self.user_data['nome'])
        
        # Mostrar mensagem de boas-vindas personalizada
        self.notification_manager.show_welcome_message(self.user_data['nome'])
        
        # Agendar lembrete de água para 1h após login
        self.notification_manager.schedule_water_reminder(self)
        
        # Initialize system tray icon
        self.system_tray = SystemTrayIcon(self, self.on_quit)
        # Override close button behavior
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # Configurar observer do valor base
        self.base_value_observer = BaseValueObserver(self.db)
        self.base_value_observer.attach(self)
        self.base_value = self.base_value_observer.get_base_value(self.user_data['id'])
        
        # Configuração dos pesos do grid para responsividade
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.current_frame = None
        
        # Forçar foco na janela
        self.lift()  # Traz a janela para frente
        self.focus_force()  # Força o foco
        
        # Configurar protocolo de foco
        self.bind('<FocusIn>', self._on_focus_in)
        self.bind('<FocusOut>', self._on_focus_out)
        
        # Adicionar atributo para controlar estado da janela
        self._minimized = False
        
        # Inicializar interface
        self.setup_ui()

        # Bind para salvar posição ao mover/redimensionar
        self.bind('<Configure>', self._on_configure)

    def _on_focus_in(self, event):
        """Garante opacidade total quando a janela ganha foco"""
        self.attributes('-alpha', 1.0)
        self.update_idletasks()

    def _on_focus_out(self, event):
        """Mantém opacidade quando perde foco"""
        self.attributes('-alpha', 1.0)
        self.update_idletasks()

    def _show_window(self):
        """Mostra a janela após um pequeno delay"""
        self.deiconify()
        self.attributes('-alpha', 1.0)
        self.lift()
        self.focus_force()
        # Abrir em tela cheia (maximizada)
        try:
            self.state('zoomed')  # Windows
        except Exception:
            self.attributes('-zoomed', True)  # Linux/Mac (fallback)

    def _on_configure(self, event):
        """Salva a posição e tamanho da janela ao mover/redimensionar"""
        if self.window_manager:
            self.window_manager.positions['main_window_geometry'] = self.geometry()
            self.window_manager._save_positions()

    def setup_ui(self):
        # Menu lateral
        self.sidebar = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew", rowspan=2)
        self.sidebar.grid_propagate(False)

        # Frame superior para informações globais
        self.top_frame = ctk.CTkFrame(self)
        self.top_frame.grid(row=0, column=1, sticky="new", padx=10, pady=(10, 0))

        # Adiciona o ActivityTopFrame ao top_frame
        self.activity_top_frame = ActivityTopFrame(self.top_frame, self.user_data)
        self.activity_top_frame.pack(fill="both", expand=True)

        # Área principal (abaixo do top_frame)
        self.main_area = ctk.CTkFrame(self)
        self.main_area.grid(row=1, column=1, sticky="nsew", padx=10, pady=10)
        self.main_area.grid_columnconfigure(0, weight=1)
        self.main_area.grid_rowconfigure(0, weight=1)

        # Setup do menu lateral
        self.setup_sidebar()
        
        # Inicialmente, mostrar a tela correta conforme tipo de usuário
        tipo = self.user_data.get('tipo_usuario', '').lower()
        if tipo == 'master':
            self.show_search()
        elif tipo == 'comum':
            self.show_activities()
        else:
            self.show_search()  # fallback

        # Adicionar após criar os frames principais
        self.update_idletasks()  # Força atualização do layout
        self.attributes('-alpha', 1.0)  # Reforça opacidade

        # Aplicar visibilidade conforme tipo de usuário
        self._apply_user_type_visibility()

    def show_search(self):
        """Mostra o frame de pesquisa"""
        self.clear_main_area()
        self._update_button_states("pesquisa")
        search_frame = SearchFrame(self.main_area, self.user_data)
        search_frame.pack(expand=True, fill="both", padx=10, pady=10)
        self.current_frame = search_frame

    def show_register_dialog(self):
        """Mostra o diálogo de gerenciamento de usuários"""
        self.clear_main_area()
        self._update_button_states("usuarios")  # Ajuste conforme o nome do botão
        user_frame = UserManager(self.main_area, self.user_data)
        user_frame.pack(expand=True, fill="both", padx=10, pady=10)
        self.current_frame = user_frame

    def setup_sidebar(self):
        """Configura o menu lateral"""
        # Logo
        self.logo_label = ctk.CTkLabel(
            self.sidebar,
            text="CHRONOS",
            font=("Roboto", 20, "bold")
        )
        self.logo_label.pack(pady=10, padx=20)

        self.btn_perfil = ctk.CTkButton(
            self.sidebar,
            text="Meu Perfil",
            fg_color="#FF5722",
            hover_color="#CE461B",
            command=self.show_profile,
        )
        self.btn_perfil.pack(pady=10, padx=20)

        # Botão de Atividades
        self.btn_atividades = ctk.CTkButton(
            self.sidebar,
            text="Atividades",
            fg_color="#FF5722",
            hover_color="#CE461B",
            command=self.show_activities,
        )
        self.btn_atividades.pack(pady=10, padx=20)
        
        self.btn_pesquisar = ctk.CTkButton(
            self.sidebar,
            text="Pesquisar",
            fg_color="#FF5722",
            hover_color="#CE461B",
            command=self.show_search
        )
        self.btn_pesquisar.pack(pady=10, padx=20)

         # Botão do Dashboard
        self.btn_dashboard = ctk.CTkButton(
            self.sidebar,
            text="Dashboard",
            fg_color="#FF5722",
            hover_color="#CE461B",
            command=self.show_dashboard,
        )
        self.btn_dashboard.pack(pady=10, padx=20)
        
        self.btn_cadastrar = ctk.CTkButton(
            self.sidebar,
            text="Usuários",
            fg_color="#FF5722",
            hover_color="#CE461B",
            command=self.open_user_management
        )
        self.btn_cadastrar.pack(pady=10, padx=20)

        # Adiciona botão de relatório
        self.btn_relatorio = ctk.CTkButton(
            self.sidebar,
            text="Gerar Relatório",
            command=self.generate_report,
            fg_color="#FF5722",
            hover_color="#CE461B"
        )
        self.btn_relatorio.pack(pady=10, padx=20)

        # Botão de logout
        self.btn_logout = ctk.CTkButton(
            self.sidebar,
            text="Logout",
            command=self.logout,
            fg_color="#FF5722",
            hover_color="#CE461B"
        )
        self.btn_logout.pack(side="bottom", pady=20, padx=20)

    def _update_button_states(self, active_section):
        # Define o mapeamento dos botões e suas seções
        button_section_map = {
            'perfil': self.btn_perfil,
            'atividades': self.btn_atividades,
            'pesquisa': self.btn_pesquisar,
            'dashboard': self.btn_dashboard,
            'usuarios': self.btn_cadastrar,
            'relatorio': self.btn_relatorio,
        }
        # Cores para ativo/inativo
        active_fg = ("#FF5722", "#CE461B")
        inactive_fg = "#FF5722"
        # Atualiza todos os botões exceto logout
        for section, btn in button_section_map.items():
            if section == active_section:
                btn.configure(fg_color=active_fg)
                btn.configure(state="disabled")
            else:
                btn.configure(fg_color=inactive_fg)
                btn.configure(state="normal")
        # O botão de logout nunca é afetado

    def show_profile(self):
        self.clear_main_area()
        self._update_button_states("perfil")
        perfil_frame = PerfilFrame(self.main_area, self.user_data)
        perfil_frame.pack(expand=True, fill="both", padx=10, pady=10)
        self.current_frame = perfil_frame

    def clear_main_area(self):
        if self.current_frame:
            self.current_frame.destroy()
            self.current_frame = None

    def logout(self):
        """Realiza o logout do administrador"""
        if messagebox.askyesno("Logout", "Deseja realmente sair?"):
            try:
                # Pausar todas as atividades ativas do usuário
                logic = ActivityControlsLogic(self.db)
                logic.pause_all_active_activities(self.user_data['id'])
                # Primeiro desregistrar observers e limpar outros recursos
                self.unregister_observers()
                
                # Remover o ícone da bandeja antes de fechar a janela
                if hasattr(self, 'system_tray'):
                    self.system_tray.cleanup()
                    time.sleep(0.2)  # Dar tempo para o Windows processar
                
                logger.info("Administrador realizou logout")
                self.master.deiconify()
                self.destroy()
                
            except Exception as e:
                logger.error(f"Erro durante logout do administrador: {str(e)}")
                messagebox.showerror("Erro", "Erro ao realizar logout")

    def on_close(self):
        """Manipula o evento de fechar a janela"""
        try:
            logger.debug("Minimizando MainWindow para bandeja")
            self._minimized = True
            if hasattr(self, 'system_tray'):
                self.system_tray.minimize_to_tray()
        except Exception as e:
            logger.error(f"Erro ao minimizar MainWindow: {e}")

    def on_quit(self):
            if messagebox.askyesno("Sair", "Deseja realmente sair do sistema?"):
                self.master.deiconify()
                self.destroy()

    def generate_report(self, selected_date=None):
        """Gera o relatório de atividades"""
        try:
            self._show_activities_report_dialog()
        except Exception as e:
            logger.error(f"[REPORT] Erro ao abrir diálogo de relatório: {str(e)}")
            messagebox.showerror(
                "Erro",
                "Erro ao abrir gerador de relatório"
            )

    def _show_activities_report_dialog(self):
        """Mostra o diálogo de relatório de atividades"""
        try:
            logger.debug("[REPORT] Abrindo diálogo de relatório")
            ActivitiesPrinterDialog(
                self, 
                self._generate_report,
                self.db,
                self.user_data['id']
            )
        except Exception as e:
            logger.error(f"[REPORT] Erro ao abrir diálogo de relatório: {str(e)}")
            messagebox.showerror(
                "Erro",
                "Não foi possível abrir o diálogo de relatório. Por favor, tente novamente."
            )

    def update_base_value(self, value):
        """
        Compatibilidade: método chamado por observers de valor base.
        """
        logger.info(f"[MainWindow] update_base_value chamado com valor: {value}")
        self.base_value = value

    def _generate_report(self, selected_date=None):
        """Implementação da geração do relatório"""
        try:
            logger.info("[REPORT] Iniciando geração de relatório")
            
            # Solicita local para salvar o arquivo
            file_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                title="Salvar Relatório"
            )
            
            if file_path:
                logger.debug(f"[REPORT] Local selecionado para salvar: {file_path}")
                
                try:
                    # Tenta abrir o arquivo para verificar se está em uso
                    with open(file_path, 'wb') as _:
                        pass
                    
                    # Busca os dados das atividades
                    query = QueryActivities(self.db)
                    month = selected_date['month'] if selected_date else None
                    year = selected_date['year'] if selected_date else None
                    data = query.get_activities_report_data(
                        self.user_data['id'],
                        month=month,
                        year=year
                    )
                    
                    # Adiciona o valor base aos dados
                    data['base_value'] = self.base_value
                    
                    # Gera o relatório
                    printer = ActivitiesPrinter()
                    logo_path = os.path.join('icons', 'logo_light.png')
                    
                    logger.debug("[REPORT] Iniciando impressão do relatório")
                    printer.generate_report(file_path, data, logo_path)
                    
                    logger.info("[REPORT] Relatório gerado com sucesso")
                    messagebox.showinfo(
                        "Sucesso",
                        "Relatório gerado com sucesso!"
                    )
                    
                except PermissionError:
                    logger.error("[REPORT] Arquivo PDF está aberto ou em uso")
                    messagebox.showerror(
                        "Erro",
                        "Não foi possível gerar o relatório pois o arquivo PDF está aberto.\n"
                        "Por favor, feche o arquivo e tente novamente."
                    )
                except IOError as e:
                    logger.error(f"[REPORT] Erro de acesso ao arquivo: {str(e)}")
                    messagebox.showerror(
                        "Erro",
                        "Não foi possível acessar o arquivo para gravação.\n"
                        "Verifique se você tem permissão para salvar neste local."
                    )
                
        except Exception as e:
            logger.error(f"[REPORT] Erro ao gerar relatório: {str(e)}")
            messagebox.showerror(
                "Erro",
                "Ocorreu um erro inesperado ao gerar o relatório.\n"
                "Por favor, tente novamente."
            )

    def show_dashboard(self):
        """Abre a janela do dashboard"""
        self.clear_main_area()
        self._update_button_states("dashboard")
        dashboard_frame = DashboardDaily(self.main_area, db_connection=self.db, user_data=self.user_data)
        dashboard_frame.pack(expand=True, fill="both", padx=10, pady=10)
        self.current_frame = dashboard_frame

    def destroy(self):
        """Sobrescreve o método destroy para limpar recursos"""
        try:
            # Primeiro limpar o system tray
            if hasattr(self, 'system_tray'):
                self.system_tray.cleanup()
                time.sleep(0.1)  # Pequeno delay para processamento
            
            # Então destruir a janela
            super().destroy()
        except Exception as e:
            logger.error(f"Erro ao destruir MainWindow: {e}")
            # Tentar forçar a destruição mesmo em caso de erro
            super().destroy()

    def open_user_management(self):
        """Abre o painel de gerenciamento de usuários"""
        self.clear_main_area()
        self._update_button_states("usuarios")
        self.current_frame = UserManager(self.main_area, self.user_data)
        self.current_frame.pack(expand=True, fill="both", padx=10, pady=10)

    def unregister_observers(self):
        """Remove todos os observadores registrados de forma segura"""
        try:
            # Remover observador do BaseValueObserver
            if hasattr(self, 'base_value_observer'):
                try:
                    self.base_value_observer.detach(self)
                    logger.debug("[OBSERVER] Observador removido do BaseValueObserver")
                except ValueError:
                    logger.debug("[OBSERVER] AdminWindow não estava registrado no BaseValueObserver")
            
            logger.info("[OBSERVER] Processo de remoção de observadores concluído")
            
        except Exception as e:
            logger.error(f"[OBSERVER] Erro inesperado ao remover observadores: {str(e)}", exc_info=True)

    def show_window(self):
        """Método para mostrar a janela novamente"""
        try:
            self._minimized = False
            self.deiconify()
            self.lift()
            self.state('normal')
            self.focus_force()
            # Restaurar geometria se existir
            if self.window_manager and 'main_window_geometry' in self.window_manager.positions:
                self.geometry(self.window_manager.positions['main_window_geometry'])
        except Exception as e:
            logger.error(f"Erro ao mostrar MainWindow: {e}")


    def withdraw(self):
        """Sobrescreve o método withdraw para controlar o estado"""
        try:
            self._minimized = True
            super().withdraw()
        except Exception as e:
            logger.error(f"Erro ao ocultar MainWindow: {e}")

    def winfo_exists(self):
        """Sobrescreve winfo_exists para considerar estado minimizado"""
        try:
            exists = super().winfo_exists()
            return exists and (not self._minimized or hasattr(self, 'system_tray'))
        except Exception:
            return False

    def show_activities(self):
        self.clear_main_area()
        self._update_button_states("atividades")
        activities_frame = ctk.CTkFrame(self.main_area)
        activities_frame.pack(fill="both", expand=True)
        from ..components.activities.activity_table import ActivityTable
        self.activity_controls = ActivityControls(
            activities_frame,
            self.user_data,
            self.db,
            self.handle_activity_action,
            active_label=self.activity_top_frame.active_activity_label,
            selected_label=self.activity_top_frame.selected_activity_label,
            daily_time_manager=getattr(self.activity_top_frame, 'daily_time_manager', None)
        )
        self.activity_controls.pack(fill="x", pady=5)
        self.activity_table = ActivityTable(
            activities_frame,
            self.user_data,
            self.db
        )
        self.activity_table.pack(fill="both", expand=True)
        self.current_frame = activities_frame

        # BIND: sempre que selecionar uma linha, atualiza o label do topo
        def on_tree_select(event):
            selected = self.activity_table.tree.selection()
            if selected:
                item = self.activity_table.tree.item(selected[0])
                activity_data = {
                    'id': item['values'][0],
                    'description': item['values'][1],
                    'atividade': item['values'][2],
                    'start_time': item['values'][3],
                    'end_time': item['values'][4],
                    'status': item['values'][-1]
                }
                self.activity_controls.on_activity_selected(activity_data)
            else:
                self.activity_controls.on_activity_selected(None)
        self.activity_table.tree.bind("<<TreeviewSelect>>", on_tree_select)

    def handle_activity_action(self, action):
        if hasattr(self, 'activity_controls'):
            self.activity_controls.handle_activity_action(action)

    def _apply_user_type_visibility(self):
        """
        Oculta botões e frames conforme o tipo de usuário:
        - master: oculta botão Atividades e o top_frame
        - comum: oculta botões Pesquisar e Usuários
        """
        tipo = self.user_data.get('tipo_usuario', '').lower()
        if tipo == 'master':
            # Oculta botão Atividades
            if hasattr(self, 'btn_atividades'):
                self.btn_atividades.pack_forget()
            # Oculta o top_frame
            if hasattr(self, 'top_frame'):
                self.top_frame.grid_remove()
        elif tipo == 'comum':
            # Oculta botões Pesquisar e Usuários
            if hasattr(self, 'btn_pesquisar'):
                self.btn_pesquisar.pack_forget()
            if hasattr(self, 'btn_cadastrar'):
                self.btn_cadastrar.pack_forget()
    