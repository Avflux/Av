import os
import sys
import logging
import customtkinter as ctk
from tkinter import ttk
from tkinter import filedialog
from datetime import datetime
from PIL import Image
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import numpy as np
from app.core.printer.templates.dashboard_printer import DashboardPrinter
from app.core.printer.query.dashboard_query import DashboardQuery
from app.config.settings import APP_CONFIG
from ...database.connection import DatabaseConnection

logger = logging.getLogger(__name__)

class DashboardDaily(ctk.CTkFrame):
    def __init__(self, parent, db_connection=None, user_data=None):
        super().__init__(parent)
        
        # Inicializar dicionário de imagens ANTES de qualquer configuração de interface
        self.images = {}
        
        logger.info("[THEME] Inicializando DashboardDaily")
        
        # Armazenar referências importantes
        self.parent = parent
        self.db = db_connection or DatabaseConnection()
        self.user_data = user_data
        self.dashboard_query = DashboardQuery(self.db)
        self.container = ctk.CTkFrame(self, fg_color="#FFF")
        self.container.pack(fill="both", expand=True)
        self.cores = {}
        self.dados_atrasos = {}
        self.dados_atrasos_periodos = {}
        self.selected_equipe = None
        
        self._on_theme_change_callback_id = self.after(100, self._on_theme_change)
        self._on_theme_change_callback = self._on_theme_change
        ctk.AppearanceModeTracker.add(self._on_theme_change_callback, self)
        logger.debug("[THEME] Callback de tema registrado")
        
        # Primeiro carregar os dados, depois criar a interface
        self.carregar_dados_atrasos()
        
        # Forçar atualização inicial do tema
        self._on_theme_change()
        
        # Criar interface após configuração inicial
        self.setup_interface()
        
        # Garantir novamente que a janela esteja em foco após a configuração completa
        self.after(100, self.lift)
        self.after(200, self.focus_force)

    def criar_title_bar(self):
        """Cria a barra de título personalizada"""
        title_bar = ctk.CTkFrame(self, height=35)
        title_bar.pack(fill="x", pady=0)
        title_bar.pack_propagate(False)
        
        # Título
        title_label = ctk.CTkLabel(
            title_bar,
            text="Interest Dashboard Dailys 2025",
            font=("Arial", 12, "bold")
        )
        title_label.pack(side="left", padx=10)
        
        # Frame para botões
        buttons_frame = ctk.CTkFrame(title_bar, fg_color="transparent")
        buttons_frame.pack(side="right", padx=5)
        
        # Botão Minimizar
        min_btn = ctk.CTkButton(
            buttons_frame,
            text="—",
            width=35,
            height=25,
            command=self.iconify
        )
        min_btn.pack(side="left", padx=2)
        
        # Botão Maximizar
        self.max_btn = ctk.CTkButton(
            buttons_frame,
            text="□",
            width=35,
            height=25,
            command=self.toggle_maximize
        )
        self.max_btn.pack(side="left", padx=2)
        
        # Botão Fechar
        close_btn = ctk.CTkButton(
            buttons_frame,
            text="✕",
            width=35,
            height=25,
            command=self.destroy,
            fg_color="#D41919",
            hover_color="#AA1515"
        )
        close_btn.pack(side="left", padx=2)
        
        # Variável para controlar estado maximizado
        self.is_maximized = False

    def toggle_maximize(self):
        """Alterna entre janela maximizada e normal"""
        if self.is_maximized:
            self.restore_window()
        else:
            self.maximize_window()
        
        # Garantir que a janela permaneça no topo
        self.after(100, lambda: self.attributes('-topmost', True))
        self.after(200, self.lift)
        self.after(300, self.focus_force)

    def carregar_dados_atrasos(self, equipe_nome=None):
        """Carrega os dados de atrasos do banco de dados para a equipe selecionada ou todas"""
        try:
            if self.db is None:
                logger.error("Conexão com banco de dados não inicializada")
                return
            
            # Buscar equipe_id se equipe_nome não for 'Todos'
            team_id = None
            if equipe_nome and equipe_nome != "Todos":
                equipe_query = "SELECT id FROM equipes WHERE nome = %s"
                equipe_result = self.db.execute_query(equipe_query, (equipe_nome,))
                if equipe_result:
                    team_id = equipe_result[0]['id']
            
            # Usar o DashboardQuery para buscar os dados
            dados = self.dashboard_query.get_dashboard_data(
                user_id=None,
                team_id=team_id,
                period='week'
            )
            
            if dados and 'atrasos' in dados:
                # Armazenar todos os períodos
                self.dados_atrasos_periodos = dados['atrasos']
                # Processar apenas os motivos da semana_atual
                self.dados_atrasos = self._processar_dados_atrasos(self.dados_atrasos_periodos.get('semana_atual', {}))
                logger.info(f"Dados de atrasos carregados: {len(self.dados_atrasos)} registros da semana atual")
                logger.info(f"Períodos carregados: {list(self.dados_atrasos_periodos.keys())}")
            else:
                logger.warning("Nenhum dado de atraso encontrado")
                self.dados_atrasos = {}
                self.dados_atrasos_periodos = {}
                
        except Exception as e:
            logger.error(f"Erro ao carregar dados de atrasos: {e}")
            self.dados_atrasos = {}
            self.dados_atrasos_periodos = {}

    def _processar_dados_atrasos(self, dados_brutos):
        """Processa os dados brutos do banco para garantir estrutura consistente"""
        dados_processados = {}
        
        try:
            # Verificar se dados_brutos é um dicionário ou lista
            if isinstance(dados_brutos, dict):
                for motivo, dados in dados_brutos.items():
                    dados_processados[motivo] = self._normalizar_dados_atraso(dados)
            elif isinstance(dados_brutos, list):
                # Se for uma lista, processar cada item
                for item in dados_brutos:
                    if isinstance(item, dict):
                        motivo = item.get('motivo', 'Motivo não especificado')
                        dados_processados[motivo] = self._normalizar_dados_atraso(item)
            else:
                logger.warning(f"Formato de dados não reconhecido: {type(dados_brutos)}")
                
        except Exception as e:
            logger.error(f"Erro ao processar dados de atrasos: {e}")
            
        return dados_processados

    def _normalizar_dados_atraso(self, dados):
        """Normaliza um item de dados de atraso para ter a estrutura esperada"""
        if not isinstance(dados, dict):
            return {
                'quantidade': 0,
                'tempo': 0,
                'impacto': 'Baixo'
            }
        
        # Calcular impacto baseado na quantidade e tempo
        quantidade = dados.get('quantidade', 0)
        tempo = dados.get('tempo', 0)
        
        # Lógica para determinar impacto
        if quantidade >= 5 or tempo >= 3:
            impacto = 'Alto'
        elif quantidade >= 3 or tempo >= 2:
            impacto = 'Médio'
        else:
            impacto = 'Baixo'
        
        return {
            'quantidade': quantidade,
            'tempo': tempo,
            'impacto': dados.get('impacto', impacto)
        }

    def recarregar_dashboard_por_equipe(self, equipe_nome):
        """Recarrega apenas os dados do dashboard para a equipe selecionada, sem recriar toda a interface."""
        try:
            logger.info(f"Recarregando dashboard para equipe: {equipe_nome}")
            # Carregar novos dados
            self.carregar_dados_atrasos(equipe_nome)
            # Limpar container antes de recriar a interface
            for widget in self.container.winfo_children():
                if widget != self.container:
                    widget.destroy()
            # Recriar interface com novos dados
            self.setup_interface()
            # Forçar atualização do tema
            self._on_theme_change()
            # Atualizar Treeview de atrasos para 'semana_atual'
            if hasattr(self, 'treeview_atrasos'):
                self.carregar_controle_atraso(self.dados_atrasos_periodos, periodo='semana_atual')
            logger.info("Dashboard recarregado com sucesso")
        except Exception as e:
            logger.error(f"Erro ao recarregar dashboard: {e}")
            # Em caso de erro, tentar criar interface básica
            try:
                self.dados_atrasos = {}
                self.dados_atrasos_periodos = {}
                self.setup_interface()
            except Exception as e2:
                logger.error(f"Erro crítico ao criar interface básica: {e2}")

    def setup_cores(self):
        """Configura as cores baseadas no tema atual"""
        is_light = ctk.get_appearance_mode().lower() == "light"
        logger.debug(f"[THEME] Configurando cores para modo: {'Light' if is_light else 'Dark'}")
        
        self.cores = {
            'primaria': "#FF5722",     # Laranja Interest (mantido fixo)
            'secundaria': "#343333" if is_light else "#FFFFFF",
            'destaque': "#DE2020",     # Vermelho
            'sucesso': "#107C10",      # Verde (modificado)
            'fundo': "#FFFFFF" if is_light else "#2B2B2B",
            'texto': "#333333" if is_light else "#FFFFFF",
            'texto_card': "#FF5722",    # Laranja Interest (mantido fixo)
            'fundo_card': "#FFFFFF" if is_light else "#2B2B2B",
            'borda': "#E0E0E0" if is_light else "#555555"
        }
    
    def setup_interface(self):
        """Configura a interface do usuário"""
        # Atualizar a cor do container existente
        self.container.configure(fg_color=self.cores['fundo'])
        
        # Cabeçalho com logo e informações
        self.criar_cabecalho()
        
        # Área principal dividida em duas colunas
        self.criar_area_principal()

    def criar_cabecalho(self):
        """Cria o cabeçalho com logo e informações"""
        # Salvar valor atual da equipe selecionada
        valor_atual = self.selected_equipe.get() if self.selected_equipe else "Todos"
        header = ctk.CTkFrame(self.container, fg_color=self.cores['fundo'], height=60)
        header.pack(fill="x", pady=0)
        header.pack_propagate(False)
        
        # Frame para a logo
        logo_frame = ctk.CTkFrame(header, fg_color="transparent")
        logo_frame.pack(side="left", padx=20)
        
        # Carregando e exibindo a logo
        try:
            # Obter caminho base do projeto
            if hasattr(sys, '_MEIPASS'):
                base_path = sys._MEIPASS
            else:
                # Subir 4 níveis: dialogs -> ui -> app -> raiz do projeto
                current_dir = os.path.dirname(__file__)  # dialogs
                ui_dir = os.path.dirname(current_dir)    # ui
                app_dir = os.path.dirname(ui_dir)        # app
                base_path = os.path.dirname(app_dir)     # raiz do projeto
            
            # Construir caminhos das imagens usando APP_CONFIG
            logo_light = os.path.join(base_path, APP_CONFIG['icons']['logo_light'])
            logo_dark = os.path.join(base_path, APP_CONFIG['icons']['logo_dark'])
            
            logger.debug(f"[LOGO] Caminho base: {base_path}")
            logger.debug(f"[LOGO] Caminho light: {logo_light}")
            logger.debug(f"[LOGO] Caminho dark: {logo_dark}")
            
            if os.path.exists(logo_light) and os.path.exists(logo_dark):
                # Carregar imagens
                light_image = Image.open(logo_light)
                dark_image = Image.open(logo_dark)
                
                # Definir tamanho para o logo
                logo_width = 200
                logo_height = int(logo_width * (65.2/318.6))
                
                # Criar e armazenar a referência da imagem
                self.images['logo'] = ctk.CTkImage(
                    light_image=light_image,
                    dark_image=dark_image,
                    size=(logo_width, logo_height)
                )
                
                # Criar label com a imagem
                logo_label = ctk.CTkLabel(
                    logo_frame,
                    image=self.images['logo'],
                    text=""
                )
                logo_label.pack()
                logger.info("[LOGO] Logo carregado com sucesso")
                
            else:
                raise FileNotFoundError(f"Arquivos de logo não encontrados. Light: {os.path.exists(logo_light)}, Dark: {os.path.exists(logo_dark)}")
            
        except Exception as e:
            logger.error(f"[LOGO] Erro ao carregar logo: {str(e)}")
            # Fallback para texto
            ctk.CTkLabel(
                logo_frame,
                text="INTEREST ENGENHARIA",
                font=("Arial Black", 24),
                text_color=self.cores['primaria']
            ).pack()
        
        # Informações do relatório
        info_frame = ctk.CTkFrame(header, fg_color="transparent")
        info_frame.pack(side="right", padx=20)
        
        # Buscar informações do banco
        
        # Buscar nomes das equipes para o menu suspenso
        try:
            equipes = ["Todos"]
            equipes_query = "SELECT nome FROM equipes ORDER BY nome"
            equipes_result = self.db.execute_query(equipes_query)
            if equipes_result:
                equipes += [row['nome'] for row in equipes_result]
        except Exception as e:
            logger.error(f"Erro ao buscar equipes para OptionMenu: {e}")
            equipes = ["Todos"]

        # Variável de controle para equipe selecionada
        self.selected_equipe = ctk.StringVar(value=valor_atual)

        # Função callback para mudança de equipe
        def on_equipe_change(selected):
            self.selected_equipe.set(selected)
            self.recarregar_dashboard_por_equipe(selected)

        # Frame para equipe (menu suspenso)
        equipe_frame = ctk.CTkFrame(info_frame, fg_color="transparent")
        equipe_frame.pack(side="left", padx=15)
        ctk.CTkLabel(
            equipe_frame,
            text="Equipe:",
            font=("Arial", 12),
            text_color=self.cores['secundaria']
        ).pack(side="left", padx=5)
        ctk.CTkOptionMenu(
            equipe_frame,
            variable=self.selected_equipe,
            values=equipes,
            command=on_equipe_change,
            fg_color="#FF5722",
            button_color="#FF5722",
            button_hover_color="#CE461B",
            width=140,
            height=28
        ).pack(side="left")
        # Restaurar valor selecionado após criar OptionMenu
        self.selected_equipe.set(valor_atual)

        # Frame para data
        data_frame = ctk.CTkFrame(info_frame, fg_color="transparent")
        data_frame.pack(side="left", padx=15)
        ctk.CTkLabel(
            data_frame,
            text="Data:",
            font=("Arial", 12),
            text_color=self.cores['secundaria']
        ).pack(side="left", padx=5)
        ctk.CTkLabel(
            data_frame,
            text=datetime.now().strftime("%d/%m/%Y"),
            font=("Arial", 12, "bold"),
            text_color=self.cores['secundaria']
        ).pack(side="left")
        
        # Alterar o botão de impressão para seguir o tema
        print_button = ctk.CTkButton(
            info_frame,
            text="Salvar PDF",
            command=self.salvar_relatorio,
            font=("Arial", 12),
            fg_color=self.cores['primaria'],
            hover_color="#CE461B",  # Cor de hover padrão do projeto
            width=80
        )
        print_button.pack(side="left", padx=15)
    
    def criar_area_principal(self):
        # Frame principal com duas colunas
        main_frame = ctk.CTkFrame(self.container, fg_color=self.cores['fundo'])
        main_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Configurar o grid do main_frame
        main_frame.grid_columnconfigure(0, weight=7)  # 70% para coluna esquerda
        main_frame.grid_columnconfigure(1, weight=3)  # 30% para coluna direita
        main_frame.grid_rowconfigure(0, weight=1)     # Permite expansão vertical
        
        # Frame geral da coluna esquerda
        frame_esquerda = ctk.CTkFrame(
            main_frame,
            fg_color=self.cores['fundo_card'],
            border_width=1,
            border_color=self.cores['borda']
        )
        frame_esquerda.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        
        # Configurar grid do frame esquerda para expansão
        frame_esquerda.grid_rowconfigure(1, weight=1)  # Row do frame_grafico expande
        frame_esquerda.grid_columnconfigure(0, weight=1)
        
        # Indicadores principais (row 0)
        frame_indicadores = ctk.CTkFrame(frame_esquerda, fg_color=self.cores['fundo_card'], height=100)
        frame_indicadores.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        frame_indicadores.grid_propagate(False)
        self.criar_indicadores_principais(frame_indicadores)
        
        # Container do gráfico (row 1 - expansível)
        frame_grafico = ctk.CTkFrame(frame_esquerda, fg_color=self.cores['fundo_card'])
        frame_grafico.grid(row=1, column=0, sticky="nsew", padx=5, pady=(0, 5))
        self.criar_grafico_combinado(frame_grafico)
        
        # Legendas (row 2)
        frame_legendas = ctk.CTkFrame(
            frame_esquerda,
            fg_color=self.cores['fundo_card'],
            height=60,
            border_width=1,            # Adiciona borda
            border_color=self.cores['borda']  # Usa a cor de borda padrão
        )
        frame_legendas.grid(row=2, column=0, sticky="ew", padx=5, pady=(0, 5))
        frame_legendas.grid_propagate(False)
        self.criar_legenda_unificada(frame_legendas)
        
        # Frame geral da coluna direita
        frame_direita = ctk.CTkFrame(
            main_frame,
            fg_color=self.cores['fundo_card'],
            border_width=1,
            border_color=self.cores['borda']
        )
        frame_direita.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        
        # Configurar grid do frame direita para expansão
        frame_direita.grid_rowconfigure(0, weight=1)
        frame_direita.grid_columnconfigure(0, weight=1)
        
        # Frame para controle de atrasos
        frame_atrasos = ctk.CTkFrame(
            frame_direita,
            fg_color=self.cores['fundo_card']
        )
        frame_atrasos.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self.criar_controle_atrasos(frame_atrasos)
        
        # Vincular evento de redimensionamento da janela
        self.bind('<Configure>', self.atualizar_layout)

    def atualizar_layout(self, event=None):
        """Atualiza as proporções quando a janela é redimensionada"""
        if event and event.widget == self:
            largura = event.width
            altura = event.height
            
            # Garantir tamanho mínimo
            if largura < 1300 or altura < 700:
                return
            
            # Atualizar larguras das colunas
            for widget in self.container.winfo_children():
                if isinstance(widget, ctk.CTkFrame):
                    for col in widget.winfo_children():
                        if col.winfo_name() == '!ctkframe':  # Coluna esquerda
                            col.configure(width=int(largura * 0.7))
                        elif col.winfo_name() == '!ctkframe2':  # Coluna direita
                            col.configure(width=int(largura * 0.3))
            
            # Atualizar gráfico matplotlib (SUBSTITUIR a parte do canvas)
            if hasattr(self, 'canvas_matplotlib'):
                self.atualizar_matplotlib_grafico()

    def criar_indicadores_principais(self, parent):
        # Frame principal dos indicadores sem borda
        frame = ctk.CTkFrame(
            parent,
            fg_color=self.cores['fundo_card'],
            border_width=0  # Removida a borda intermediária
        )
        frame.pack(fill="x", pady=(0, 5))
        
        # Container interno para os cards com padding
        container = ctk.CTkFrame(frame, fg_color="transparent")
        container.pack(fill="x", padx=10, pady=10)
        
        # Calcular valores dos indicadores
        indicadores = [
            ("Indicador Semana Anterior", self.calcular_valor_periodo('semana_anterior'), "▲" if self.calcular_valor_periodo('semana_anterior') >= self.calcular_valor_periodo('semana_atual') else "▼"),
            ("Indicador Semana Atual", self.calcular_valor_periodo('semana_atual'), "▲" if self.calcular_valor_periodo('semana_atual') >= self.calcular_valor_periodo('semana_anterior') else "▼"),
            ("Indicador Média Semestral", self.calcular_valor_periodo('semestral'), "▲" if self.calcular_valor_periodo('semestral') >= self.calcular_valor_periodo('anual') else "▼"),
            ("Indicador Média Anual", self.calcular_valor_periodo('anual'), "▲" if self.calcular_valor_periodo('anual') >= self.calcular_valor_periodo('semestral') else "▼")
        ]

        for titulo, valor, tendencia in indicadores:
            # Card individual com borda mais suave
            card = ctk.CTkFrame(
                container,
                fg_color=self.cores['fundo_card'],
                border_width=1,
                border_color=self.cores['borda']
            )
            card.pack(side="left", expand=True, fill="both", padx=5)
            
            ctk.CTkLabel(
                card,
                text=titulo,
                font=("Arial", 12),
                text_color=self.cores['texto_card']
            ).pack(pady=(10, 5))
            
            valor_frame = ctk.CTkFrame(card, fg_color="transparent")
            valor_frame.pack(pady=(0, 10))
            
            ctk.CTkLabel(
                valor_frame,
                text=f"{valor:.1f}%",  # Formatando para uma casa decimal
                font=("Arial Black", 24),
                text_color=self.cores['primaria']
            ).pack(side="left", padx=5)
            
            # Modificado: Cor fixa para cada símbolo
            cor_tendencia = self.cores['sucesso'] if tendencia == "▲" else self.cores['destaque']  # Verde para ▲, Vermelho para ▼
            ctk.CTkLabel(
                valor_frame,
                text=tendencia,
                font=("Arial", 24),  # Aumentado de 18 para 24
                text_color=cor_tendencia
            ).pack(side="left")
    
    def criar_grafico_combinado(self, parent):
        frame = ctk.CTkFrame(
            parent, 
            fg_color=self.cores['fundo_card'], 
            border_width=1, 
            border_color=self.cores['borda']
        )
        frame.pack(fill="both", expand=True)
        
        # Título do gráfico
        ctk.CTkLabel(
            frame,
            text="ANÁLISE INTEGRADA DE INDICADORES E ATRASOS",
            font=("Arial Black", 14),
            text_color=self.cores['texto_card']
        ).pack(pady=10)
        
        # Converter dados do banco para o formato do gráfico
        self.dados_grafico = {
            'Semana Anterior': {
                'valor': self.calcular_valor_periodo('semana_anterior'),
                'tempo': self.calcular_tempo_periodo('semana_anterior'),
                'atrasos': self.calcular_atrasos_periodo('semana_anterior'),
                'cor': '#015BFF'
            },
            'Semana Atual': {
                'valor': self.calcular_valor_periodo('semana_atual'),
                'tempo': self.calcular_tempo_periodo('semana_atual'),
                'atrasos': self.calcular_atrasos_periodo('semana_atual'),
                'cor': '#DE2020'
            },
            'Média Semestral': {
                'valor': self.calcular_valor_periodo('semestral'),
                'tempo': self.calcular_tempo_periodo('semestral'),
                'atrasos': self.calcular_atrasos_periodo('semestral'),
                'cor': '#0F510C'
            },
            'Média Anual': {
                'valor': self.calcular_valor_periodo('anual'),
                'tempo': self.calcular_tempo_periodo('anual'),
                'atrasos': self.calcular_atrasos_periodo('anual'),
                'cor': '#FFB800'
            }
        }
        
        # Criar figura matplotlib
        self.criar_matplotlib_grafico(frame)

    def criar_matplotlib_grafico(self, parent):
        """Cria o gráfico usando matplotlib"""
        # Configurar estilo baseado no tema
        is_light = ctk.get_appearance_mode().lower() == "light"
        
        # Criar figura com fundo transparente
        self.fig = Figure(figsize=(12, 6), dpi=100)
        self.fig.patch.set_facecolor(self.cores['fundo_card'])
        
        # Criar subplot
        self.ax = self.fig.add_subplot(111)
        
        # Configurar cores do gráfico
        self.ax.set_facecolor(self.cores['fundo_card'])
        
        # Criar o canvas
        self.canvas_matplotlib = FigureCanvasTkAgg(self.fig, parent)
        self.canvas_matplotlib.get_tk_widget().pack(fill="both", expand=True, padx=20, pady=20)
        
        # Desenhar o gráfico inicial
        self.atualizar_matplotlib_grafico()

    def atualizar_matplotlib_grafico(self):
        """Atualiza o gráfico matplotlib"""
        if not hasattr(self, 'ax') or not hasattr(self, 'dados_grafico'):
            return
        
        # Limpar o gráfico
        self.ax.clear()
        
        # Configurar cores baseadas no tema
        is_light = ctk.get_appearance_mode().lower() == "light"
        cor_texto = "#333333" if is_light else "#FFFFFF"
        cor_grade = "#E0E0E0" if is_light else "#555555"
        
        # Configurar fundo
        self.ax.set_facecolor(self.cores['fundo_card'])
        self.fig.patch.set_facecolor(self.cores['fundo_card'])
        
        # Preparar dados
        periodos = list(self.dados_grafico.keys())
        valores = [self.dados_grafico[p]['valor'] for p in periodos]
        tempos = [self.dados_grafico[p]['tempo'] for p in periodos]
        cores = [self.dados_grafico[p]['cor'] for p in periodos]
        
        # Calcular atrasos totais
        atrasos_totais = []
        for p in periodos:
            if 'atrasos' in self.dados_grafico[p]:
                total = sum(self.dados_grafico[p]['atrasos'].values())
            else:
                total = 0
            atrasos_totais.append(total)
        
        # Posições das barras
        x_pos = np.arange(len(periodos))
        width = 0.25
        
        # Criar barras principais (valores percentuais)
        bars1 = self.ax.bar(x_pos - width, valores, width, 
                        color=cores, alpha=0.8, label='Eficiência (%)')
        
        # Criar barras de atrasos (motivos)
        bars2 = self.ax.bar(x_pos, atrasos_totais, width, 
                        color='#FF6B6B', alpha=0.7, label='Qtd. Atrasos')
        
        # Criar linha de tempo (eixo secundário)
        ax2 = self.ax.twinx()
        line = ax2.plot(x_pos + width/2, tempos, 'o-', 
                    color='#4ECDC4', linewidth=2, markersize=8, 
                    label='Tempo Atraso (dias)')
        
        # Configurar eixos
        self.ax.set_xlabel('Períodos', color=cor_texto, fontweight='bold')
        self.ax.set_ylabel('Eficiência (%) / Qtd. Atrasos', color=cor_texto, fontweight='bold')
        ax2.set_ylabel('Tempo de Atraso (dias)', color=cor_texto, fontweight='bold')
        
        # Configurar ticks
        self.ax.set_xticks(x_pos)
        self.ax.set_xticklabels(periodos, rotation=15, ha='right', color=cor_texto)
        self.ax.tick_params(colors=cor_texto)
        ax2.tick_params(colors=cor_texto)
        
        # Configurar grades
        self.ax.grid(True, alpha=0.3, color=cor_grade)
        self.ax.set_axisbelow(True)
        
        # Adicionar valores nas barras
        for i, (bar1, bar2) in enumerate(zip(bars1, bars2)):
            # Valor percentual
            height1 = bar1.get_height()
            self.ax.text(bar1.get_x() + bar1.get_width()/2., height1 + 1,
                        f'{height1:.1f}%', ha='center', va='bottom', 
                        color=cor_texto, fontweight='bold', fontsize=9)
            
            # Quantidade de atrasos
            height2 = bar2.get_height()
            if height2 > 0:
                self.ax.text(bar2.get_x() + bar2.get_width()/2., height2 + 0.5,
                            f'{int(height2)}', ha='center', va='bottom', 
                            color=cor_texto, fontweight='bold', fontsize=9)
            
            # Tempo formatado
            tempo_formatado = self.formatar_tempo(tempos[i])
            ax2.text(x_pos[i] + width/2, tempos[i] + 0.1,
                    tempo_formatado, ha='center', va='bottom', 
                    color=cor_texto, fontweight='bold', fontsize=8)
        
        # Configurar limites dos eixos
        self.ax.set_ylim(0, max(max(valores), max(atrasos_totais)) * 1.15)
        ax2.set_ylim(0, max(tempos) * 1.2 if tempos else 1)
        
        # Configurar legendas
        lines1, labels1 = self.ax.get_legend_handles_labels()
        lines2, labels2 = ax2.get_legend_handles_labels()
        
        legend = self.ax.legend(lines1 + lines2, labels1 + labels2, 
                            loc='upper left', facecolor=self.cores['fundo_card'],
                            edgecolor=cor_grade, labelcolor=cor_texto)
        legend.get_frame().set_alpha(0.9)
        
        # Configurar bordas do gráfico
        for spine in self.ax.spines.values():
            spine.set_color(cor_grade)
        for spine in ax2.spines.values():
            spine.set_color(cor_grade)
        
        # Ajustar layout
        self.fig.tight_layout()
        
        # Redesenhar
        self.canvas_matplotlib.draw()

    def calcular_valor_periodo(self, periodo):
        """Calcula o valor percentual para o período"""
        try:
            dados_periodo = self.dados_atrasos_periodos.get(periodo, {})
            if not dados_periodo:
                return 0
            
            # Calcular média dos valores de eficiência
            total_atividades = sum(item['quantidade'] for item in dados_periodo.values())
            if total_atividades == 0:
                return 100
            
            return round((1 - (total_atividades / 100)) * 100, 2)
        except Exception as e:
            logger.error(f"Erro ao calcular valor do período {periodo}: {e}")
            return 0

    def calcular_tempo_periodo(self, periodo):
        """Calcula o tempo total de atrasos para o período"""
        try:
            dados_periodo = self.dados_atrasos_periodos.get(periodo, {})
            if not dados_periodo:
                return 0
            
            # Somar todos os tempos de atraso
            return sum(item['tempo'] for item in dados_periodo.values())
        except Exception as e:
            logger.error(f"Erro ao calcular tempo do período {periodo}: {e}")
            return 0

    def calcular_atrasos_periodo(self, periodo):
        """Calcula os atrasos agrupados por motivo para o período"""
        try:
            dados_periodo = self.dados_atrasos_periodos.get(periodo, {})
            if not dados_periodo:
                return {'material': 0, 'equipamento': 0, 'mao_obra': 0}
            
            # Agrupar atrasos por categoria
            atrasos = {
                'material': 0,
                'equipamento': 0,
                'mao_obra': 0
            }
            
            for motivo, dados in dados_periodo.items():
                # Classificar o motivo em uma das categorias
                if 'material' in motivo.lower():
                    atrasos['material'] += dados['quantidade']
                elif 'equip' in motivo.lower():
                    atrasos['equipamento'] += dados['quantidade']
                else:
                    atrasos['mao_obra'] += dados['quantidade']
            
            return atrasos
        except Exception as e:
            logger.error(f"Erro ao calcular atrasos do período {periodo}: {e}")
            return {'material': 0, 'equipamento': 0, 'mao_obra': 0}

    def criar_legenda_unificada(self, frame):
        # Título da legenda
        ctk.CTkLabel(
            frame,
            text="LEGENDAS",
            font=("Arial Black", 14),
            text_color=self.cores['texto_card']
        ).pack(pady=10)
        
        # Frame único para todas as legendas
        legendas_container = ctk.CTkFrame(frame, fg_color="transparent")
        legendas_container.pack(fill="x", padx=10, pady=5)
        
        # Frame para linha única de legendas
        linha_frame = ctk.CTkFrame(legendas_container, fg_color="transparent")
        linha_frame.pack(fill="x")
        
        # Tamanho aumentado para os quadrados de cor
        square_size = 20  # Aumentado de 12 para 20
        
        # Todos os indicadores em uma linha
        for nome, info in self.dados_grafico.items():
            item_frame = ctk.CTkFrame(linha_frame, fg_color="transparent")
            item_frame.pack(side="left", padx=5, fill="x", expand=True)
            
            cor_box = ctk.CTkCanvas(
                item_frame,
                width=square_size,
                height=square_size,
                highlightthickness=0
            )
            cor_box.pack(side="left", padx=2)
            cor_box.create_rectangle(0, 0, square_size, square_size, fill=info['cor'], outline='')
            
            ctk.CTkLabel(
                item_frame,
                text=nome,
                font=("Arial", 15),
                text_color=self.cores['texto_card']
            ).pack(side="left", padx=2)
        
        # Adicionar indicadores de atraso na mesma linha
        indicadores_atraso = {
            'Motivo do Atraso': '#FF6B6B',
            'Tempo de Atraso': '#4ECDC4'
        }
        
        for tipo, cor in indicadores_atraso.items():
            item_frame = ctk.CTkFrame(linha_frame, fg_color="transparent")
            item_frame.pack(side="left", padx=5, fill="x", expand=True)
            
            cor_box = ctk.CTkCanvas(
                item_frame,
                width=square_size,
                height=square_size,
                highlightthickness=0
            )
            cor_box.pack(side="left", padx=2)
            cor_box.create_rectangle(0, 0, square_size, square_size, fill=cor, outline='')
            
            ctk.CTkLabel(
                item_frame,
                text=tipo,
                font=("Arial", 15),
                text_color=self.cores['texto_card']
            ).pack(side="left", padx=2)
        
        # Ajustar cores dos textos da legenda
        for widget in frame.winfo_children():
            if isinstance(widget, ctk.CTkLabel):
                widget.configure(text_color=self.cores['texto_card'])
    
    def formatar_tempo(self, tempo_em_dias):
        """Converte tempo em dias (float ou int) para o formato 'Xd/HH:MMh'"""
        try:
            total_minutos = int(float(tempo_em_dias) * 24 * 60)
            dias = total_minutos // (24 * 60)
            horas = (total_minutos % (24 * 60)) // 60
            minutos = total_minutos % 60
            if dias > 0:
                return f"{dias}d/{horas:02d}:{minutos:02d}h"
            else:
                return f"{horas:02d}:{minutos:02d}h"
        except Exception:
            return str(tempo_em_dias)

    def criar_controle_atrasos(self, parent):
        frame = ctk.CTkFrame(parent, fg_color=self.cores['fundo_card'])
        frame.pack(fill="both", expand=True)
        
        # Título com período
        titulo_frame = ctk.CTkFrame(frame, fg_color="transparent")
        titulo_frame.pack(fill="x", pady=10)
        
        ctk.CTkLabel(
            titulo_frame,
            text="CONTROLE DE ATRASOS",
            font=("Arial Black", 14),
            text_color=self.cores['texto_card']
        ).pack(side="left", padx=10)
        
        ctk.CTkLabel(
            titulo_frame,
            text="(Semana Atual)",
            font=("Arial", 12),
            text_color=self.cores['texto_card']
        ).pack(side="left")
        
        # Container para a Treeview
        lista_container = ctk.CTkFrame(frame, fg_color="transparent")
        lista_container.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # Treeview
        columns = ("Motivo", "Qtd.", "Tempo", "Impacto")
        tree = ttk.Treeview(lista_container, columns=columns, show="headings", height=10)
        tree.heading("Motivo", text="Motivo")
        tree.heading("Qtd.", text="Qtd.")
        tree.heading("Tempo", text="Tempo")
        tree.heading("Impacto", text="Impacto")

        tree.column("Motivo", width=200, anchor="w")
        tree.column("Qtd.", width=60, anchor="center")
        tree.column("Tempo", width=80, anchor="center")
        tree.column("Impacto", width=100, anchor="center")

        # Scrollbar
        scrollbar = ttk.Scrollbar(lista_container, orient="vertical", command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Preencher Treeview apenas com os motivos da semana_atual
        dados_periodo = self.dados_atrasos
        if not dados_periodo or len(dados_periodo) == 0:
            tree.insert("", "end", values=("Nenhum registro de atraso encontrado", "", "", ""))
        else:
            for motivo, dados in dados_periodo.items():
                quantidade = dados.get('quantidade', 0)
                tempo = dados.get('tempo', 0)
                impacto = dados.get('impacto', 'Baixo')
                tree.insert("", "end", values=(motivo, quantidade, self.formatar_tempo(tempo), impacto))

        # Armazenar referência se precisar atualizar depois
        self.treeview_atrasos = tree

    def carregar_controle_atraso(self, dados_atraso, periodo='semana_atual'):
        self.treeview_atrasos.delete(*self.treeview_atrasos.get_children())

        if periodo in dados_atraso:
            for motivo, info in dados_atraso[periodo].items():
                self.treeview_atrasos.insert(
                    "", "end",
                    values=(motivo, info['quantidade'], self.formatar_tempo(info['tempo']), info['impacto'])
                )

    def salvar_relatorio(self):
        """Prepara os dados e salva o relatório PDF"""
        # Coletar dados do dashboard
        dados_dashboard = {
            'equipe': 'SPCS',
            'gerado_por': 'WRP',
            'indicadores': {
                'semana_anterior': '100%',
                'semana_atual': '90%',
                'media_semestral': '20%',
                'media_anual': '0.5%'
            },
            'atrasos': self.dados_atrasos
        }
        
        # Criar instância do printer
        printer = DashboardPrinter(dados_dashboard)
        
        # Abrir diálogo para salvar arquivo
        arquivo = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile=f"dashboard_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        )
        
        if arquivo:
            printer.gerar_relatorio(arquivo)

    def _on_theme_change(self):
        """Atualiza a interface quando o tema muda"""
        try:
            logger.info(f"[THEME] Atualizando interface para tema: {ctk.get_appearance_mode()}")
            
            # Atualizar cores
            self.setup_cores()
            
            # Atualizar container principal
            if hasattr(self, 'container'):
                self.container.configure(fg_color=self.cores['fundo'])
            
            # Atualizar todos os frames e labels
            self._update_widget_colors(self)
            
            # Atualizar gráfico matplotlib
            if hasattr(self, 'canvas_matplotlib') and hasattr(self, 'dados_grafico'):
                self.atualizar_matplotlib_grafico()
            
            # Atualizar lista de atrasos
            if hasattr(self, 'treeview_atrasos'):
                for item in self.treeview_atrasos.get_children():
                    self.treeview_atrasos.delete(item)
                for motivo, dados in self.dados_atrasos.items():
                    quantidade = dados.get('quantidade', 0)
                    tempo = dados.get('tempo', 0)
                    impacto = dados.get('impacto', 'Baixo')
                    self.treeview_atrasos.insert("", "end", values=(motivo, quantidade, self.formatar_tempo(tempo), impacto))
                                
        except Exception as e:
            logger.error(f"[THEME] Erro ao atualizar tema: {e}")
            
    def _update_widget_colors(self, widget):
        """Atualiza recursivamente as cores de todos os widgets"""
        try:
            for child in widget.winfo_children():
                # Atualizar cor do frame
                if isinstance(child, ctk.CTkFrame):
                    if 'transparent' not in str(child.cget('fg_color')):
                        child.configure(fg_color=self.cores['fundo_card'])
                
                # Atualizar cor do label
                elif isinstance(child, ctk.CTkLabel):
                    if not child.cget('image'):
                        texto = child.cget('text')
                        # Preservar cores dos indicadores de tendência
                        if texto in ["▲", "▼"]:
                            cor = self.cores['sucesso'] if texto == "▲" else self.cores['destaque']
                            child.configure(text_color=cor)
                        # Preservar cores dos indicadores de impacto
                        elif texto in ["Alto", "Médio", "Baixo"]:
                            cores_impacto = {
                                'Alto': self.cores['destaque'],    # Vermelho
                                'Médio': '#FFB800',   # Amarelo
                                'Baixo': self.cores['sucesso']    # Verde
                            }
                            child.configure(text_color=cores_impacto[texto])
                        # Verificar se é um label de título ou cabeçalho
                        elif "Arial Black" in str(child.cget('font')):
                            child.configure(text_color=self.cores['texto_card'])
                        else:
                            child.configure(text_color=self.cores['texto'])
                
                # Atualizar cor do botão
                elif isinstance(child, ctk.CTkButton):
                    if child.cget('fg_color') != self.cores['primaria']:
                        child.configure(text_color=self.cores['texto'])
                
                # Atualizar canvas do gráfico
                elif isinstance(child, ctk.CTkCanvas):
                    child.configure(bg=self.cores['fundo_card'])
                    if hasattr(self, 'dados_grafico'):
                        self.desenhar_grafico(self.dados_grafico)
                
                # Recursivamente atualizar widgets filhos
                if child.winfo_children():
                    self._update_widget_colors(child)
                    
        except Exception as e:
            logger.error(f"[THEME] Erro ao atualizar cores do widget: {e}")

    def _on_focus_out(self, event):
        """Chamado quando a janela perde o foco"""
        self.lift()
        self.focus_force()

    def cleanup(self):
        """Limpa todos os recursos antes de destruir a janela"""
        try:
            logger.info("Iniciando limpeza de recursos do DashboardDaily")
            
            # Remover atributo topmost antes de destruir
            self.attributes('-topmost', False)
            
            # Cancelar callbacks pendentes
            if hasattr(self, '_on_theme_change_callback_id'):
                self.after_cancel(self._on_theme_change_callback_id)
                
            # Remover callback de tema
            ctk.AppearanceModeTracker.remove(self._on_theme_change_callback)
            
            # Limpar matplotlib
            if hasattr(self, 'fig'):
                plt.close(self.fig)
            if hasattr(self, 'canvas_matplotlib'):
                self.canvas_matplotlib.get_tk_widget().destroy()
            
            # Limpar referências cíclicas
            self.parent = None
            self.db = None
            self.user_data = None
            
            # Limpar dados
            self.dados_atrasos.clear()
            self.dados_atrasos_periodos.clear()
            
            # Limpar imagens
            for img in self.images.values():
                img = None
            self.images.clear()
            
            # Forçar coleta de lixo
            import gc
            gc.collect()
            
            logger.info("Limpeza de recursos do DashboardDaily concluída")
            
        except Exception as e:
            logger.error(f"Erro durante limpeza de recursos: {e}", exc_info=True)

    def destroy(self):
        """Sobrescreve o método destroy para garantir limpeza de recursos"""
        try:
            self.cleanup()
        finally:
            super().destroy()

    def __del__(self):
        """Destrutor da classe"""
        try:
            # Limpar referências das imagens
            self.images.clear()
            
            # Remover callback ao destruir a janela
            if hasattr(self, '_on_theme_change_callback_id'):
                self.after_cancel(self._on_theme_change_callback_id)
            ctk.AppearanceModeTracker.remove(self._on_theme_change_callback)
            logger.debug("[THEME] Callback de tema removido")
        except Exception as e:
            logger.error(f"[THEME] Erro ao limpar recursos: {e}")