from customtkinter import CTkFrame
import customtkinter as ctk
from tkinter import ttk, messagebox
import tkinter as tk
import os, sys
import logging
from PIL import Image
from datetime import datetime, timedelta
from ....database.connection import DatabaseConnection
from ....utils.excel_selector import ExcelSelector

logger = logging.getLogger(__name__)

class ActivitiesTab(CTkFrame):
    def __init__(self, parent, user_data, manager=None):
        super().__init__(parent)
        self.user_data = user_data
        self.manager = manager
        self.db = DatabaseConnection()
        self.setup_activities()
        self.pack(expand=True, fill="both")

    def setup_activities(self):
        """Configura a aba de atividades"""
        # Título da aba com estatísticas
        title_frame = ctk.CTkFrame(self, fg_color="transparent")
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

        activities_frame = ctk.CTkFrame(self, fg_color="transparent")
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
            # Corrigido: busca o ícone na pasta raiz 'icons/'
            python_icon_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), '..', '..', '..', '..', 'icons', 'reset.png')
            python_icon_path = os.path.abspath(python_icon_path)
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
                    equipes_raw = self.manager.get_equipes() if self.manager else []
                    equipes = [e[0] if isinstance(e, (tuple, list)) else str(e) for e in equipes_raw]
                    widget = ctk.CTkComboBox(
                        field_frame,
                        values=equipes,
                        height=28
                    )
                    if equipes:
                        widget.set(equipes[0])
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
            python_icon_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), '..', '..', '..', '..', 'icons', 'python.png')
            python_icon_path = os.path.abspath(python_icon_path)
            if not os.path.exists(python_icon_path):
                python_icon_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), '..', '..', '..', '..', 'icons', 'logo_login_light.png')
                python_icon_path = os.path.abspath(python_icon_path)
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

    def open_excel_selector(self):
        """Abre o seletor de Excel para escolher a descrição"""
        def on_select(description):
            if description and "descricao" in self.activity_entries:
                self.activity_entries["descricao"].delete(0, "end")
                self.activity_entries["descricao"].insert(0, description)
        
        ExcelSelector(self, on_select)

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
            equipes = self.manager.get_equipes() if self.manager else []
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