import customtkinter as ctk

from ...database.connection import DatabaseConnection
from ..components.activities.activity_form import ActivityForm
from ..components.logic.activity_controls_logic import ActivityControlsLogic
from ..notifications.notification_manager import NotificationManager
from ...core.time.time_observer import TimeObserver, Dict, Optional
from ...core.time.time_manager import TimeManager
from ...core.time.daily_time_manager import DailyTimeManager
from datetime import datetime, timedelta
import logging

logger = logging.getLogger(__name__)

class ActivityTopFrame(ctk.CTkFrame, TimeObserver):
    def __init__(self, parent, user_data):
        super().__init__(parent)
        self.db = DatabaseConnection()
        self.user_data = user_data
        self.logic = ActivityControlsLogic(self.db)
        self.notification_manager = NotificationManager()
        self.notification_manager.initialize(self, self.user_data['nome'])

        # Inicialização dos atributos do timer
        self.time_controller = None
        self._timer_id = None
        self.active_activity = None
        self.selected_activity = None
        self.time_manager = TimeManager()
        self.time_manager.state.set_user_id(user_data['id'])
        self.time_manager.add_observer(self)
        self.daily_time_manager = DailyTimeManager()
        self.daily_time_manager.add_observer(self)

        # Frame do relógio
        clock_frame = ctk.CTkFrame(self)
        clock_frame.pack(side="left", padx=10, pady=5)

        # Título do frame do relógio
        clock_title = ctk.CTkLabel(
            clock_frame,
            text="HORA/DATA",
            font=("Roboto", 12, "bold"),
            text_color="#ff5722"  # Cor laranja para destaque
        )
        clock_title.pack(pady=(5,2))

        # Relógio
        self.clock_label = ctk.CTkLabel(
            clock_frame,
            text="00:00:00",
            font=("Roboto", 24, "bold"),
        )
        self.clock_label.pack(pady=(0,2))

        # Data
        self.date_label = ctk.CTkLabel(
            clock_frame,
            text="DD/MM/YYYY",
            font=("Roboto", 12),
        )
        self.date_label.pack(pady=(0,5))

        # Frame do tempo diário
        daily_hours_frame = ctk.CTkFrame(self)
        daily_hours_frame.pack(side="left", padx=10, pady=5)

        # Título do frame de tempo diário
        daily_title = ctk.CTkLabel(
            daily_hours_frame,
            text="TEMPO DIÁRIO",
            font=("Roboto", 12, "bold"),
            text_color="#ff5722"
        )
        daily_title.pack(pady=(5,2))

        # Timer diário
        self.daily_hours_label = ctk.CTkLabel(
            daily_hours_frame,
            text="00:00:00",
            font=("Roboto", 24, "bold"),
        )
        self.daily_hours_label.pack(pady=(0,2))

        # Valor decimal
        self.decimal_hours_label = ctk.CTkLabel(
            daily_hours_frame,
            text="0.000",
            font=("Roboto", 12),
        )
        self.decimal_hours_label.pack(pady=(0,5))

        # Frame dos horários da empresa
        company_hours_frame = ctk.CTkFrame(self)  # Removida largura fixa
        company_hours_frame.pack(side="left", padx=10, pady=5)

        # Título do frame dos horários
        company_hours_title = ctk.CTkLabel(
            company_hours_frame,
            text="HORÁRIOS",
            font=("Roboto", 12, "bold"),
            text_color="#ff5722"
        )
        company_hours_title.pack(pady=(5,2))

        # Horários da empresa
        company_hours = {
            'start': TimeManager.get_time_tuple(TimeManager.COMPANY_START_TIME),
            'break_start': TimeManager.get_time_tuple(TimeManager.BREAK_START_TIME),
            'break_end': TimeManager.get_time_tuple(TimeManager.BREAK_END_TIME),
            'end': TimeManager.get_time_tuple(TimeManager.COMPANY_END_TIME)
        }

        # Container para os horários
        hours_container = ctk.CTkFrame(company_hours_frame, fg_color="transparent")
        hours_container.pack(fill="x", padx=5, pady=(0,5))

        # Label para expediente
        expediente_container = ctk.CTkFrame(hours_container, fg_color="transparent")
        expediente_container.pack(fill="x", pady=2)

        expediente_label = ctk.CTkLabel(
            expediente_container,
            text="Expediente:",
            font=("Roboto", 12, "bold"),
        )
        expediente_label.pack(side="left")

        expediente_time_label = ctk.CTkLabel(
            expediente_container,
            text=f"{company_hours['start'][0]:02d}:{company_hours['start'][1]:02d} - {company_hours['end'][0]:02d}:{company_hours['end'][1]:02d}",
            font=("Roboto", 12),
        )
        expediente_time_label.pack(side="left", padx=(5,0))

        # Label para intervalo
        intervalo_container = ctk.CTkFrame(hours_container, fg_color="transparent")
        intervalo_container.pack(fill="x", pady=2)

        intervalo_label = ctk.CTkLabel(
            intervalo_container,
            text="Intervalo:",
            font=("Roboto", 12, "bold"),
        )
        intervalo_label.pack(side="left")

        intervalo_time_label = ctk.CTkLabel(
            intervalo_container,
            text=f"{company_hours['break_start'][0]:02d}:{company_hours['break_start'][1]:02d} - {company_hours['break_end'][0]:02d}:{company_hours['break_end'][1]:02d}",
            font=("Roboto", 12),
        )
        intervalo_time_label.pack(side="left", padx=(5,0))

        # Frame de atividades
        info_frame = ctk.CTkFrame(self, height=50)  # Altura fixa
        info_frame.pack(side="left", expand=True, fill="both", padx=10, pady=5)
        info_frame.pack_propagate(False)  # Bloqueia expansão

        # Título do frame de atividades
        activity_title = ctk.CTkLabel(
            info_frame,
            text="ATIVIDADE",
            font=("Roboto", 12, "bold"),
            text_color="#ff5722"
        )
        activity_title.pack(pady=(5,2))

        # Labels de atividade
        self.active_activity_label = ctk.CTkLabel(
            info_frame,
            text="Atividade atual: Nenhuma",
            font=("Roboto", 12, "bold"),
            anchor="w"
        )
        self.active_activity_label.pack(fill="x", padx=5, pady=(0, 2))

        self.selected_activity_label = ctk.CTkLabel(
            info_frame,
            text="Atividade selecionada: Nenhuma",
            font=("Roboto", 12),
            anchor="w"
        )
        self.selected_activity_label.pack(fill="x", padx=5)

        # Frame do timer da atividade
        timer_frame = ctk.CTkFrame(self)
        timer_frame.pack(side="right", padx=10)

        # Título do frame do timer
        timer_title = ctk.CTkLabel(
            timer_frame,
            text="TEMPO ATIVIDADE",
            font=("Roboto", 12, "bold"),
            text_color="#ff5722"
        )
        timer_title.pack(pady=(5,2))

        # Timer
        self.timer_label = ctk.CTkLabel(
            timer_frame,
            text="00:00:00",
            font=("Roboto", 20, "bold"),
        )
        self.timer_label.pack(pady=(0,2))

        # Tempo total
        self.total_time_label = ctk.CTkLabel(
            timer_frame,
            text="Total: 00:00:00",
            font=("Roboto", 12),
        )
        self.total_time_label.pack(pady=(0,5))

        # Iniciar atualização do relógio
        self.update_clock()

    def handle_activity_action(self, action):
        """Redireciona ações de atividade para o ActivityControls"""
        if self.activity_controls:
            self.activity_controls.handle_activity_action(action)

    def show_activity_form(self):
        """Exibe o formulário de criação de atividade"""
        ActivityForm(self, self.user_data, self.activity_controls.on_activity_created)

    def on_activity_created(self):
        """Callback chamado quando uma atividade é criada"""
        if hasattr(self, 'activity_table'):
            self.activity_table.update_activities()
        if hasattr(self, 'activity_controls'):
            self.activity_controls.update_button_states(True)
        
        # Reset timer and total time
        self.timer_label.configure(text="00:00:00")
        self.total_time_label.configure(text="Total: 00:00:00")

    def update_activity_status(self, activity_info: Optional[Dict]) -> None:
        """Atualiza o status da atividade e notifica se necessário"""
        try:
            if activity_info:
                status = activity_info.get('status', '').upper()
                self.active_activity_label.configure(
                    text=f"Atividade atual ({status}): {activity_info['atividade']}"
                )
                
                # Verificar horário comercial usando o notification_manager
                time_status = self.notification_manager.check_company_hours()
                
                if time_status != "working_hours":
                    self.notification_manager.notify_company_hours(time_status, activity_info)
            else:
                self.active_activity_label.configure(text="Atividade atual: Nenhuma")
                
        except Exception as e:
            logger.error(f"Erro ao atualizar status da atividade: {e}")

    def notify_time_exceeded(self, activity_info: Dict) -> None:
        """Notifica quando o tempo é excedido"""
        try:
            # Apenas notificar através do notification manager
            self.notification_manager.notify_time_exceeded(activity_info)
        except Exception as e:
            logger.error(f"Erro ao processar notificação de tempo excedido: {e}")

    def update_timer_display(self, timer_value: timedelta, total_time: timedelta):
        """Atualiza o display do timer"""
        try:
            # Verificar se os widgets ainda existem
            if not self.winfo_exists():
                return
            
            if hasattr(self, 'timer_label') and self.timer_label.winfo_exists():
                timer_str = self.time_manager.format_total_time(timer_value)
                self.timer_label.configure(text=timer_str)
            
            if hasattr(self, 'total_time_label') and self.total_time_label.winfo_exists():
                total_str = self.time_manager.format_total_time(total_time)
                self.total_time_label.configure(text=total_str)
            
        except Exception as e:
            logger.error(f"Erro ao atualizar display do timer: {e}")

    def update_clock(self):
        current_time = datetime.now()
        self.clock_label.configure(text=current_time.strftime("%H:%M:%S"))
        self.date_label.configure(text=current_time.strftime("%d/%m/%Y"))
        
        # Verificar mudança de dia
        self.daily_time_manager.check_day_change()
        
        # Verificar horários comerciais
        time_status = self.notification_manager.check_company_hours()
        if time_status != "working_hours":
            # Buscar atividade ativa se houver
            query = """
                SELECT id, atividade 
                FROM atividades 
                WHERE user_id = %s
                AND ativo = TRUE 
                AND concluido = FALSE
                AND pausado = FALSE
                LIMIT 1
            """
            result = self.db.execute_query(query, (self.user_data['id'],))
            activity_info = result[0] if result else None
            
            # Notificar sobre horário comercial
            self.notification_manager.notify_company_hours(time_status, activity_info)
        
        current_time = current_time.time()
        company_end = TimeManager.get_time_object(TimeManager.COMPANY_END_TIME)
        
        try:
            # Verificar horário de fim do expediente
            if current_time >= company_end:
                query = """
                    SELECT id, atividade 
                    FROM atividades 
                    WHERE user_id = %s
                    AND ativo = TRUE 
                    AND concluido = FALSE
                    AND pausado = FALSE
                    LIMIT 1
                """
                result = self.db.execute_query(query, (self.user_data['id'],))
                
                if result and result[0]:
                    activity_id = result[0]['id']
                    activity_name = result[0]['atividade']
                    
                    # Pausa a atividade diretamente
                    update_query = """
                        UPDATE atividades 
                        SET pausado = TRUE,
                            ativo = TRUE,
                            concluido = FALSE,
                            total_time = COALESCE(
                                (SELECT SEC_TO_TIME(
                                    TIME_TO_SEC(TIMEDIFF(NOW(), start_time))
                                )
                            ), '00:00:00')
                        WHERE id = %s
                    """
                    self.db.execute_query(update_query, (activity_id,))
                    
                    # Pausa o timer sem mensagens extras
                    self.time_manager.pause_activity()
                    
                    # Atualiza interface
                    if hasattr(self, 'activity_controls'):
                        self.activity_controls.check_current_status()
                        self.activity_controls.refresh_activities()
                        
                    logger.debug(f"[END_TIME] Atividade '{activity_name}' pausada automaticamente por fim de expediente")
            
        except Exception as e:
            logger.error(f"Erro ao processar pausa automática: {e}")
        
        # Atualizar tempo diário se estiver rodando
        if self.daily_time_manager.is_running:
            self.daily_time_manager.update_daily_hours()
        
        self.after(1000, self.update_clock)

    def _pause_active_activities(self):
        """Pausa todas as atividades ativas"""
        try:
            logger.debug("[LOCK] Iniciando pausa de atividades ativas")
            
            # Buscar atividades ativas
            query = """
                SELECT id, atividade 
                FROM atividades 
                WHERE user_id = %s
                AND ativo = TRUE 
                AND concluido = FALSE
                AND pausado = FALSE
            """
            active_activities = self.db.execute_query(query, (self.user_data['id'],))
            
            for activity in active_activities:
                logger.debug(f"[LOCK] Pausando atividade {activity['id']}: {activity['atividade']}")
                
                # Pausar no TimeManager
                if hasattr(self, 'time_manager'):
                    self.time_manager.pause_activity()
                
                # Atualizar status no banco
                update_query = """
                    UPDATE atividades 
                    SET pausado = TRUE,
                        ativo = TRUE,
                        concluido = FALSE
                    WHERE id = %s
                """
                self.db.execute_query(update_query, (activity['id'],))
                
                # Notificar interface
                if hasattr(self, 'activity_controls'):
                    self.activity_controls.check_current_status()
                    self.activity_controls.refresh_activities()
                    
            logger.info("[LOCK] Todas as atividades ativas foram pausadas")
            
        except Exception as e:
            logger.error(f"Erro ao pausar atividades ativas: {e}")

    def _has_active_activities(self):
        """Verifica se existem atividades ativas"""
        try:
            query = """
                SELECT COUNT(*) as count
                FROM atividades 
                WHERE user_id = %s
                AND ativo = TRUE 
                AND concluido = FALSE
                AND pausado = FALSE
            """
            result = self.db.execute_query(query, (self.user_data['id'],))
            return result[0]['count'] > 0 if result else False
        except Exception as e:
            logger.error(f"Erro ao verificar atividades ativas: {e}")
            return False
        
    def _convert_time_to_decimal(self, time_value) -> str:
        """
        Converte tempo para decimal, usando o mesmo padrão do ExcelProcessor
        """
        try:
            if not time_value:
                return "0,0000"

            # Se for timedelta, extrair horas, minutos e segundos
            if isinstance(time_value, timedelta):
                total_seconds = int(time_value.total_seconds())
                hours = total_seconds // 3600
                minutes = (total_seconds % 3600) // 60
                seconds = total_seconds % 60
            # Se for string, fazer split
            elif isinstance(time_value, str):
                if time_value == "00:00:00":
                    return "0,0000"
                hours, minutes, seconds = map(int, time_value.split(':'))
            else:
                logger.error(f"Tipo de tempo não suportado: {type(time_value)}")
                return "0,0000"

            decimal_hours = hours + (minutes / 60.0) + (seconds / 3600.0)
            # Formata para ter 4 casas decimais e usa vírgula
            return f"{decimal_hours:.4f}".replace('.', ',')
                
        except Exception as e:
            logger.error(f"Erro ao converter tempo {time_value}: {e}")
            return "0,0000"

    def update_daily_time(self, daily_time: timedelta) -> None:
        """Atualiza o display do tempo diário e decimal"""
        try:
            total_seconds = int(daily_time.total_seconds())
            hours = total_seconds // 3600
            minutes = (total_seconds % 3600) // 60
            seconds = total_seconds % 60
            
            # Atualizar o label de tempo em HH:MM:SS
            self.daily_hours_label.configure(text=f"{hours:02d}:{minutes:02d}:{seconds:02d}")
            
            # Converter para decimal e atualizar o label
            decimal_value = self._convert_time_to_decimal(daily_time)
            self.decimal_hours_label.configure(text=decimal_value)
            
        except Exception as e:
            logger.error(f"Erro ao atualizar display do tempo diário: {e}")

    def update_active_activity(self, activity_info=None):
        """Atualiza informações da atividade ativa"""
        try:
            self.active_activity = activity_info
            if activity_info:
                status = "ATIVA" if activity_info.get('ativo') else "PAUSADA"
                self.active_activity_label.configure(
                    text=f"Atividade atual ({status}): {activity_info['atividade']}"
                )
            else:
                self.active_activity_label.configure(text="Atividade atual: Nenhuma")
        except Exception as e:
            logger.error(f"Erro ao atualizar atividade ativa: {e}")

    def check_current_status(self):
        """Verifica e atualiza status atual"""
        try:
            query = """
                SELECT id, description, atividade, start_time, end_time, 
                    ativo, pausado, concluido
                FROM atividades 
                WHERE user_id = %s
                AND ativo = TRUE 
                AND concluido = FALSE
                AND pausado = FALSE
                LIMIT 1
            """
            
            def on_status_received(result):
                if result and result[0]:
                    activity = result[0]
                    self.active_activity = activity
                    if self.active_activity_label:
                        self.active_activity_label.configure(
                            text=f"Atividade atual (ATIVA): {activity['atividade']}"
                        )
                else:
                    self.active_activity = None
                    if self.active_activity_label:
                        self.active_activity_label.configure(text="Atividade atual: Nenhuma")
                self._update_button_states("atividades")
            
            # Usar versão assíncrona
            self.logic.db.execute_query_async(
                query, 
                (self.user_data['id'],),
                callback=on_status_received
            )
            
        except Exception as e:
            logger.error(f"Erro ao verificar status atual: {e}")
            if self.active_activity_label:
                self.active_activity_label.configure(text="Atividade atual: Erro ao carregar")
            self.active_activity = None

    def update_idle_status(self, status: str):
        """
        Atualiza o status de ociosidade do usuário.
        
        Args:
            status: 'idle' para inativo, 'active' para ativo
        """
        try:
            if status == 'idle':
                # Lógica para quando o usuário está inativo
                logger.debug("Usuário está inativo (UserWindow)")
            else:
                # Lógica para quando o usuário retoma a atividade
                logger.debug("Usuário está ativo (UserWindow)")
                
        except Exception as e:
            logger.error(f"Erro ao atualizar status de ociosidade: {e}")