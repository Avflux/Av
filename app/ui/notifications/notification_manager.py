import tkinter as tk
import logging, os, sys
from tkinter import messagebox
from datetime import datetime, time
from winotify import Notification, audio
from typing import Dict
from ...database.connection import DatabaseConnection
from ...config.settings import APP_CONFIG

logger = logging.getLogger(__name__)

class TimeConverter:
    """Classe utilitária para converter entre diferentes formatos de tempo"""
    
    @staticmethod
    def string_to_time(time_str: str) -> time:
        """Converte string no formato HH:MM:SS para objeto time"""
        try:
            h, m, s = map(int, time_str.split(':'))
            return time(h, m, s)
        except ValueError:
            logger.error(f"Erro ao converter string de tempo: {time_str}")
            return None
            
    @staticmethod
    def tuple_to_time(time_tuple: tuple) -> time:
        """Converte tupla (hora, minuto) para objeto time"""
        try:
            return time(time_tuple[0], time_tuple[1])
        except (IndexError, TypeError):
            logger.error(f"Erro ao converter tupla de tempo: {time_tuple}")
            return None

class MessageBoxManager:
    """Gerencia quase todas as caixas de diálogo do sistema"""
    
    @staticmethod
    def confirm_action(action: str) -> bool:
        """Confirma uma ação com o usuário"""
        return messagebox.askyesno("Confirmar", f"Deseja realmente {action} esta atividade?")
    
    @staticmethod
    def confirm_break_time() -> bool:
        """Confirma ação durante horário de intervalo"""
        return messagebox.askyesno(
            "Horário de Intervalo",
            "Você está no horário de intervalo (12:15 - 13:15).\n"
            "Deseja continuar mesmo assim?"
        )
    
    @staticmethod
    def show_error(title: str, message: str) -> None:
        """Exibe mensagem de erro"""
        messagebox.showerror(title, message)
    
    @staticmethod
    def show_warning(title: str, message: str) -> None:
        """Exibe mensagem de aviso"""
        messagebox.showwarning(title, message)
    
    @staticmethod
    def show_info(title: str, message: str) -> None:
        """Exibe mensagem informativa"""
        messagebox.showinfo(title, message)
    
    @staticmethod
    def show_active_activity_warning() -> None:
        """Aviso de atividade já em andamento"""
        messagebox.showwarning(
            "Aviso",
            "Você já possui uma atividade em andamento!"
        )

    @staticmethod
    def show_select_activity_warning() -> None:
        """Aviso para selecionar uma atividade"""
        messagebox.showwarning("Aviso", "Selecione uma atividade primeiro!")

    @staticmethod
    def show_break_suggestions() -> None:
        """Mostra sugestões para o intervalo"""
        suggestions = (
            "🧘‍♂️ Faça alguns exercícios de alongamento\n"
            "🚶‍♂️ Dê uma curta caminhada\n"
            "👀 Pratique exercícios para os olhos\n"
            "🌱 Aproveite para se hidratar\n"
            "🧘‍♀️ Pratique alguns minutos de respiração consciente"
        )
        messagebox.showinfo("Sugestões para o Intervalo", suggestions)

    @staticmethod
    def show_day_summary() -> None:
        """Mostra resumo do dia"""
        messagebox.showinfo(
            "Resumo do Dia",
            "Obrigado pelo seu trabalho hoje!\n"
            "Todas as suas atividades foram registradas com sucesso."
        )

    @staticmethod
    def confirm_day_end() -> bool:
        """Confirma encerramento do dia"""
        return messagebox.askyesno(
            "Encerrar Dia",
            "Deseja realmente encerrar o dia?\n"
            "Isso irá finalizar todas as atividades em andamento."
        )

class NotificationManager:
    _instance = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(NotificationManager, cls).__new__(cls)
            cls._instance.main_window = None
            cls._instance.username = None
            cls._instance.message_box = MessageBoxManager()
            cls._instance._notification_states = {
                "before_hours": False,
                "break_time": False,
                "break_end": False,
                "after_hours": False
            }
            cls._instance.last_checked_date = None
            
            # Definir horários fixos da empresa (substituindo TimeManager)
            time_converter = TimeConverter()
            cls._instance.COMPANY_START = time_converter.string_to_time('08:00:00')  # Início expediente
            cls._instance.BREAK_START = time_converter.string_to_time('12:15:00')   # Início intervalo
            cls._instance.BREAK_END = time_converter.string_to_time('13:15:00')     # Fim intervalo
            cls._instance.COMPANY_END = time_converter.string_to_time('17:00:00')   # Fim expediente
            
        return cls._instance

    def initialize(self, main_window, username):
        """Inicializa o gerenciador com a janela principal e nome do usuário"""
        try:
            self.main_window = main_window
            self.username = username
            self.db = DatabaseConnection()
            
            # Resetar estados de notificação
            self._notification_states = {
                "before_hours": False,
                "break_time": False,
                "break_end": False,
                "after_hours": False
            }
            
            # Verificar horário de login para mensagem de bom dia
            current_hour = datetime.now().hour
            if 0 <= current_hour < 8:
                message = (
                    f"Bom dia, {self.username}! Você está começando cedo hoje.\n"
                    "Que seu dia seja produtivo e cheio de realizações!"
                )
                self.show_system_notification(
                    title="Início do Expediente 🌅",
                    message=message,
                    icon_path=APP_CONFIG['icons']['morning']
                )
                logger.debug(f"Mensagem de início de expediente exibida às {current_hour}h")
            
        except Exception as e:
            logger.error(f"Erro ao inicializar notification manager: {e}")

    def show_system_notification(self, title, message, icon_path=None):
        """Exibe uma notificação do sistema usando winotify"""
        try:
            # Verificar se o PowerShell está disponível
            import subprocess
            try:
                # Usando CREATE_NO_WINDOW para evitar que a janela apareça
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = subprocess.SW_HIDE
                
                subprocess.run(
                    ['powershell', '-Command', 'exit 0'],
                    capture_output=True,
                    check=True,
                    startupinfo=startupinfo
                )
            except (subprocess.SubprocessError, FileNotFoundError):
                # Se PowerShell não estiver disponível, usar messagebox como fallback
                logger.warning("PowerShell não disponível, usando messagebox como fallback")
                self.main_window.after(1, lambda: messagebox.showinfo(title, message))
                return

            # Resto do código original para notificações
            if icon_path:
                if not os.path.isabs(icon_path):
                    if hasattr(sys, "_MEIPASS"):
                        icon = os.path.join(sys._MEIPASS, icon_path)
                    else:
                        base_path = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
                        icon = os.path.join(base_path, icon_path)
            else:
                if hasattr(sys, "_MEIPASS"):
                    icon = os.path.join(sys._MEIPASS, APP_CONFIG['icons']['app'])
                else:
                    base_path = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
                    icon = os.path.join(base_path, APP_CONFIG['icons']['app'])

            if not os.path.exists(icon):
                logger.warning(f"Ícone não encontrado: {icon}")
                icon = ""

            app_name = "Sistema Chronos.exe" if hasattr(sys, "_MEIPASS") else "Sistema Chronos"
            toast = Notification(
                app_id=app_name,
                title=title,
                msg=message,
                icon=icon,
                duration="short"
            )

            toast.set_audio(audio.Default, loop=False)
            toast.show()
            logger.info(f"Notificação enviada: {title}")

        except Exception as e:
            logger.error(f"Erro ao mostrar notificação do sistema: {e}", exc_info=True)
            # Usar messagebox como fallback
            self.main_window.after(1, lambda: messagebox.showinfo(title, message))

    #---------------------Horário Comercial--------------------------   
    def notify_company_hours(self, status, active_activity=None):
        """Notifica sobre horários da empresa usando mensagens dinâmicas"""
        try:
            logger.debug(f"Notificando horário comercial - Status: {status}, Atividade: {active_activity}")
            
            messages = {
                "before_hours": {
                    "title": "Início Antecipado ⏰",
                    "msg": f"Atividade iniciada antes do horário comercial ({self.COMPANY_START.strftime('%H:%M')})",
                    "icon": APP_CONFIG['icons']['time_exceeded']
                },
                "break_time": {
                    "title": "Horário de Intervalo 🍽️",
                    "msg": (
                        f"Você está no horário de intervalo ({self.BREAK_START.strftime('%H:%M')} - "
                        f"{self.BREAK_END.strftime('%H:%M')})\n"
                        "Aproveite para descansar e se alimentar bem!"
                    ),
                    "icon": APP_CONFIG['icons']['break']
                },
                "break_end": {
                    "title": "Retorno do Intervalo ⏰",
                    "msg": (
                        "Bem-vindo de volta!\n"
                        "Hora de retomar as atividades com energia renovada."
                    ),
                    "icon": APP_CONFIG['icons']['success']
                },
                "after_hours": {
                    "title": "Fim do Expediente 🌙",
                    "msg": (
                        f"O expediente da empresa se encerrou às {self.COMPANY_END.strftime('%H:%M')}\n"
                        "Não se esqueça de concluir suas atividades!"
                    ),
                    "icon": APP_CONFIG['icons']['night']
                }
            }
            
            if status in messages:
                notification = messages[status]
                logger.debug(f"Enviando notificação: {notification['title']} - {notification['msg']}")
                
                if not self._notification_states[status]:
                    self.show_system_notification(
                        title=notification["title"],
                        message=notification["msg"],
                        icon_path=notification.get("icon", APP_CONFIG['icons']['app'])
                    )
                    self._notification_states[status] = True
                    logger.info(f"Notificação de horário comercial enviada: {status}")
                
            return True
            
        except Exception as e:
            logger.error(f"Erro ao notificar horário comercial: {e}", exc_info=True)
            return False

    def reset_notification_states(self):
        """Reseta os estados das notificações no início de um novo dia"""
        self._notification_states = {
            "before_hours": False,
            "break_time": False,
            "break_end": False,
            "after_hours": False
        }

    def check_company_hours(self):
        """Verifica e notifica sobre horários da empresa usando horários fixos definidos na classe"""
        try:
            current_time = datetime.now().time()
            current_date = datetime.now().date()

            # Resetar estados de notificação no início de um novo dia
            if current_date != self.last_checked_date:
                self.reset_notification_states()
                self.last_checked_date = current_date

            # Verificar início do expediente
            if current_time < self.COMPANY_START:
                return "before_hours"
                
            # Verificar intervalo
            elif self.BREAK_START <= current_time <= self.BREAK_END:
                # Se estiver no início do intervalo (dentro de 1 minuto)
                if abs((datetime.combine(datetime.today(), current_time) - 
                       datetime.combine(datetime.today(), self.BREAK_START)).total_seconds()) <= 60:
                    return "break_time"
                # Se estiver no fim do intervalo (dentro de 1 minuto)
                elif abs((datetime.combine(datetime.today(), current_time) - 
                        datetime.combine(datetime.today(), self.BREAK_END)).total_seconds()) <= 60:
                    return "break_end"
                return "break_time"
                
            # Verificar fim do expediente
            elif current_time >= self.COMPANY_END:
                return "after_hours"
                
            return "working_hours"
            
        except Exception as e:
            logger.error(f"Erro ao verificar horário comercial: {e}")
            return "working_hours"

    #------------------Atividades-----------------
    def notify_time_exceeded(self, activity_info: Dict) -> None:
        """Notifica quando o tempo é excedido"""
        try:
            if not activity_info:
                return
            
            # Verificar se já foi notificado
            query = "SELECT notification_shown FROM atividades WHERE id = %s"
            result = self.db.execute_query(query, (activity_info['id'],))
            
            if result and not result[0]['notification_shown']:
                title = "⚠️ Tempo Excedido"
                message = (
                    f"A atividade '{activity_info.get('atividade', '')}' "
                    "excedeu o tempo previsto!"
                )
                
                # Mostrar notificação
                self.show_system_notification(
                    title=title,
                    message=message,
                    icon_path=APP_CONFIG['icons']['time_exceeded']
                )
                
                # Marcar como notificado
                update_query = "UPDATE atividades SET notification_shown = TRUE WHERE id = %s"
                self.db.execute_query(update_query, (activity_info['id'],))
                
                logger.info(f"Notificação de tempo excedido enviada para atividade {activity_info.get('id')}")
            
        except Exception as e:
            logger.error(f"Erro ao notificar tempo excedido: {e}")
    
    def _handle_time_exceeded_click(self, activity_info):
        """Manipula o clique na notificação de tempo excedido"""
        try:
            # Traz a janela principal para frente
            if self.main_window:
                self.main_window.deiconify()
                self.main_window.lift()
                self.main_window.focus_force()
                
                # Se tiver uma tabela de atividades, seleciona a atividade excedida
                if hasattr(self.main_window, 'activity_table'):
                    self.main_window.activity_table.select_activity(activity_info['id'])
                    
        except Exception as e:
            logger.error(f"Erro ao manipular clique em notificação de tempo excedido: {e}")

    def notify_day_end(self):
        """Notifica o encerramento do dia"""
        self.show_system_notification(
            title="Encerramento do Dia",
            message="O dia de trabalho está sendo encerrado. Clique para verificar.",
            icon_path=APP_CONFIG['icons']['time_exceeded']
        )

    def confirm_day_end(self):
        """Confirma o encerramento do dia"""
        return self.message_box.confirm_day_end()

    #---------------Eventos do Sistema------------
    def notify_system_event(self, event_type):
        """Notifica eventos do sistema"""
        try:
            messages = {
                "lembrete_agua": {
                    "title": "Hora de se Hidratar! 💧",
                    "msg": (
                        "Mantenha-se hidratado para manter a produtividade!\n"
                        "Uma pausa rápida para beber água faz toda diferença."
                    ),
                    "icon": APP_CONFIG['icons']['water']
                }
            }

            if event_type in messages:
                notification = messages[event_type]
                self.show_system_notification(
                    title=notification["title"],
                    message=notification["msg"],
                    icon_path=notification.get("icon", APP_CONFIG['icons']['app'])
                )
                logger.info(f"Evento do sistema notificado: {event_type}")

        except Exception as e:
            logger.error(f"Erro ao notificar evento do sistema: {e}", exc_info=True)
    
    def _handle_system_event_click(self, event_type):
        """Manipula o clique em notificações do sistema"""
        try:
            if self.main_window:
                self.main_window.deiconify()
                self.main_window.lift()
                self.main_window.focus_force()
                
                # Ações específicas baseadas no tipo de evento
                if event_type == "intervalo":
                    # Poderia abrir uma janela com sugestões de alongamentos
                    self._show_break_suggestions()
                elif event_type == "encerramento":
                    # Poderia mostrar um resumo do dia
                    self._show_day_summary()
                
        except Exception as e:
            logger.error(f"Erro ao manipular clique em notificação do sistema: {e}")
    
    def _show_break_suggestions(self):
        """Mostra sugestões para o intervalo"""
        self.message_box.show_break_suggestions()
    
    def _show_day_summary(self):
        """Mostra um resumo do dia de trabalho"""
        self.message_box.show_day_summary()
    
    def schedule_water_reminder(self, window):
        """Agenda lembretes para beber água começando 1h após login"""
        try:
            # Agenda o primeiro lembrete para 1h após o login
            window.after(3600000, self._show_water_reminder, window)
            logger.debug("Lembrete de água agendado para 1h após login")
            
        except Exception as e:
            logger.error(f"Erro ao agendar lembrete de água: {e}")
        
    def _show_water_reminder(self, window):
        """Mostra o lembrete de água e agenda o próximo"""
        try:
            # Envia notificação
            self.notify_system_event("lembrete_agua")
            
            # Agenda próximo lembrete em 1 hora (3600000 milliseconds)
            window.after(3600000, self._show_water_reminder, window)
            
            logger.debug("Lembrete de água exibido e próximo agendado")
            
        except Exception as e:
            logger.error(f"Erro ao mostrar lembrete de água: {e}")
    
    def show_welcome_message(self, username):
        """Mostra mensagem de saudação personalizada baseada no horário"""
        try:
            current_hour = datetime.now().hour
            
            # Definir saudação baseada no horário
            if 6 <= current_hour < 12:
                greeting = "Bom dia"
                icon = APP_CONFIG['icons']['morning']
                emoji = "🌅"
            elif 12 <= current_hour < 18:
                greeting = "Boa tarde"
                icon = APP_CONFIG['icons']['afternoon']
                emoji = "☀️"
            elif 18 <= current_hour < 24:
                greeting = "Boa noite"
                icon = APP_CONFIG['icons']['night']
                emoji = "🌙"
            else:  # Entre 0h e 6h não mostra mensagem
                return
            
            # Mensagem personalizada
            message = (
                f"{greeting}, {username}! Bem-vindo ao sistema.\n"
                "Que seu trabalho seja produtivo e cheio de conquistas!"
            )
            
            self.show_system_notification(
                title=f"{greeting} {emoji}",
                message=message,
                icon_path=icon
            )
            logger.debug(f"Mensagem de {greeting.lower()} exibida para {username} às {current_hour}h")
            
        except Exception as e:
            logger.error(f"Erro ao mostrar mensagem de saudação: {e}")