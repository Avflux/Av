import customtkinter as ctk
from PIL import Image
import os
import sys
import logging
from ...utils.window_manager import WindowManager

logger = logging.getLogger(__name__)

class TimeExceededDialog(ctk.CTkToplevel):
    _instance = None  # Variável de classe para controlar a instância única
    
    def __new__(cls, *args, **kwargs):
        logger.debug(f"[TIME_EXCEEDED] Tentando criar nova instância. Instance existente: {cls._instance}")
        if cls._instance is None or not cls._instance.winfo_exists():
            logger.info("[TIME_EXCEEDED] Criando nova instância do diálogo")
            cls._instance = super().__new__(cls)
            return cls._instance
        else:
            logger.info("[TIME_EXCEEDED] Instância já existe, trazendo para frente")
            try:
                cls._instance.lift()
                cls._instance.focus_force()
                cls._instance.bell()
                logger.debug("[TIME_EXCEEDED] Janela existente atualizada com sucesso")
            except Exception as e:
                logger.error(f"[TIME_EXCEEDED] Erro ao atualizar janela existente: {e}")
            return None
    
    def __init__(self, parent, activity_name=None):
        if not hasattr(self, '_initialized'):
            logger.info(f"[TIME_EXCEEDED] Inicializando diálogo. Parent: {parent}, Activity: {activity_name}")
            
            # Verificar se o parent é válido
            valid_parent = None
            try:
                if parent and parent.winfo_exists():
                    try:
                        # Testar se conseguimos acessar a janela pai
                        _ = parent.winfo_id()
                        valid_parent = parent
                        logger.debug(f"[TIME_EXCEEDED] Info da janela pai - Existe: {parent.winfo_exists()}, "
                                   f"ID: {parent.winfo_id()}, Geometria: {parent.winfo_geometry()}")
                    except Exception as e:
                        logger.warning(f"[TIME_EXCEEDED] Janela pai inválida: {e}")
                        valid_parent = None
                else:
                    logger.warning("[TIME_EXCEEDED] Janela pai não existe ou é None")
                    valid_parent = None
                
                # Inicializar com o parent válido ou None
                super().__init__(valid_parent)
                self._initialized = True
                
                # Configurações básicas da janela
                self.title("Tempo Excedido")
                self.resizable(False, False)
                
                # Definir geometria inicial correta
                self.geometry("400x350")  # Definir tamanho antes de qualquer outra coisa
                
                self.attributes('-topmost', True)
                
                logger.debug("[TIME_EXCEEDED] Configurações básicas aplicadas")
                
                # Remover botões de minimizar e maximizar
                self.overrideredirect(False)
                
                # Tornar a janela modal se tivermos um parent válido
                if valid_parent:
                    self.transient(valid_parent)
                    self.grab_set()
                    logger.debug("[TIME_EXCEEDED] Configurações modais aplicadas com sucesso")
                else:
                    logger.warning("[TIME_EXCEEDED] Não foi possível aplicar configurações modais - parent inválido")
                
                # Usar o WindowManager para posicionar a janela
                try:
                    window_manager = WindowManager()
                    window_manager.position_window(self, valid_parent)
                    logger.debug("[TIME_EXCEEDED] Janela posicionada pelo WindowManager")
                except Exception as e:
                    logger.error(f"[TIME_EXCEEDED] Erro ao posicionar janela: {e}")
                
                # Carregar ícone de alerta
                try:
                    if hasattr(sys, "_MEIPASS"):
                        icons_dir = os.path.join(sys._MEIPASS, 'icons')
                    else:
                        icons_dir = os.path.join(os.path.abspath("."), 'icons')
                        
                    alert_path = os.path.join(icons_dir, 'alert.png')
                    if os.path.exists(alert_path):
                        alert_image = Image.open(alert_path)
                        self.alert_icon = ctk.CTkImage(
                            light_image=alert_image,
                            dark_image=alert_image,
                            size=(64, 64)
                        )
                        
                        # Adicionar ícone
                        self.icon_label = ctk.CTkLabel(
                            self,
                            image=self.alert_icon,
                            text=""
                        )
                        self.icon_label.pack(pady=20)
                        logger.debug("[TIME_EXCEEDED] Ícone carregado com sucesso")
                except Exception as e:
                    logger.error(f"[TIME_EXCEEDED] Erro ao carregar ícone: {e}")
            
                # Título
                self.title_label = ctk.CTkLabel(
                    self,
                    text="Tempo Excedido!",
                    font=("Roboto", 24, "bold"),
                    text_color="#FF5722"
                )
                self.title_label.pack(pady=10)
                
                # Mensagem
                message_text = "O tempo estimado para esta atividade foi excedido!"
                if activity_name:
                    message_text += f"\n\nAtividade: {activity_name}"
                
                self.message_label = ctk.CTkLabel(
                    self,
                    text=message_text,
                    font=("Roboto", 14),
                    justify="center"
                )
                self.message_label.pack(pady=20)
                
                # Botão de OK
                self.ok_button = ctk.CTkButton(
                    self,
                    text="OK",
                    command=self.on_close,
                    fg_color="#FF5722",
                    hover_color="#CE461B",
                    width=120,
                    height=35
                )
                self.ok_button.pack(pady=20)
                
                # Trazer janela para frente e forçar foco
                self.lift()
                self.focus_force()
                
                # Log do estado final da janela
                logger.debug(f"[TIME_EXCEEDED] Estado final da janela - "
                           f"Geometria: {self.winfo_geometry()}, "
                           f"Visível: {self.winfo_viewable()}, "
                           f"ID: {self.winfo_id()}")
                
                # Configurar som de notificação
                try:
                    self.bell()
                except Exception as e:
                    logger.error(f"[TIME_EXCEEDED] Erro ao tocar som: {e}")
                
                # Garantir que a janela permaneça visível
                self.after(100, self._ensure_visibility)
                
                # Protocolo para fechar a janela
                self.protocol("WM_DELETE_WINDOW", self.on_close)
                
            except Exception as e:
                logger.error(f"[TIME_EXCEEDED] Erro durante inicialização: {e}")

    def _ensure_visibility(self):
        try:
            logger.debug("[TIME_EXCEEDED] Verificando visibilidade da janela")
            
            # Obter informações do monitor atual
            try:
                window_manager = WindowManager()
                current_monitor = window_manager.get_monitor_from_window(self)
                if current_monitor:
                    monitor_left, monitor_top, monitor_right, monitor_bottom = current_monitor
                    screen_width = monitor_right - monitor_left
                    screen_height = monitor_bottom - monitor_top
                    base_x = monitor_left
                    base_y = monitor_top
                    logger.debug(f"[TIME_EXCEEDED] Usando monitor: {current_monitor}")
                else:
                    # Fallback para monitor primário
                    screen_width = self.winfo_screenwidth()
                    screen_height = self.winfo_screenheight()
                    base_x = 0
                    base_y = 0
                    logger.debug("[TIME_EXCEEDED] Usando monitor primário como fallback")
            except Exception as e:
                logger.error(f"[TIME_EXCEEDED] Erro ao obter monitor: {e}")
                screen_width = self.winfo_screenwidth()
                screen_height = self.winfo_screenheight()
                base_x = 0
                base_y = 0
            
            width = self.winfo_width()
            height = self.winfo_height()
            
            # Se as dimensões estiverem erradas, corrigir
            if width != 400 or height != 350:
                width = 400
                height = 350
                logger.warning(f"[TIME_EXCEEDED] Dimensões incorretas detectadas: {self.winfo_geometry()}, corrigindo para {width}x{height}")
            
            x = self.winfo_x()
            y = self.winfo_y()
            
            original_geometry = self.winfo_geometry()
            
            # Ajustar coordenadas relativas ao monitor atual
            if x + width > base_x + screen_width:
                x = base_x + screen_width - width
            if x < base_x:
                x = base_x
                
            if y + height > base_y + screen_height:
                y = base_y + screen_height - height
            if y < base_y:
                y = base_y
            
            new_geometry = f'{width}x{height}+{x}+{y}'
            
            # Só aplicar nova geometria se realmente mudou
            if new_geometry != original_geometry:
                self.geometry(new_geometry)
                logger.debug(f"[TIME_EXCEEDED] Geometria ajustada - Original: {original_geometry}, Nova: {new_geometry}")
            
            # Verificar se a janela está realmente visível
            if not self.winfo_viewable():
                logger.warning("[TIME_EXCEEDED] Janela não está visível após ajuste")
                self.lift()
                self.focus_force()
                
            # Verificar se a janela está no monitor correto
            if current_monitor:
                actual_monitor = window_manager.get_monitor_from_window(self)
                if actual_monitor != current_monitor:
                    logger.warning(f"[TIME_EXCEEDED] Janela está no monitor errado. Esperado: {current_monitor}, Atual: {actual_monitor}")
                    # Reposicionar usando WindowManager
                    window_manager.position_window(self, self.master)
            
        except Exception as e:
            logger.error(f"[TIME_EXCEEDED] Erro ao ajustar visibilidade: {e}")
    
    def on_close(self):
        """Método chamado quando a janela é fechada"""
        try:
            logger.info("[TIME_EXCEEDED] Fechando janela")
            TimeExceededDialog._instance = None
            self.grab_release()
            self.destroy()
            logger.debug("[TIME_EXCEEDED] Janela fechada com sucesso")
        except Exception as e:
            logger.error(f"[TIME_EXCEEDED] Erro ao fechar janela: {e}")
