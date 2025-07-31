import customtkinter as ctk
from PIL import Image
import os
import sys
import logging
from app.core.time.time_manager import TimeManager
from ...utils.window_manager import WindowManager

logger = logging.getLogger(__name__)

class BreakEndDialog(ctk.CTkToplevel):
    def __init__(self, parent=None):
        super().__init__(parent)
        
        # Configurações básicas da janela
        self.title("Fim do Intervalo")
        self.resizable(False, False)
        self.attributes('-topmost', True)  # Mantém a janela sempre no topo
        
        # Usar o WindowManager para posicionar a janela
        window_manager = WindowManager()
        window_manager.position_window(self, parent)
        
        # Carregar ícone
        try:
            if hasattr(sys, "_MEIPASS"):
                icons_dir = os.path.join(sys._MEIPASS, 'icons')
            else:
                icons_dir = os.path.join(os.path.abspath("."), 'icons')
                
            alert_path = os.path.join(icons_dir, 'clock.png')
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
        except Exception as e:
            logger.error(f"Erro ao carregar ícone: {e}")
            
        # Título
        self.title_label = ctk.CTkLabel(
            self,
            text="Fim do Intervalo!",
            font=("Roboto", 24, "bold"),
            text_color="#FF5722"
        )
        self.title_label.pack(pady=10)
        
        # Mensagem
        self.message_label = ctk.CTkLabel(
            self,
            text="O intervalo terminou!\nNão se esqueça de registrar seu retorno.",
            font=("Roboto", 14),
            justify="center"
        )
        self.message_label.pack(pady=20)
        
        # Botão de OK
        self.ok_button = ctk.CTkButton(
            self,
            text="OK",
            command=self.destroy,
            fg_color="#FF5722",
            hover_color="#CE461B",
            width=120,
            height=35
        )
        self.ok_button.pack(pady=20)
        
        # Trazer janela para frente
        self.lift()
        self.focus_force()
        
        # Configurar som de notificação (se disponível)
        try:
            self.bell()
        except:
            pass