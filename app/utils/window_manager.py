import win32api, json, os, ctypes, win32con, logging
from typing import Tuple
from ctypes import wintypes

logger = logging.getLogger(__name__)

class WindowManager:
    def __init__(self):
        self.config_file = "window_positions.json"
        self.positions = self._load_positions()
        self.last_monitor = self.positions.get('last_monitor', None)  # Carrega o último monitor usado
        logger.debug(f"[WINDOW] Carregou último monitor: {self.last_monitor}")
        self.window_monitors = {}

    def _load_positions(self) -> dict:
        """Carrega as posições salvas das janelas"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    return json.load(f)
        except Exception as e:
            logger.error(f"Erro ao carregar posições das janelas: {e}")
        return {}

    def _save_positions(self):
        """Salva as posições das janelas"""
        try:
            with open(self.config_file, 'w') as f:
                json.dump(self.positions, f)
        except Exception as e:
            logger.error(f"Erro ao salvar posições das janelas: {e}")

    def _save_monitor_position(self, monitor_rect):
        """Salva a posição do monitor"""
        try:
            self.last_monitor = monitor_rect
            self.positions['last_monitor'] = monitor_rect
            self._save_positions()
            logger.debug(f"[WINDOW] Salvou posição do monitor: {monitor_rect}")
        except Exception as e:
            logger.error(f"[WINDOW] Erro ao salvar posição do monitor: {e}")

    def get_cursor_pos(self) -> Tuple[int, int]:
        """Obtém a posição atual do cursor usando ctypes"""
        try:
            point = wintypes.POINT()
            ctypes.windll.user32.GetCursorPos(ctypes.byref(point))
            return (point.x, point.y)
        except Exception as e:
            logger.error(f"Erro ao obter posição do cursor: {e}")
            return (0, 0)

    def get_current_monitor(self) -> Tuple[int, int, int, int]:
        """Retorna as dimensões do monitor onde o mouse está"""
        try:
            # Pega a posição atual do mouse usando ctypes
            cursor_x, cursor_y = self.get_cursor_pos()
            logger.debug(f"Posição do cursor: x={cursor_x}, y={cursor_y}")
            
            # Pega o monitor onde o mouse está
            monitor = win32api.MonitorFromPoint((cursor_x, cursor_y), win32api.MONITOR_DEFAULTTONEAREST)
            monitor_info = win32api.GetMonitorInfo(monitor)
            monitor_rect = monitor_info['Monitor']
            
            logger.debug(f"Monitor encontrado: {monitor_rect}")
            return monitor_rect
            
        except Exception as e:
            logger.error(f"Erro ao obter monitor atual: {e}")
            return None

    def is_window_minimized(self, window) -> bool:
        """Verifica se uma janela está minimizada"""
        try:
            hwnd = window.winfo_id()
            # Usar GetWindowLong ao invés de GetWindowPlacement
            style = win32api.GetWindowLong(hwnd, win32con.GWL_STYLE)
            return (style & win32con.WS_MINIMIZE) != 0
        except Exception as e:
            logger.error(f"[WINDOW] Erro ao verificar estado da janela: {e}")
            return False

    def get_monitor_from_window(self, window) -> Tuple[int, int, int, int]:
        """Obtém o monitor onde a janela está"""
        try:
            # Se a janela estiver minimizada, usar o monitor do cursor (como no system tray)
            if self.is_window_minimized(window):
                logger.debug(f"[WINDOW] Janela {window} está minimizada, usando monitor do cursor")
                return self.get_current_monitor()

            # Se não estiver minimizada, obter o monitor atual
            hwnd = window.winfo_id()
            monitor = win32api.MonitorFromWindow(hwnd, win32con.MONITOR_DEFAULTTONEAREST)
            monitor_info = win32api.GetMonitorInfo(monitor)
            monitor_rect = monitor_info['Monitor']
            
            # Atualizar o último monitor conhecido
            self.last_monitor = monitor_rect
            self._save_monitor_position(monitor_rect)
            logger.debug(f"[WINDOW] Monitor atual da janela: {monitor_rect}")
            
            return monitor_rect
            
        except Exception as e:
            logger.error(f"[WINDOW] Erro ao obter monitor da janela: {e}")
            return self.get_current_monitor()  # Fallback para o monitor do cursor

    def position_window(self, window, parent=None):
        """Posiciona a janela no monitor correto"""
        try:
            # Define as dimensões da janela
            if window.__class__.__name__ == 'LoginWindow':
                width, height = 900, 600
            elif window.__class__.__name__ in ['MainWindow', 'UserWindow']:
                width, height = 1300, 700
            elif window.__class__.__name__ == 'ChangePasswordDialog':
                width, height = 400, 300
            elif window.__class__.__name__ in ['ReasonExceededDialog', 'RegisterDialog', 'ActivitiesPrinterDialog', 'ActivityForm']:
                width, height = 500, 600
            elif window.__class__.__name__ in ['BreakStartDialog', 'BreakEndDialog', 'CompanyEndDialog', 'CompanyEndWarningDialog', 'TimeExceededDialog']:
                width, height = 400, 350
            elif window.__class__.__name__ == 'DashboardDaily':
                width, height = 1200, 700
            elif window.__class__.__name__ == 'ExcelViewer':
                width, height = 1200, 600
            elif window.__class__.__name__ == 'ExcelSelector':
                width = int(window.winfo_screenwidth() * 0.8)
                height = int(window.winfo_screenheight() * 0.8)
            else:
                width = window.winfo_reqwidth()
                height = window.winfo_reqheight()

            # Determina o monitor correto
            monitor_rect = None
            
            # Se tem janela pai, usa ela como referência
            if parent and parent.winfo_exists():
                monitor_rect = self.get_monitor_from_window(parent)
                if monitor_rect:
                    logger.debug(f"[WINDOW] Usando monitor da janela pai: {monitor_rect}")
                else:
                    logger.warning("[WINDOW] Não foi possível determinar monitor da janela pai")
                    monitor_rect = self.get_current_monitor()
            
            # Se não tem pai ou falhou em obter monitor do pai
            if not monitor_rect:
                if window.__class__.__name__ == 'LoginWindow':
                    if self.last_monitor:
                        monitor_rect = self.last_monitor
                        logger.debug(f"[WINDOW] Login window usando último monitor salvo: {monitor_rect}")
                    else:
                        monitor = win32api.MonitorFromWindow(window.winfo_id(), win32con.MONITOR_DEFAULTTOPRIMARY)
                        monitor_rect = win32api.GetMonitorInfo(monitor)['Monitor']
                        logger.debug(f"[WINDOW] Login window usando monitor primário: {monitor_rect}")
                else:
                    monitor_rect = self.last_monitor or self.get_current_monitor()
                    logger.debug(f"[WINDOW] Usando monitor alternativo: {monitor_rect}")

            # Fallback absoluto: monitor principal
            if not monitor_rect:
                logger.warning(f"[WINDOW] Nenhum monitor encontrado para {window.__class__.__name__}, usando monitor principal")
                screen_width = win32api.GetSystemMetrics(0)
                screen_height = win32api.GetSystemMetrics(1)
                monitor_rect = (0, 0, screen_width, screen_height)

            # Salva o monitor usado
            self._save_monitor_position(monitor_rect)
            
            # Extrai as coordenadas do monitor
            monitor_left, monitor_top, monitor_right, monitor_bottom = monitor_rect
            
            # Calcula o centro do monitor
            monitor_width = monitor_right - monitor_left
            monitor_height = monitor_bottom - monitor_top
            
            # Calcula a posição centralizada
            x = monitor_left + (monitor_width - width) // 2
            y = monitor_top + (monitor_height - height) // 2

            # Restaurar geometry salva se for MainWindow
            if window.__class__.__name__ == 'MainWindow' and 'main_window_geometry' in self.positions:
                try:
                    window.geometry(self.positions['main_window_geometry'])
                except Exception as e:
                    logger.warning(f"[WINDOW] Falha ao restaurar geometry salva: {e}, centralizando...")
                    window.geometry(f"{width}x{height}+{x}+{y}")
            else:
                window.geometry(f"{width}x{height}+{x}+{y}")
            window.update_idletasks()  # Força atualização imediata
            
            # Registra a posição para debug
            logger.debug(f"[WINDOW] Posicionando {window.__class__.__name__} em monitor {monitor_rect}")
            logger.debug(f"[WINDOW] Geometria calculada: {width}x{height}+{x}+{y}")
            
        except Exception as e:
            logger.error(f"[WINDOW] Erro ao posicionar janela: {e}")
            # Fallback para centro do monitor principal
            try:
                screen_width = win32api.GetSystemMetrics(0)
                screen_height = win32api.GetSystemMetrics(1)
                x = (screen_width - width) // 2
                y = (screen_height - height) // 2
                window.geometry(f"{width}x{height}+{x}+{y}")
            except:
                window.geometry(f"{width}x{height}+0+0")

    def is_window_maximized(self, window):
        # Implemente a lógica para verificar se uma janela está maximizada
        # Isso pode ser feito usando win32api.GetWindowPlacement ou outras APIs
        # Por enquanto, vamos assumir que todas as janelas estão maximizadas
        return False 