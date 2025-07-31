import customtkinter as ctk
from app.admin.ui.main_window import DatabaseConfigDialog
import logging
import os

logger = logging.getLogger(__name__)

# Ensure logs directory exists
if not os.path.exists('logs'):
    os.makedirs('logs')

# Logging configuration
logging.basicConfig(
    filename='logs/chronos_admin.log',
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)

class App:
    def __init__(self):
        # CustomTkinter settings
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")

        # Create main window
        self.root = ctk.CTk()
        self.root.withdraw()  # Hide the main window
        
        # Create and show database config dialog
        self.db_config = DatabaseConfigDialog(self.root)
        
        # Configure closing behavior
        self.db_config.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        logger.info("Administrative application started")
        
    def on_closing(self):
        """Handle application closing"""
        self.root.quit()
        
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = App()
    app.run()