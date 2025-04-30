from PySide6.QtGui import QFont

# Constants
CONFIG_FILE = "app_config.json"
AUTOSAVE_INTERVAL = 300000  # 5 minutes in milliseconds

# Item list for dropdown
ITEM_LIST = [
    "MAIZE", "SOYABEAN", "LOBHA", "HULLI", "KADLI", "BLACK MOONG", 
    "CHAMAKI MOONG", "RAGI", "WHEAT", "RICE", "BILAJOLA", "BIJAPUR", 
    "CHS-5", "FEEDS", "KUSUBI", "SASAVI", "SAVI", "CASTER SEEDS", 
    "TOOR RED", "TOOR WHITE", "HUNASIBIKA", "SF", "AWARI"
]

# Font configurations
HEADER_FONT = QFont("Segoe UI", 28, QFont.Bold)
SUBHEADER_FONT = QFont("Segoe UI", 16)
LABEL_FONT = QFont("Segoe UI", 13)
ENTRY_FONT = QFont("Segoe UI", 13)
TABLE_HEADER_FONT = QFont("Segoe UI", 13, QFont.Bold)
TABLE_FONT = QFont("Segoe UI", 13)
BUTTON_FONT = QFont("Segoe UI", 13)

# Color scheme
PRIMARY_COLOR = "#1976d2"      # Blue
SECONDARY_COLOR = "#2196f3"    # Lighter blue
ACCENT_COLOR = "#64b5f6"       # Even lighter blue
BACKGROUND_COLOR = "#ffffff"    # White
FRAME_COLOR = "#f5f5f5"        # Light gray
BORDER_COLOR = "#e0e0e0"       # Border gray
TEXT_COLOR = "#212121"         # Dark gray for text
ERROR_COLOR = "#f44336"        # Red