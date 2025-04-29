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

# Font configurations (Tkinter style)
HEADER_FONT = ("Arial", 24, "bold")
SUBHEADER_FONT = ("Arial", 14)
LABEL_FONT = ("Arial", 12)
ENTRY_FONT = ("Arial", 12)
TABLE_HEADER_FONT = ("Arial", 12, "bold")
TABLE_FONT = ("Arial", 12)
BUTTON_FONT = ("Arial", 12)

# Tally Prime inspired color scheme
PRIMARY_COLOR = "#2c5282"      # Tally blue
SECONDARY_COLOR = "#2d3748"    # Dark blue gray
ACCENT_COLOR = "#4299e1"       # Light blue
BACKGROUND_COLOR = "#f7fafc"   # Light background
FRAME_COLOR = "#edf2f7"        # Frame background
BORDER_COLOR = "#e2e8f0"       # Border color
TEXT_COLOR = "#1a202c"         # Dark text
ERROR_COLOR = "#e53e3e"        # Red for errors
SUCCESS_COLOR = "#2c5282"      # Blue for success
WARNING_COLOR = "#d69e2e"      # Yellow for warnings
TABLE_TEXT_COLOR = "#1a202c"   # Dark text for table
TABLE_GRID_COLOR = "#e2e8f0"   # Grid color for table
 