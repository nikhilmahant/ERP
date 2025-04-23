import customtkinter as ctk
from tkinter import messagebox, TclError
import logging
from logging.handlers import RotatingFileHandler
from datetime import datetime
import os
import json
import re
import shutil
import platform
import tempfile
import win32api
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from excel_handler import save_to_excel
from printer_handler import print_invoice as direct_print_invoice, generate_pdf
from print_utils import print_with_dialog, load_print_setting, open_settings_window
from tkinter import ttk
# Define translation function globally
_ = lambda x: x  # Default to identity function (English)

# Localization setup (English and Kannada)
lang = 'en'  # Default to English; can be changed to 'kn' for Kannada
if lang == 'kn':
    try:
        import gettext
        kn = gettext.translation('invoice_app', localedir='locale', languages=['kn'])
        kn.install()
        _ = kn.gettext
    except Exception as e:
        print(f"Translation error: {e}")
        # Keep default English translation

# Make _ available globally
import builtins
builtins._ = _  # This makes _ available to all modules without importing

# Configure logging with rotation
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger()
handler = RotatingFileHandler('invoice_app.log', maxBytes=1_000_000, backupCount=5)
handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
logger.addHandler(handler)

# Font configurations
HEADER_FONT = ("Segoe UI", 24, "bold")
SUBHEADER_FONT = ("Segoe UI", 14)
LABEL_FONT = ("Segoe UI", 12)
ENTRY_FONT = ("Segoe UI", 12)
TABLE_HEADER_FONT = ("Segoe UI", 12, "bold")
TABLE_FONT = ("Segoe UI", 12)
BUTTON_FONT = ("Segoe UI", 12)

# Item list for dropdown
ITEM_LIST = [
    "MAIZE",
    "SOYABEAN",
    "LOBHA",
    "HULLI",
    "KADLI",
    "BLACK MOONG",
    "CHAMAKI MOONG",
    "RAGI",
    "WHEAT",
    "RICE",
    "BILAJOLA",
    "BIJAPUR",
    "CHS-5",
    "FEEDS",
    "KUSUBI",
    "SASAVI",
    "SAVI",
    "CASTER SEEDS",
    "TOOR RED",
    "TOOR WHITE",
    "HUNASIBIKA",
    "SF",
    "AWARI"
]

# Mode configurations
MODES = {
    "Patti": {
        "headers": ["Item", "Packet", "Quantity", "Rate", "Hamali", "Amount"],
        "fields": 5,
        "calc": lambda w: validate_float(w[2].get()) * validate_float(w[3].get()) + 
                        validate_float(w[1].get()) * validate_float(w[4].get())
    },
    "Kata": {
        "headers": ["Item", "Net Wt", "Less%", "Rate", "Hamali Rate", "Amount"],
        "fields": 5,
        "calc": lambda w: (validate_float(w[1].get()) * (1 - min(validate_float(w[2].get()), 100) / 100) * 
                          validate_float(w[3].get()) + 
                          int(validate_float(w[1].get()) / 60) * validate_float(w[4].get()))
    },
    "Barthe": {
        "headers": ["Item", "Packet", "Weight", "+/-", "Rate", "Hamali", "Amount"],
        "fields": 6,
        "calc": lambda w: ((validate_float(w[1].get()) * validate_float(w[2].get()) + 
                           validate_float(w[3].get())) * validate_float(w[4].get()) + 
                          validate_float(w[1].get()) * validate_float(w[5].get()))
    }
}

def validate_float(value):
    """Validate if a string can be converted to a non-negative float.

    Args:
        value (str): Input string to validate.

    Returns:
        float: Validated float value, or 0 if invalid.
    """
    try:
        result = float(value) if value.strip() else 0
        if result < 0:
            raise ValueError("Negative numbers are not allowed")
        return result
    except ValueError:
        return 0

def load_config():
    """Load configuration from config.json.

    Returns:
        dict: Configuration settings.
    """
    default_config = {
        "printer": "default",
        "paper_size": "A4",
        "orientation": "portrait",
        "max_rows": 100,
        "print_mode": "dialog"
    }
    try:
        with open('config.json', 'r', encoding='utf-8') as f:
            content = f.read().strip()
            if not content:  # Handle empty file
                with open('config.json', 'w', encoding='utf-8') as fw:
                    json.dump(default_config, fw, indent=4)
                return default_config
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        with open('config.json', 'w', encoding='utf-8') as f:
            json.dump(default_config, f, indent=4)
        return default_config


def load_print_setting():
    """Load print setting from config.json"""
    default = {"print_mode": "dialog"}  # Options: dialog, direct, pdf
    try:
        if os.path.exists("config.json"):
            with open("config.json", "r", encoding="utf-8") as f:
                data = json.load(f)
                return data.get("print_mode", "dialog")
        return default["print_mode"]
    except Exception:
        return default["print_mode"]

def load_recent_customers():
    """Load recent customer names from recent_customers.json.

    Returns:
        list: List of recent customer names.
    """
    try:
        with open('recent_customers.json', 'r') as f:
            return json.load(f)
    except FileNotFoundError:
        return []

def save_recent_customer(name):
    """Save a customer name to recent_customers.json.

    Args:
        name (str): Customer name to save.
    """
    if not name or name == "Unknown Customer":
        return
    customers = load_recent_customers()
    if name not in customers:
        customers.append(name)
        customers = customers[-10:]  # Keep last 10 customers
        with open('recent_customers.json', 'w') as f:
            json.dump(customers, f, indent=4)

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

class InvoiceApp(ctk.CTk):
    """Main application class for the Invoice Management System."""

    def __init__(self):
        """Initialize the application."""
        super().__init__()
        self.title(_("G.V. Mahant Brothers - Invoice"))
        self.geometry("1200x800")
        self.resizable(True, True)
        self.configure(padx=20, pady=20)

        self.config = load_config()
        self.current_mode = ctk.StringVar(value="Patti")
        self.rows = []
        self.undo_stack = []
        self.redo_stack = []
        self.max_rows = self.config.get("max_rows", 100)

        self.build_ui()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def build_ui(self):
        """Build the main user interface."""
        # Header
        ctk.CTkLabel(self, text="", height=1).pack()
        ctk.CTkLabel(self, text="|ಶ್ರೀ|", font=HEADER_FONT, text_color="#1976d2").pack()
        ctk.CTkLabel(self, text="G.V. Mahant Brothers", font=HEADER_FONT, text_color="#1976d2").pack()
        self.date_label = ctk.CTkLabel(self, text=datetime.now().strftime("%A, %d %B %Y %I:%M %p"), font=SUBHEADER_FONT)
        self.date_label.pack()

        # Menu bar
        menubar = ctk.CTkFrame(self, fg_color="transparent")
        ctk.CTkButton(menubar, text=_("Help"), command=self.show_help, font=BUTTON_FONT).pack(side="right", padx=10)
        ctk.CTkButton(menubar, text=_("Settings"), command=self.open_settings, font=BUTTON_FONT).pack(side="right", padx=10)
        menubar.pack(fill="x")

        # Mode navigation
        nav_frame = ctk.CTkFrame(self, fg_color="transparent")
        for i, mode in enumerate(MODES.keys()):
            rb = ctk.CTkRadioButton(nav_frame, text=_(mode), variable=self.current_mode, value=mode, 
                                   command=self.switch_mode, font=LABEL_FONT)
            rb.grid(row=0, column=i, padx=20, pady=10)
        nav_frame.pack(pady=(30, 20))

        # Customer name entry
        customer_frame = ctk.CTkFrame(self, fg_color="transparent")
        ctk.CTkLabel(customer_frame, text=_("Customer Name:"), font=LABEL_FONT).pack(side="left", padx=(0, 10))
        self.customer_entry = ctk.CTkEntry(customer_frame, width=400, font=ENTRY_FONT, height=35)
        self.customer_entry.pack(side="left")
        customer_frame.pack(pady=(20, 30))

        # Table container
        self.table_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.table_frame.pack(fill="both", expand=True, pady=(0, 20))
        self.create_table_headers()

        # Create a frame for buttons and total below the table rows
        self.buttons_frame = ctk.CTkFrame(self.table_frame, fg_color="transparent")

        self.add_row()

        # Update date every minute
        self.after(60000, self.update_date)

    def update_date(self):
        """Update the date label every minute."""
        self.date_label.configure(text=datetime.now().strftime("%A, %d %B %Y %I:%M %p"))
        self.after(60000, self.update_date)

    def create_table_headers(self):
        """Create table headers based on the current mode."""
        for widget in self.table_frame.winfo_children():
            widget.destroy()
        self.rows.clear()
        
        # Create a frame for buttons that will appear below the rows
        self.buttons_frame = ctk.CTkFrame(self.table_frame, fg_color="transparent")

        headers = MODES[self.current_mode.get()]["headers"]
        for i, h in enumerate(headers):
            ctk.CTkLabel(
                self.table_frame,
                text=_(h),
                font=TABLE_HEADER_FONT,
                text_color="white",
                fg_color="#1976d2",
                corner_radius=5,
                width=100,
                height=35,
                anchor="center"
            ).grid(row=0, column=i, sticky="nsew", padx=2, pady=2)
            self.table_frame.grid_columnconfigure(i, weight=1)
        # Delete column
        ctk.CTkLabel(
            self.table_frame,
            text="",
            font=TABLE_HEADER_FONT,
            text_color="white",
            fg_color="#1976d2",
            corner_radius=5,
            width=40,
            height=35
        ).grid(row=0, column=len(headers), sticky="nsew", padx=2, pady=2)
        self.table_frame.grid_columnconfigure(len(headers), weight=0)

    def add_buttons_below_rows(self):
        """Add action buttons and total amount below the table rows."""
        # Clear any existing buttons
        for widget in self.buttons_frame.winfo_children():
            widget.destroy()
            
        # Calculate the row index for buttons (after the last data row)
        buttons_row = len(self.rows) + 1 if self.rows else 1
        
        # Position the buttons frame in the grid
        self.buttons_frame.grid(row=buttons_row, column=0, columnspan=len(MODES[self.current_mode.get()]["headers"]) + 2, 
                               sticky="ew", padx=2, pady=10)
        
        # Create a container for buttons on the left
        buttons_container = ctk.CTkFrame(self.buttons_frame, fg_color="transparent")
        buttons_container.pack(side="left", fill="x", expand=True)
        
        # Add buttons to the container
        button_params = {"font": BUTTON_FONT, "width": 120, "height": 35}
        ctk.CTkButton(buttons_container, text=_("+ Add Row"), command=self.add_row, **button_params).pack(side="left", padx=10)
        ctk.CTkButton(buttons_container, text=_("Clear"), command=self.confirm_clear_rows, **button_params).pack(side="left", padx=10)
        ctk.CTkButton(buttons_container, text=_("Save"), command=self.save_to_excel, **button_params).pack(side="left", padx=10)
        ctk.CTkButton(buttons_container, text=_("Print"), command=self.print_invoice, **button_params).pack(side="left", padx=10)
        
        # Create a right container for total and Kata amount (if in Kata mode)
        right_container = ctk.CTkFrame(self.buttons_frame, fg_color="transparent")
        right_container.pack(side="right", padx=20)
        
        # Add Kata amount field if in Kata mode
        self.kata_amount_var = ctk.StringVar(value="0")
        if self.current_mode.get() == "Kata":
            kata_frame = ctk.CTkFrame(right_container, fg_color="transparent")
            kata_frame.pack(side="top", pady=(0, 10), anchor="e")
            
            ctk.CTkLabel(kata_frame, text=_("Kata:"), font=LABEL_FONT).pack(side="left", padx=(0, 5))
            kata_entry = ctk.CTkEntry(kata_frame, width=120, font=ENTRY_FONT, textvariable=self.kata_amount_var)
            kata_entry.pack(side="left")
            kata_entry.bind("<KeyRelease>", lambda e: self.update_amounts())
        
        # Add total amount
        self.total_label = ctk.CTkLabel(right_container, text=_("Amount: ₹0.00"), font=("Segoe UI", 16, "bold"))
        self.total_label.pack(side="bottom", anchor="e")
    
    def switch_mode(self):
        """Switch between Patti, Kata, and Barthe modes."""
        self.save_state_for_undo()
        self.create_table_headers()
        self.add_row()
        self.update_amounts()

    def add_row(self):
        """Add a new row to the table."""
        if len(self.rows) >= self.max_rows:
            messagebox.showwarning(_("Warning"), _("Maximum row limit reached (%d).") % self.max_rows)
            return

        # If buttons_frame exists in the grid, remove it temporarily
        if hasattr(self, 'buttons_frame') and self.buttons_frame.winfo_exists():
            self.buttons_frame.grid_forget()
            
        self.save_state_for_undo()
        mode = self.current_mode.get()
        num_fields = MODES[mode]["fields"]
        entries = []
        row_idx = len(self.rows) + 1
        bg_even = "#e3f2fd"
        bg_odd = "#ffffff"
        row_bg = bg_even if row_idx % 2 == 0 else bg_odd

        # First column is always Item - use a dropdown instead of an entry
        item_dropdown = ttk.Combobox(
            self.table_frame,
            values=ITEM_LIST,
            font=TABLE_FONT,
            height=10,
            state="readonly"
        )
        item_dropdown.grid(row=row_idx, column=0, padx=2, pady=2, sticky="nsew")
        item_dropdown.bind("<<ComboboxSelected>>", lambda e: self.update_amounts())
        entries.append(item_dropdown)
        
        # Add remaining entry fields
        for i in range(1, num_fields):
            entry = ctk.CTkEntry(
                self.table_frame,
                font=TABLE_FONT,
                justify="center",
                height=35,
                fg_color=row_bg
            )
            entry.grid(row=row_idx, column=i, padx=2, pady=2, sticky="nsew")
            entry.bind("<KeyRelease>", lambda e, ent=entry: self.validate_and_update(ent))
            entries.append(entry)

        amount_label = ctk.CTkLabel(
            self.table_frame,
            text="₹0.00",
            font=TABLE_FONT,
            anchor="e",
            height=35,
            fg_color=row_bg
        )
        amount_label.grid(row=row_idx, column=num_fields, padx=2, pady=2, sticky="nsew")
        entries.append(amount_label)

        delete_btn = ctk.CTkButton(
            self.table_frame,
            text="X",
            width=40,
            height=35,
            fg_color="red",
            hover_color="darkred",
            command=lambda: self.delete_row(row_idx)
        )
        delete_btn.grid(row=row_idx, column=num_fields + 1, padx=2, pady=2)
        entries.append(delete_btn)

        self.rows.append((row_idx, entries))
        
        # Add buttons below the last row
        self.add_buttons_below_rows()
        
        self.update_amounts()

    def delete_row(self, row_idx):
        """Delete a specific row from the table.

        Args:
            row_idx (int): Row index to delete.
        """
        self.save_state_for_undo()
        for idx, (r_idx, widgets) in enumerate(self.rows):
            if r_idx == row_idx:
                for w in widgets:
                    w.destroy()
                self.rows.pop(idx)
                break
        # Reindex rows
        for i, (idx, widgets) in enumerate(self.rows, 1):
            for w in widgets:
                try:
                    w.grid(row=i, column=w.grid_info()['column'])
                except TclError:
                    continue
        
        # Reposition buttons below the last row
        self.add_buttons_below_rows()
        
        self.update_amounts()

    def confirm_clear_rows(self):
        """Confirm before clearing all rows."""
        if messagebox.askyesno(_("Confirm"), _("Are you sure you want to clear all rows?")):
            self.save_state_for_undo()
            self.clear_rows()

    def clear_rows(self):
        """Clear all rows except headers."""
        for widget in self.table_frame.winfo_children():
            info = widget.grid_info()
            if info.get('row', 0) != 0 and widget != self.buttons_frame:
                widget.destroy()
        self.rows.clear()
        self.add_row()
        self.update_amounts()

    def validate_and_update(self, entry):
        """Validate entry input and update amounts.

        Args:
            entry (CTkEntry): Entry widget to validate.
        """
        value = entry.get()
        try:
            if value.strip():
                float_val = float(value)
                if float_val < 0:
                    entry.configure(fg_color="pink")
                    return
                if self.current_mode.get() == "Kata" and entry.grid_info()['column'] == 2:
                    if float_val > 100:
                        entry.configure(fg_color="pink")
                        return
            entry.configure(fg_color="white")
            self.update_amounts()
        except ValueError:
            entry.configure(fg_color="pink")

    def update_amounts(self):
        """Update amounts for all rows and total."""
        logger.debug("Updating amounts for all rows")
        total = 0
        mode = self.current_mode.get()
        calc_func = MODES[mode]["calc"]

        for _, widgets in self.rows:
            try:
                amount = calc_func(widgets)
                widgets[-2].configure(text=f"₹{amount:.2f}")  # Amount label
                total += amount
            except Exception as ex:
                logger.error(f"Calculation error: {ex}")
                widgets[-2].configure(text="₹0.00")
        
        # Add Kata amount if in Kata mode
        if mode == "Kata" and hasattr(self, 'kata_amount_var'):
            try:
                kata_amount = validate_float(self.kata_amount_var.get())
                total += kata_amount
            except Exception as ex:
                logger.error(f"Kata amount calculation error: {ex}")

        # Update the total label if it exists (it might not exist during initialization)
        if hasattr(self, 'total_label') and self.total_label.winfo_exists():
            self.total_label.configure(text=f"Amount: ₹{total:.2f}")

    def save_to_excel(self):
        """Save invoice data to Excel."""
        customer = self.validate_customer_name(self.customer_entry.get().strip())
        save_recent_customer(customer)
        try:
            # Get additional data for saving
            additional_data = {}
            if self.current_mode.get() == "Kata" and hasattr(self, 'kata_amount_var'):
                additional_data["kata_amount"] = validate_float(self.kata_amount_var.get())
                
            save_to_excel(self.current_mode.get(), customer, self.rows, MODES, additional_data)
            messagebox.showinfo(_("Saved"), _("Invoice data saved successfully."))
        except Exception as e:
            logger.error(f"Error saving to Excel: {str(e)}")
            messagebox.showerror(_("Save Error"), str(e))

    def validate_customer_name(self, name):
        """Validate and sanitize customer name.

        Args:
            name (str): Customer name to validate.

        Returns:
            str: Sanitized customer name.
        """
        if not name:
            return "Unknown Customer"
        name = re.sub(r'[<>:"/\\|?*]', '', name)[:100]  # Remove invalid chars, limit to 100
        return name or "Unknown Customer"

    def show_print_preview(self):
        """Show a print preview window."""
        try:
            preview = ctk.CTkToplevel(self)
            preview.title(_("Print Preview"))
            preview.geometry("400x600")
            preview.grab_set()
            x = self.winfo_x() + (self.winfo_width() - 400) // 2
            y = self.winfo_y() + (self.winfo_height() - 600) // 2
            preview.geometry(f"400x600+{x}+{y}")

            preview_frame = ctk.CTkScrollableFrame(preview)
            preview_frame.pack(fill="both", expand=True, padx=10, pady=10)

            # Use Courier New for print preview
            preview_text = ctk.CTkTextbox(preview_frame, font=("Courier New", 10))
            preview_text.pack(fill="both", expand=True)

            lines = self.generate_print_content()
            preview_text.insert("1.0", "\n".join(lines))
            preview_text.configure(state="disabled")

            button_frame = ctk.CTkFrame(preview, fg_color="transparent")
            button_frame.pack(side="bottom", fill="x", padx=10, pady=10)

            ctk.CTkButton(
                button_frame,
                text=_("Print"),
                command=lambda: [preview.destroy(), self.print_invoice()],
                width=120
            ).pack(side="left", padx=5)

            ctk.CTkButton(
                button_frame,
                text=_("Close"),
                command=preview.destroy,
                width=120
            ).pack(side="right", padx=5)

        except Exception as e:
            logger.error(f"Preview error: {str(e)}")
            messagebox.showerror(_("Preview Error"), str(e))

    def generate_print_content(self):
        """Generate content for printing or preview (80mm / 42 chars)."""
        width = 42
        lines = []

        lines += [
            "|| ಶ್ರೀ ||".center(width),
            "G.V. Mahant Brothers".center(width),
            datetime.now().strftime("%d-%m-%Y %I:%M %p").center(width),
            "-" * width,
            f"Customer: {self.validate_customer_name(self.customer_entry.get().strip())}",
            f"Mode: {self.current_mode.get()}",
            "-" * width,
        ]

        headers = MODES[self.current_mode.get()]["headers"]
        header_line = "".join(f"{_(h)[:5]:<7}" for h in headers[:-1]) + "Amt"
        lines.append(header_line[:width])
        lines.append("-" * width)

        for _, widgets in self.rows:
            if any(w.get().strip() for w in widgets[:-2]):
                row_text = []
                row_text.append(f"{widgets[0].get()[:6]:<6}")  # Item
                for w in widgets[1:-2]:
                    row_text.append(f"{w.get()[:5]:>6}")
                amount = widgets[-2].cget("text").replace("₹", "")
                row_text.append(f"{amount:>7}")
                lines.append("".join(row_text)[:width])

        total_str = self.total_label.cget("text").replace("Amount: ₹", "")
        lines += [
            "-" * width,
            f"Total: ₹{total_str}".rjust(width),
            "",
            "I have verified that everything is correct.".center(width),
            "",
            "_" * width,
            "Customer Signature".center(width),
            "\n\n"
        ]
        return lines

    def print_invoice(self):
        """Print the invoice using the configured method."""
        try:
            # Define translation function locally if needed
            translate = lambda x: x  # Default English implementation
            
            customer = self.validate_customer_name(self.customer_entry.get().strip())
            additional_data = {}
            if self.current_mode.get() == "Kata" and hasattr(self, 'kata_amount_var'):
                additional_data["kata_amount"] = validate_float(self.kata_amount_var.get())

            lines = self.generate_print_content()
            mode = load_print_setting()

            if mode == "dialog":
                # Send plain text to Windows print dialog
                print_with_dialog(lines)
                messagebox.showinfo(_("Print"), _("Print dialog opened."))
            elif mode == "pdf":
                try:
                    pdf_path = generate_pdf(
                        self.current_mode.get(), customer, self.rows,
                        self.total_label.cget("text"), MODES, additional_data
                    )
                    if platform.system() == "Windows":
                        os.startfile(pdf_path)
                    else:
                        os.system(f"open {pdf_path}")
                    messagebox.showinfo(_("PDF Generated"), _("PDF has been generated and opened."))
                except Exception as pdf_error:
                    logger.error(f"PDF generation error: {str(pdf_error)}")
                    messagebox.showerror("PDF Error", str(pdf_error))
            else:
                try:
                    direct_print_invoice(
                        self.current_mode.get(), customer, self.rows,
                        self.total_label.cget("text"), MODES, self.config, additional_data
                    )
                    messagebox.showinfo(_("Success"), _("Invoice sent to printer."))
                except Exception as print_error:
                    logger.error(f"Direct print error: {str(print_error)}")
                    messagebox.showerror("Print Error", str(print_error))
        except Exception as e:
            logger.error(f"Print error: {str(e)}")
            messagebox.showerror("Print Error", str(e))

    def save_state_for_undo(self):
        """Save current state for undo."""
        state = {
            "mode": self.current_mode.get(),
            "customer": self.customer_entry.get(),
            "rows": [[w.get() if i < len(widgets) - 2 else w.cget("text") 
                     for i, w in enumerate(widgets[:-1])] for _, widgets in self.rows]
        }
        self.undo_stack.append(state)
        self.redo_stack.clear()
        if len(self.undo_stack) > 50:  # Limit stack size
            self.undo_stack.pop(0)

    def undo(self):
        """Undo the last action."""
        if not self.undo_stack:
            return
        self.redo_stack.append({
            "mode": self.current_mode.get(),
            "customer": self.customer_entry.get(),
            "rows": [[w.get() if i < len(widgets) - 2 else w.cget("text") 
                     for i, w in enumerate(widgets[:-1])] for _, widgets in self.rows]
        })
        state = self.undo_stack.pop()
        self.restore_state(state)

    def redo(self):
        """Redo the last undone action."""
        if not self.redo_stack:
            return
        self.undo_stack.append({
            "mode": self.current_mode.get(),
            "customer": self.customer_entry.get(),
            "rows": [[w.get() if i < len(widgets) - 2 else w.cget("text") 
                     for i, w in enumerate(widgets[:-1])] for _, widgets in self.rows]
        })
        state = self.redo_stack.pop()
        self.restore_state(state)

    def restore_state(self, state):
        """Restore the application state.

        Args:
            state (dict): State to restore.
        """
        self.current_mode.set(state["mode"])
        self.customer_entry.delete(0, "end")
        self.customer_entry.insert(0, state["customer"])
        self.switch_mode()
        self.rows.clear()
        for row_data in state["rows"]:
            self.add_row()
            _, widgets = self.rows[-1]
            for i, value in enumerate(row_data[:-1]):
                if i == 0 and isinstance(widgets[i], ttk.Combobox):
                    widgets[i].set(value)
                else:
                    widgets[i].insert(0, value)
        self.update_amounts()

    def open_settings(self):
        # Use the open_settings_window function from print_utils.py
        open_settings_window(self)

    def show_help(self):
        """Show the help dialog."""
        help_text = _(
            "Invoice Management System\n\n"
            "1. Select a mode (Patti, Kata, Barthe) to change the table format.\n"
            "2. Enter customer name (optional).\n"
            "3. Add rows to enter item details.\n"
            "4. Use 'Delete' (X) to remove a row.\n"
            "5. Save to Excel or print the invoice.\n"
            "6. Use Undo/Redo for actions.\n\n"
            "For support, contact the developer."
        )
        messagebox.showinfo(_("Help"), help_text)

    def on_closing(self):
        """Handle window closing."""
        if messagebox.askyesno(_("Confirm"), _("Do you want to exit?")):
            self.destroy()

if __name__ == "__main__":
    app = InvoiceApp()
    app.mainloop()