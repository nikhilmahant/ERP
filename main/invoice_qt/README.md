# Invoice Application (Qt Version)

A modern invoice management application built with PySide6 (Qt for Python).

## Features

- Modern Qt-based user interface
- Multiple invoice modes (Patti, Kata, Barthe)
- Excel export functionality
- Print preview and printing support
- Kannada text support
- Auto-calculation of amounts
- Customer information management

## Installation

1. Make sure you have Python 3.7+ installed
2. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Running the Application

```bash
python invoice_qt.py
```

## Usage

1. Select the mode (Patti, Kata, or Barthe)
2. Enter customer details
3. Add rows of items with their respective quantities and rates
4. Use the buttons at the bottom to:
   - Add new rows
   - Clear the form
   - Save to Excel
   - Print the invoice

## Dependencies

- PySide6: Modern Qt framework for Python
- openpyxl: Excel file handling
- pywin32: Windows printing support
