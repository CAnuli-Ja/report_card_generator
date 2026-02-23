# Windows Executable Build Instructions

## For Building on Windows

### Prerequisites

1. Install Python 3.7+ from https://www.python.org/downloads/ (make sure to check "Add Python to PATH")

### Step 1: Install Required Packages

Open Command Prompt or PowerShell and run:

```
pip install pandas openpyxl python-docx PyInstaller
```

### Step 2: Build the Executable

Navigate to the directory containing `report_card_gui.py` and run:

```
pyinstaller --onefile --windowed --name "Report_Card_Generator" report_card_gui.py
```

Or use the provided spec file:

```
pyinstaller build_windows_exe.spec
```

### Step 3: Find Your Executable

After the build completes, you'll find your executable at:

```
dist\Report_Card_Generator.exe
```

You can now distribute this `.exe` file to other Windows computers. No Python installation is required on those computers!

---

## What's Inside

The executable includes:

- Full Python environment
- All required dependencies (pandas, openpyxl, python-docx, tkinter)
- Your report card generator GUI application

---

## Using the Executable

1. Double-click `Report_Card_Generator.exe` to launch
2. The app will start with the same GUI as the Python version
3. Select your Excel gradesheet, choose the sheet name, select your Word template, and choose an output folder
4. Click "Generate Report Cards" to process

---

## Building From This Repository

If building on a Mac or Linux:

1. Install Python and required packages
2. Run: `pyinstaller --onefile --windowed --name "Report_Card_Generator" report_card_gui.py`
3. The executable will be platform-specific (macOS, Linux, or Windows depending on your OS)

To build for Windows specifically, you must run the build process on a Windows machine.
