# 📘 Excel Workbook Sorter  

🧑‍💻 Author

Narender Kaushal (NKau$hal)
Senior Test Consultant | Python Developer
📎 LinkedIn:
https://www.linkedin.com/in/narender-kaushal-b72b075/

*A Windows Desktop Utility for Sorting Excel Sheets with Ease*

Excel Workbook Sorter is a professional, user-friendly Windows application built with **Python (Tkinter)** and powered by **openpyxl**, allowing users to:

- Sort Excel workbook sheets alphabetically  
- Sort sheets by numeric suffix  
- Sort sheets in calendar order (Jan→Dec, Dec→Jan)  
- Rename sheets using templates (e.g., `{title}`, `{i}`, `{index}`)  
- Use batch mode to process multiple files  
- Preview sorting without saving  
- Detect month-based sheets automatically and switch to Calendar Sort  
- Drag & drop Excel files  
- View logs and progress indicators  
- Save outputs as new files or overwrite existing files  

This project ships with a **standalone Windows Installer (EXE)** built using **PyInstaller + Inno Setup**, ensuring users do not need Python installed.

---

## 🚀 Features

### ✔ Sheet Sorting Modes
- **Alphabetical (A→Z)**  
- **Reverse Alphabetical (Z→A)**  
- **Numeric Suffix Sorting** (e.g., `Sheet1`, `Sheet2`, ...)  
- **Calendar Order**  
  - Jan → Dec  
  - Dec → Jan  
- **Automatic Month Detection** (auto-selects Month Sort when applicable)

---

### ✔ Batch Mode Support
Load and sort **multiple Excel files at once**.

---

### ✔ Sheet Name Template Renaming
Use patterns to rename sheets:

| Token     | Description |
|-----------|-------------|
| `{title}` | Original sheet name |
| `{i}`     | Running index |
| `{index}` | Running index |

Example:  
`Report_{i}` → `Report_1`, `Report_2`, …

---

### ✔ Drag-and-Drop File Support
Simply drag `.xlsx` or `.xls` files into the app.

---

### ✔ Logging and Preview
- View execution logs  
- Toggle log panel  
- Preview mode (no saving)

---

### ✔ Professional Installer
Built using **Inno Setup** with:
- Desktop icon
- Start menu entry
- Bundled PDF documentation
- Branded application icons

---

## 🧩 Application Architecture

### **Core Modules**
| File | Purpose |
|------|---------|
| `app.py` | Entry point (TkinterDnD-enabled window) |
| `ui.py` | Full Tkinter GUI with sorting UI and event handlers |
| `excel_operations.py` | Load, sort, rename, save Excel workbooks |
| `sheet_rules.py` | Sorting rules, regex sort, month detection |
| `validator.py` | Sheet name validation helpers |
| `worker.py` | Background thread for batch processing |
| `backup_util.py` | Automatic timestamp-based backup before save |

---

## 🛠 Build & Packaging System

### ✔ PyInstaller  
Uses a dedicated `.spec` file to bundle:

- EXE  
- Icons  
- PDF documentation  
- Dependencies (openpyxl, Pillow, tkinterdnd2, etc.)

### ✔ Inno Setup  
Produces a polished installer:

- Desktop + Start Menu shortcuts  
- File version metadata  
- Custom icons  
- Bundled documentation  

### ✔ Automated Build Script  
`build_excel_sorter.bat` performs:

1. Clean previous build  
2. Install dependencies  
3. Run PyInstaller with `.spec`  
4. Build Installer using Inno Setup  
5. Validate all artifacts  

---

## ⚡ How to Build the Application

### **1. Run the Build Script**
```bat
build_excel_sorter.bat
