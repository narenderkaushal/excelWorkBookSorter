"""Main entry point for Excel Sheet Sorter (Tkinter version)."""
import tkinter as tk
from ui import ExcelSorterApp

def main():
    """Launches the Tkinter application."""
    root = tk.Tk()
    app = ExcelSorterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
