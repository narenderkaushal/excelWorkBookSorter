"""Main entry point for Excel Sheet Sorter (Tkinter version)."""
import tkinter as tk
from tkinterdnd2 import TkinterDnD
from ui import ExcelSorterApp

def main():
    """Launches the Tkinter application."""
    root = TkinterDnD.Tk()
    app = ExcelSorterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
