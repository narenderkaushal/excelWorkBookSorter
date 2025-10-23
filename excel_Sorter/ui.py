"""Tkinter GUI for Excel Sheet Sorter with enhanced features and correct scoping."""
import os
import subprocess
import winsound
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
from excel_operations import ExcelHandler
from sheet_rules import alpha_key, numeric_suffix_key
from worker import BatchWorker
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    DND_FILES = None
    TkinterDnD = None

class ExcelSorterApp:
    """Tkinter GUI for sorting Excel sheets alphabetically."""
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Excel Sheet Sorter")

        # ✅ Add menu bar with About option
        menubar = tk.Menu(self.root)
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="About", command=lambda: messagebox.showinfo(
            "About",
            "Excel Sheet Sorter\nVersion 1.0\nDeveloped by!!!!"
        ))
        menubar.add_cascade(label="Help", menu=help_menu)
        self.root.config(menu=menubar)

        # ✅ Add graceful exit confirmation
        def on_close():
            if messagebox.askokcancel("Exit", "Are you sure you want to exit?"):
                self.root.destroy()
        self.root.protocol("WM_DELETE_WINDOW", on_close)

        # ✅ Apply title icon (title bar)
        try:
            self.root.iconbitmap("titleAppIcon.ico")
        except tk.TclError:
            print("⚠️ Warning: title.ico not found. Default window icon will be used.")

        # ✅ Configure ttk button style once here
        style = ttk.Style()
        style.theme_use("clam")  # ✅ enables custom background colors
        style.configure("TButton", padding=6, relief="raised", font=("Helvetica", 11))
        style.map("TButton", background=[("active", "#d0d0d0")])

        # Fixed window size (prevents maximize) but center it
        self.window_width = 650
        self.window_height = 780
        self._center_window()
        self.root.resizable(False, False)
        self.root.configure(bg="#f0f4f8")  # Soft background
        self.file_path = ""
        self.excel_handler = None
        self.batch_var = tk.BooleanVar()  # Batch mode toggle
        self.log_visible = tk.BooleanVar(value=False)

        style.configure("Green.TButton", relief="raised",
            background="#28a745", foreground="white",
            font=("Helvetica", 11, "bold"), borderwidth=1, focusthickness=3, focuscolor="none")
        style.map("Green.TButton",
            background=[("active", "#218838"), ("disabled", "#9ed5a0")],
            foreground=[("disabled", "#ffffff")])

        style.configure("Orange.TButton", relief="raised",
            background="#ff8c42", foreground="white",
            font=("Helvetica", 11, "bold"), borderwidth=1)
        style.map("Orange.TButton",
            background=[("active", "#e67e22"), ("disabled", "#f2b37a")],
            foreground=[("disabled", "#ffffff")])

        style.configure("Red.TButton", relief="raised",
            background="#dc3545", foreground="white",
            font=("Helvetica", 11, "bold"), borderwidth=1)
        style.map("Red.TButton",
            background=[("active", "#b52a37"), ("disabled", "#f5a3aa")],
            foreground=[("disabled", "#ffffff")])

        style.configure("Blue.TButton", relief="raised",
            background="#007bff", foreground="white",
            font=("Helvetica", 11, "bold"), borderwidth=1)
        style.map("Blue.TButton",
            background=[("active", "#0056b3"), ("disabled", "#8cbdf2")],
            foreground=[("disabled", "#ffffff")])

        # ✅ Keyboard shortcuts
        self.root.bind("<Control-o>", lambda _event: self.browse_file())
        self.root.bind("<Control-s>", lambda _event: self.sort_sheets())
        self.root.bind("<Control-q>", lambda _event: self.root.quit())
        self.setup_ui()

    def _center_window(self):
        """Center window on screen."""
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        x = (screen_w - self.window_width) // 2
        y = (screen_h - self.window_height) // 2
        self.root.geometry(f"{self.window_width}x{self.window_height}+{x}+{y}")

    # ---------------------------- UI SETUP ----------------------------
    def setup_ui(self):
        """Sets up all visual UI components."""
        # ✅ Apply app icon (window)
        icon_path = "appIcon.ico"
        try:
            if os.path.exists(icon_path):
                pil_icon = Image.open(icon_path)
                pil_icon = pil_icon.resize((28, 28), Image.Resampling.LANCZOS)  # Adjust size for your UI
                self.app_icon_img = ImageTk.PhotoImage(pil_icon)
            else:
                raise FileNotFoundError
        except tk.TclError:
            # Load fallback/default icon. You can use a bundled image or built-in default.
            pil_icon = Image.new("RGBA", (28, 28), (240,240,240,0))
            self.app_icon_img = ImageTk.PhotoImage(pil_icon)
            print("⚠️ Warning: app.ico not found. Default window icon will be used.")

        # App Title
        title_label = tk.Label(self.root, text="Excel Sheet Sorter",
            font=("Helvetica", 18, "bold"), bg="#f0f4f8",
            fg="#0056b3", image=self.app_icon_img, compound="left", padx=14)
        title_label.pack(pady=(10,4))
        ttk.Separator(self.root, orient="horizontal").pack(fill="x", padx=24, pady=(0, 6))

        # ---------- Begin replacement: compact file-selection + full-height Drop/Browse ----------
        frame_file = tk.Frame(self.root, bg="#dfe6ee", bd=2, relief="groove")
        frame_file.pack(pady=8, padx=20, fill="x")

        # Left content (expands and determines the height)
        content_frame = tk.Frame(frame_file, bg="#dfe6ee")
        content_frame.pack(side="left", fill="both", expand=True, padx=(8, 8), pady=8)

        # Right upload area (will match content_frame height)
        upload_frame = tk.Frame(frame_file, bg="#f5f7fb", bd=2, relief="solid", width=220)
        upload_frame.pack(side="right", fill="y", padx=(10, 10), pady=6)
        upload_frame.pack_propagate(False)  # don't shrink to child size

        # --- Content frame layout (left) ---
        # Section label
        tk.Label(content_frame, text="Select Excel File:", bg="#dfe6ee",
                 font=("Helvetica", 12, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 6))

        # Checkboxes row (grouped horizontally)
        checkbox_frame = tk.Frame(content_frame, bg="#dfe6ee")
        checkbox_frame.grid(row=1, column=0, sticky="w")

        self.batch_var = tk.BooleanVar(value=False)
        batch_check = tk.Checkbutton(checkbox_frame, text="Enable Batch Mode",
                            variable=self.batch_var, bg="#dfe6ee", font=("Helvetica", 10))
        batch_check.pack(side="left", padx=(0, 12))

        self.preview_var = tk.BooleanVar(value=False)
        preview_check = tk.Checkbutton(checkbox_frame, text="Preview Only (don't save changes)",
                            variable=self.preview_var, bg="#dfe6ee", font=("Helvetica", 10))
        preview_check.pack(side="left")

        # Helper text under the checkboxes
        tk.Label(content_frame,
                 text="Supported formats: .xlsx (recommended). Close file in Excel if open!",
                 bg="#dfe6ee", font=("Helvetica", 9), fg="#555").grid(row=2, column=0, sticky="w", pady=(8, 0))

        # Make sure grid expands properly inside content_frame
        content_frame.grid_rowconfigure(0, weight=0)
        content_frame.grid_rowconfigure(1, weight=0)
        content_frame.grid_rowconfigure(2, weight=1)
        content_frame.grid_columnconfigure(0, weight=1)

        # --- Upload label inside upload_frame (right) ---
        self.upload_label = tk.Label(upload_frame,
            text="📂 Drop file / Browse",
            bg="#f5f7fb",
            fg="#007bff",
            font=("Helvetica", 12, "bold"),
            padx=10, pady=8, anchor="center", justify="center")
        self.upload_label.pack(fill="both", expand=True)
        self.upload_label.bind("<Button-1>", lambda e: self.browse_file())

        # Hover effect
        def _on_upload_enter(_):
            upload_frame.config(bg="#eaf3ff")
            self.upload_label.config(fg="#b72b80") #0056b3

        def _on_upload_leave(_):
            upload_frame.config(bg="#f5f7fb")
            self.upload_label.config(fg="#007bff")

        self.upload_label.bind("<Enter>", _on_upload_enter)
        self.upload_label.bind("<Leave>", _on_upload_leave)

        # Register drag-and-drop on label (if available)
        try:
            if DND_FILES is not None:
                self.upload_label.drop_target_register(DND_FILES)
                self.upload_label.dnd_bind("<<Drop>>", self._on_drop)
        except Exception:
            self._log("[WARN] Drag-and-drop not available or registration failed.")
        # ---------- End replacement ----------


        # Helper Text
        tk.Label(
            frame_file,
            text="Supported formats: .xlsx (recommended). Close file in Excel if open!",
            bg="#e1e7ed",
            font=("Helvetica", 9),
            fg="#555",
        ).pack(anchor="w", padx=12, pady=(0, 6))

        # -------------------- Sheet Preview Frame --------------------
        frame_sheet = tk.Frame(self.root, bg="#dfe6ee", bd=2, relief="groove")
        frame_sheet.pack(pady=8, padx=20, fill="x", expand=False)

        tk.Label(frame_sheet, text="Sheets in Workbook:", bg="#dfe6ee", font=("Helvetica", 12)).pack(
            pady=6, anchor="w", padx=10
        )

        # Search Bar
        search_frame = tk.Frame(frame_sheet, bg="#e1e7ed")
        search_frame.pack(fill="x", padx=10)
        tk.Label(search_frame, text="Search Sheet:", bg="#e1e7ed").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(search_frame, textvariable=self.search_var, width=44)
        self.search_entry.pack(side="left", padx=(2, 0))
        self.search_entry.bind("<KeyRelease>", self._filter_sheets)

        listbox_container = tk.Frame(frame_sheet, bg="#e1e7ed")
        listbox_container.pack(padx=6, pady=(2, 4), fill="x", expand=False)
        #listbox_container.pack(padx=5, pady=(0, 2), fill="both", expand=False, ipady=8)

        self.sheet_listbox = tk.Listbox(
            listbox_container,
            selectmode="single",
            font=("Arial", 11),
            bg="white",
            activestyle="dotbox",
            height=10,
            width=67
        )
        self.sheet_listbox.pack(side="left", fill="x", expand=False)
        # Ensure scrollbar renders correctly and stays visible
        scrollbar = tk.Scrollbar(
            listbox_container,
            orient="vertical",
            command=self.sheet_listbox.yview,
            troughcolor="#d6dce2",  # gives contrast so arrows stay visible
            bg="#c5ccd3",
            activebackground="#aab3bb",
        )
        scrollbar.pack(side="right", fill="y")

        # Correctly link listbox and scrollbar before drawing
        self.sheet_listbox.config(yscrollcommand=scrollbar.set)
        self.sheet_listbox.update_idletasks()

        # Sheet Summary
        self.sheet_summary = tk.Label(
            frame_sheet, text="", bg="#e1e7ed", font=("Helvetica", 10, "italic"), fg="#333"
        )
        self.sheet_summary.pack(anchor="w", padx=10, pady=(2, 6))

        # ---------------- Advanced options (small) ---------------
        adv_frame = tk.Frame(frame_sheet, bg="#dfe6ee")
        adv_frame.pack(fill="x", padx=10, pady=(4, 6))

        # Sort mode dropdown
        tk.Label(adv_frame, text="Sort mode:", bg="#dfe6ee").pack(side="left")
        self.sort_mode_var = tk.StringVar(value="alpha")
        sort_menu = ttk.OptionMenu(
            adv_frame, self.sort_mode_var, "alpha",
            "alpha", "reverse_alpha", "numeric_suffix")
        sort_menu.pack(side="left", padx=(6, 10))

        # Template rename entry
        tk.Label(adv_frame, text="Rename template:", bg="#dfe6ee").pack(side="left")
        self.rename_template_var = tk.StringVar(value="{title}")
        tk.Entry(adv_frame, textvariable=self.rename_template_var, width=20).pack(
            side="left", padx=(6, 10))

        # Background worker checkbox
        self.bg_var = tk.BooleanVar(value=False)
        tk.Checkbutton(adv_frame, text="Run in background",
                       variable=self.bg_var, bg="#dfe6ee").pack(side="right")

        # -------------------- Action Buttons --------------------
        frame_actions = tk.Frame(self.root, bg="#f0f4f8")
        frame_actions.pack(pady=8)
        for widget in frame_actions.winfo_children():
            widget.grid_configure(pady=2)

        self.btn_sort = ttk.Button(frame_actions, text="✅ Sort Workbook", style="Green.TButton", width=18,
            command=self.sort_sheets)
        self.btn_sort.grid(row=0, column=0, padx=4)
        self.add_tooltip(self.btn_sort, "Sorts sheets alphabetically in the selected workbook(s).")
        ttk.Button(frame_actions, text="🔁 Clear", width=12, style="Orange.TButton",
            command=self.clear_selection).grid(row=0, column=1, padx=4)
        ttk.Button(frame_actions, text="❌ Exit", style="Red.TButton", width=12,
            command=self.root.quit).grid(row=0, column=2, padx=4)

        # ✅ Add graceful exit confirmation
        def on_close():
            if messagebox.askokcancel("Exit", "Are you sure you want to exit?"):
                self.root.destroy()

        # -------------------- Progress Bar --------------------
        self.progress = ttk.Progressbar(self.root, mode="indeterminate")
        self.progress.pack(fill="x", padx=20, pady=(0, 10))
        self.status_label = tk.Label(
            self.root,
            text="Ready.",
            bg="#f0f4f8",
            font=("Helvetica", 9),
            fg="#555")
        self.status_label.pack(anchor="w", padx=22, pady=(0, 5))

        # -------------------- Log Panel --------------------
        self.log_frame = tk.LabelFrame(self.root, text="Application Log", bg="#f9fbfd", fg="#555", font=("Segoe UI", 10, "bold"))
        self.log_frame.pack(fill="x", padx=10, pady=(0, 5))

        toggle_btn = ttk.Button(self.log_frame, text="Show / Hide Log", width=17, style="Blue.TButton", command=self.toggle_log)
        toggle_btn.pack(anchor="w", padx=8, pady=(2, 2))
        self.log_text = tk.Text(self.log_frame, height=5, state="disabled", wrap="word")
        self.log_text.pack_forget()
        self.root.update_idletasks()
        self.root.geometry(f"{self.window_width}x{self.root.winfo_reqheight()}")
        self.root.minsize(self.root.winfo_width(), self.root.winfo_height())

    # ---------------------------- EVENT HANDLERS ----------------------------
    def toggle_log(self):
        """Show or hide the log console."""
        if self.log_visible.get():
            self.log_text.pack_forget()
            self.log_visible.set(False)
        else:
            self.log_text.pack(fill="both", expand=True, padx=10, pady=5)
            self.log_visible.set(True)

    def _log(self, msg: str):
        """Write message to log console."""
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.config(state="disabled")
        self.log_text.see(tk.END)

    def _filter_sheets(self, _event=None):
        """Filter sheet names by search input."""
        query = self.search_var.get().lower()
        self.sheet_listbox.delete(0, tk.END)
        if self.excel_handler:
            for name in self.excel_handler.get_sheet_names():
                if query in name.lower():
                    self.sheet_listbox.insert(tk.END, name)

    def _batch_callback(self, idx, total, path, state):
        """
        Handle events from BatchWorker on UI thread.
        state values: started, loaded, locked, sorted, done, finished, error:...
        """
        if state == "started":
            self.status_label.config(text=f"🔄 Starting: {os.path.basename(path)} ({idx}/{total})")
        elif state == "loaded":
            self.status_label.config(text=f"🔄 Loaded: {os.path.basename(path)}")
        elif state == "locked":
            self.status_label.config(text=f"⚠️ Locked: {os.path.basename(path)}")
        elif state == "sorted":
            self.status_label.config(text=f"✅ Sorted: {os.path.basename(path)}")
        elif state.startswith("error"):
            self.status_label.config(text=f"❌ Error: {os.path.basename(path)}")
            self._log(f"[ERROR] {state} for {path}")
        elif state == "finished":
            self.status_label.config(text="✅ Batch finished.")
            self.progress["value"] = total
        # update progress bar if available
        try:
            self.progress["value"] = idx
        except (tk.TclError, KeyError):
            self.log("[WARN] Failed to update progress bar.")
            pass

    def add_tooltip(self, widget, text):
        tip = tk.Toplevel(widget)
        tip.withdraw()
        tip.overrideredirect(True)
        label = tk.Label(tip, text=text, bg="#ffffe0", relief="solid", borderwidth=1, font=("Segoe UI", 9))
        def show(event):
            x, y, _, _ = widget.bbox("insert")
            x += widget.winfo_rootx() + 25
            y += widget.winfo_rooty() + 20
            tip.geometry(f"+{x}+{y}")
            label.pack()
            tip.deiconify()
        def hide(event):
            tip.withdraw()
        widget.bind("<Enter>", show)
        widget.bind("<Leave>", hide)

    # ---------------------------- MAIN ACTIONS ----------------------------
    def _load_paths(self, paths: list):
        """Shared loader for a list of file paths.
        Both browse_file() and drag-and-drop handler call this."""
        if not paths:
            return

        # Ensure proper list type and normalize
        paths = list(paths)

        # ✅ Store selected paths (list)
        self.file_path = paths

        # ✅ Clear any previous listbox data
        self.sheet_listbox.delete(0, tk.END)

        # ✅ Handle single vs multiple file preview
        if len(paths) == 1:
            selected_path = paths[0]
            self.excel_handler = ExcelHandler(selected_path)

            loaded = self.excel_handler.load_workbook()
            if getattr(self.excel_handler, "file_open_locked", False):
                self.excel_handler = None
                self.sheet_listbox.delete(0, tk.END)
                messagebox.showwarning(
                    "File Open",
                    "The selected file appears to be open in Excel.\nPlease close it and try again."
                )
                return

            if loaded:
                sheets = self.excel_handler.get_sheet_names()
                for s in sheets:
                    self.sheet_listbox.insert(tk.END, s)
                # file size display (existing behavior)
                try:
                    size_kb = os.path.getsize(selected_path) / 1024.0
                except OSError:
                    size_kb = 0.0
                self.sheet_summary.config(text=f"Sheets detected: {len(sheets)} | File size: {size_kb:.1f} KB")
            else:
                messagebox.showerror("Load Error", "Failed to load workbook. Check console logs.")
        else:
            # Batch Mode: show file names in the listbox instead of sheets
            self.excel_handler = None
            for p in paths:
                self.sheet_listbox.insert(tk.END, os.path.basename(p))
            self.sheet_summary.config(text=f"{len(paths)} files selected (batch mode)")
            self._log(f"[INFO] Batch mode: {len(paths)} files selected")

        # ✅ Update display
        self.sheet_listbox.update_idletasks()
        self.sheet_listbox.yview_moveto(0)
        self.root.after(200, lambda: self.sheet_listbox.yview_moveto(0))

        filename_display = "; ".join(os.path.basename(p) for p in paths)
        self.upload_label.config(text=f"✅ Loaded: {filename_display}", fg="#28a745")

        # ✅ Enable/disable search field depending on mode
        if len(paths) == 1:
            self.search_entry.config(state="normal")
        else:
            self.search_entry.config(state="disabled")

            self.status_label.config(text="✅ File(s) loaded successfully.")

    def browse_file(self):
        """Handles single or multi-file selection."""
        filetypes = [("Excel Files", "*.xlsx *.xls")]

        # ✅ Get paths and normalize to list
        if self.batch_var.get():
            paths = list(filedialog.askopenfilenames(title="Select Excel Files", filetypes=filetypes))
        else:
            path = filedialog.askopenfilename(title="Select Excel File", filetypes=filetypes)
            paths = [path] if path else []

        if not paths:
            return
        # ✅ Unified loader (Browse + Drag & Drop)
        self._load_paths(paths)

    def _on_drop(self, event):
        """Handle dropped files. event.data is a string of file paths;
        on Windows it may be like '{C:/path/file1.xlsx} {C:/path/file2.xlsx}'
        We parse carefully and then call _load_paths(paths)."""
        data = event.data
        if not data:
            return
        paths = []
        cur = ""
        in_brace = False
        for ch in data:
            if ch == '{':
                in_brace = True
                cur = ""
            elif ch == '}':
                in_brace = False
                if cur:
                    paths.append(cur)
                cur = ""
            elif ch == ' ' and not in_brace:
                if cur:
                    paths.append(cur)
                    cur = ""
            else:
                cur += ch
        if cur:
            paths.append(cur)

        # filter non-files and only keep Excel files
        paths = [p for p in paths if p and os.path.isfile(p) and p.lower().endswith(('.xlsx', '.xls'))]

        if not paths:
            messagebox.showinfo("Drop Files", "No valid Excel files detected in the drop.")
            return
        # call the common loader
        self._load_paths(paths)

    def sort_sheets(self):
        """Sorts sheets alphabetically and saves workbook(s)."""
        if not self.file_path:
            self.status_label.config(text="⚠️ No file selected.")
            messagebox.showinfo("No File Selected", "Please select one or more Excel files first.")
            return

        paths = list(self.file_path)
        # Determinate progress setup (safe to use paths now)
        self.progress["mode"] = "determinate"
        self.progress["maximum"] = len(paths)
        self.progress["value"] = 0
        self.progress.update()

        if not paths:
            self.root.update()
            messagebox.showwarning("Warning", "No Excel files to process.")
            self.status_label.config(text="⚠️ No files.")
            return

        self.status_label.config(text=f"🔄 Process starting batch ({len(paths)} files)…")
        self.root.update()
        self.progress["value"] = 0
        self.progress.update()
        self.root.after(1700)    # small delay for UI draw

                # Determine key function from UI selection
        mode = getattr(self, "sort_mode_var", None)
        key_func = None
        if mode:
            mode = mode.get()
            if mode == "alpha":
                key_func = alpha_key
            elif mode == "reverse_alpha":
                # wrap alpha_key to invert sort
                def _rev(ws):
                    return tuple([-ord(c) for c in ws.title.lower()][:16])
                key_func = _rev
            elif mode == "numeric_suffix":
                key_func = numeric_suffix_key

        # If background requested, use worker
        if getattr(self, "bg_var", None) and self.bg_var.get():
            def cb(idx, total, path, state):
                # executed in worker thread; schedule UI updates on main thread
                self.root.after(0, lambda: self._batch_callback(idx, total, path, state))

            worker = BatchWorker(paths, ExcelHandler, cb)
            worker.start()
            self._log("[INFO] Batch worker started.")
            return

        try:
            for idx, path in enumerate(paths, start=1):
                # Set per-file progress and status
                self.progress["value"] = idx - 1
                self.progress.update()
                self.status_label.config(text=f"🔄 Sorting: {os.path.basename(path)} ({idx}/{len(paths)})")
                self.root.update_idletasks()

                try:
                    self.excel_handler = ExcelHandler(path)
                    loaded = self.excel_handler.load_workbook()
                except (OSError, AttributeError, RuntimeError) as exc:
                    self._log(f"[ERROR] Exception while initializing ExcelHandler for {path}: {exc}")
                    messagebox.showerror("Error", f"Unexpected error opening {os.path.basename(path)}.\nSee log.")
                    # update progress and continue
                    self.progress["value"] = idx
                    self.progress.update()
                    continue

                # If loaded and a key_func was chosen, use it
                if loaded and key_func is not None:
                    ok = self.excel_handler.apply_custom_sort(key_func)
                    if not ok:
                        self._log(f"[WARN] custom sort failed for {path}")
                    else:
                        self._log(f"[INFO] custom sort applied: {path}")

                # rename if template not default
                tpl = getattr(self, "rename_template_var", None)
                if tpl and tpl.get() and tpl.get() != "{title}":
                    renamed = self.excel_handler.rename_sheets_with_template(tpl.get())
                    if renamed:
                        self._log(f"[INFO] sheets renamed using template for: {path}")
                    else:
                        self._log(f"[WARN] rename failed for: {path}")

                # File locked detection
                if getattr(self.excel_handler, "file_open_locked", False):
                    self._log(f"[WARN] File is open/locked: {path}")
                    self.status_label.config(text="⚠️ File Locked.")
                    self.progress["value"] = idx
                    self.root.update()
                    messagebox.showwarning(
                        "File Open",
                        f"The file '{os.path.basename(path)}' is open in Excel. Close it and retry.")
                    continue    # Skip this file cleanly

                if not loaded:
                    self._log(f"[ERROR] Cannot load: {path}")
                    self.root.update()
                    messagebox.showwarning("Warning", f"Cannot load {os.path.basename(path)}")
                    self.progress["value"] = idx
                    self.root.update_idletasks()
                    continue

                # Try sorting using chosen method; catch exceptions so batch continues
                try:
                    if key_func is not None:
                        # Use the user-selected custom key (numeric/reverse/alpha)
                        success = self.excel_handler.apply_custom_sort(key_func)
                    else:
                        # Default alphabetical sort
                        success = self.excel_handler.sort_sheets_alphabetically()
                except (OSError, AttributeError, RuntimeError) as exc:
                    self._log(f"[ERROR] Exception during sorting for {path}: {exc}")
                    messagebox.showerror("Error", f"Error sorting {os.path.basename(path)}. See log.")
                    self.progress["value"] = idx
                    self.root.update_idletasks()
                    continue

                # Preview-only (dry run): show order but still update progress
                if self.preview_var.get():
                    sheets = self.excel_handler.get_sheet_names()
                    preview_text = "\n".join(sheets)
                    messagebox.showinfo(
                        "Preview Mode",
                        f"The sheets will be sorted as follows:\n\n{preview_text}"
                    )
                    self._log(f"[PREVIEW] Sorted order for {os.path.basename(path)}: {sheets}")
                    # mark this file as processed in progress and continue
                    self.progress["value"] = idx
                    self.root.update_idletasks()
                    continue
                # If sorting succeeded, handle save flow
                if success:
                    try:
                        if messagebox.askyesno(
                            "Save As",
                            f"Do you want to save '{os.path.basename(path)}' as a new file instead of overwriting?"
                        ):
                            new_path = filedialog.asksaveasfilename(
                                title="Save Sorted Workbook As",
                                defaultextension=".xlsx",
                                filetypes=[("Excel Files", "*.xlsx")]
                            )
                            if new_path:
                                saved = self.excel_handler.save_as(new_path)
                                if saved:
                                    self._log(f"[INFO] Saved as: {new_path}")
                                else:
                                    self._log(f"[ERROR] Save-as failed for: {new_path}")
                            else:
                                self._log(f"[INFO] Save-as cancelled for: {path}")
                        else:
                            saved = self.excel_handler.save_workbook()
                            if saved:
                                self._log(f"[INFO] Overwritten: {path}")
                            else:
                                self._log(f"[ERROR] Overwrite failed: {path}")
                    except (PermissionError, OSError) as exc:
                        self._log(f"[ERROR] Exception while saving {path}: {exc}")
                        messagebox.showerror("Error", f"Error saving {os.path.basename(path)}. See log.")

                    # Play success sound (best-effort)
                    try:
                        winsound.MessageBeep(winsound.MB_ICONASTERISK)
                    except Exception:
                        pass

                    # Offer to open in Excel (non-blocking)
                    if messagebox.askyesno("Open File", f"Do you want to open '{os.path.basename(path)}' in Excel?"):
                        try:
                            subprocess.Popen(["start", "", path], shell=True)
                        except (OSError, subprocess.SubprocessError) as exc:
                            self._log(f"[WARN] Could not open file in Excel: {exc}")

                    self._log(f"[INFO] Sorted successfully: {path}")
                else:
                    self._log(f"[ERROR] Failed to sort: {path}")
                    messagebox.showerror("Error", f"Failed to sort or save: {os.path.basename(path)}")

                # Update progress for this file
                self.progress["value"] = idx
                self.root.update_idletasks()

            # All done
            self.progress["value"] = len(paths)
            self.progress.update()
            self.status_label.config(text="✅ All files processed successfully.")
            self.root.update()

        except (OSError, RuntimeError, tk.TclError) as e:
            self._log(f"[FATAL] Unexpected runtime error: {e}")
            messagebox.showerror("Fatal Error", f"Unexpected runtime error.\n{e}")
        finally:
            self.root.after(1200, lambda: self.status_label.config(text="✅ Ready."))
            self.root.update_idletasks()

    def clear_selection(self):
        """Clears current selection, resets file path and preview list."""
        self.file_path = ""
        self.excel_handler = None
        self.upload_label.config(text="📂  Drop Excel file here or Click to Browse", fg="#007bff")
        self.sheet_listbox.delete(0, tk.END)
        self.sheet_summary.config(text="")
        self.batch_var.set(0)
        self.preview_var.set(0)
        self.status_label.config(text="✅ Ready.")
        self.root.update_idletasks()
        messagebox.showinfo("Cleared", "Selection cleared.")
