import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
from core.converter import XliffConverter
import os
import json

class ExcelXliffConverter(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel <-> XLIFF Tool")
        self.converter = XliffConverter()
        self.languages = self._load_languages()

        self.exclude_hidden_var = tk.BooleanVar(value=False) 
        self.blank_if_equal_var = tk.BooleanVar(value=False) 
        self._setup_ui()
        icon_path = os.path.join(os.path.dirname(__file__), "..", "data", "excel2xliff2excel.ico")
        if os.path.exists(icon_path):
            self.iconbitmap(icon_path)
        

    def _load_languages(self):
        base_path = os.path.dirname(os.path.dirname(__file__))
        path = os.path.join(base_path, "data", "languages.json")

        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
                # Alphabetically order languages
                return dict(sorted(data.items(), key=lambda item: item[1]))
        except (FileNotFoundError, json.JSONDecodeError):
            # Fallback
            return {"N/A": "N/A", "N/A": "N/A"}

    def _setup_ui(self):
        # Language Selection
        display_names = list(self.languages.values())

        frame = tk.Frame(self)
        frame.pack(pady=10)
        
        tk.Label(frame, text="Source:").grid(row=0, column=0, sticky="w")
        self.src_lang = ttk.Combobox(frame, values=display_names, width=35)
        self.src_lang.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(frame, text="Target:").grid(row=1, column=0, sticky="w")
        self.tgt_lang = ttk.Combobox(frame, values=display_names, width=35)
        self.tgt_lang.grid(row=1, column=1, padx=5, pady=5)

        export_options_frame = tk.LabelFrame(self, text=" Export to XLIFF", padx=10, pady=10)
        export_options_frame.pack(pady=10, padx=20, fill="x")
        import_options_frame = tk.LabelFrame(self, text=" Import to Excel", padx=10, pady=10)
        import_options_frame.pack(pady=10, padx=20, fill="x")

        self.hide_check = tk.Checkbutton(
            export_options_frame, 
            text="Exclude hidden sheets/rows/columns", 
            variable=self.exclude_hidden_var
        )
        self.hide_check.pack(anchor="w")

        self.blank_check = tk.Checkbutton(
            export_options_frame, 
            text="Leave Target empty if Source and Target strings match", 
            variable=self.blank_if_equal_var
        )
        self.blank_check.pack(anchor="w")

        # Import-Export buttons
        tk.Button(export_options_frame, text="Export Excel → XLIFF", command=self.handle_export).pack(pady=10)
        tk.Button(import_options_frame, text="Import XLIFF → Excel", command=self.handle_import).pack(pady=10)

        self.progress = ttk.Progressbar(self, length=300, mode='determinate')
        self.progress.pack(pady=20)

    def get_iso_code(self, display_name):
        for code, name in self.languages.items():
            if name == display_name:
                return code
        return "N/A"
    
    def update_progress(self, current, total):
        # Bar progress
        self.progress["maximum"] = total
        self.progress["value"] = current
        self.update_idletasks()

    def handle_export(self):
        
        excel_types = [("Excel files", "*.xlsx"), ("All files", "*.*")]
        xliff_types = [("XLIFF files", "*.xliff"), ("All files", "*.*")]

        src_path = filedialog.askopenfilename(title="Select Source Excel",defaultextension=".xlsx",filetypes=excel_types)
        tgt_path = filedialog.askopenfilename(title="Select Target Excel",defaultextension=".xlsx",filetypes=excel_types)
        out_path = filedialog.asksaveasfilename(title="Save XLIFF file", defaultextension=".xliff", filetypes=xliff_types)

        exclude_hidden = self.exclude_hidden_var.get()
        blank_if_equal = self.blank_if_equal_var.get()

        src_lang_iso = self.get_iso_code(self.src_lang.get())
        tgt_lang_iso = self.get_iso_code(self.tgt_lang.get())

        if all([src_path, tgt_path, out_path]):
            def task():
                try:
                    self.converter.export_excel(
                        src_path, 
                        tgt_path, 
                        out_path, 
                        src_lang_iso, 
                        tgt_lang_iso,
                        hidden= exclude_hidden,
                        progress_callback=self.update_progress,
                        blank_if_equal=blank_if_equal
                    )
                    messagebox.showinfo("Success", "Export completed")
                except Exception as e:
                    messagebox.showerror("Error", str(e))

            threading.Thread(target=task).start()

    def handle_import(self):
        excel_types = [("Excel files", "*.xlsx"), ("All files", "*.*")]
        xliff_types = [("XLIFF files", "*.xliff"), ("All files", "*.*")]

        orig_path = filedialog.askopenfilename(title="Select Source Excel",defaultextension=".xlsx",filetypes=excel_types)
        xliff_path = filedialog.askopenfilename(title="Select XLIFF File", defaultextension=".xliff", filetypes=xliff_types)
        out_path = filedialog.asksaveasfilename(title="Save Target Excel file", defaultextension=".xlsx", filetypes=excel_types)
        
        if all([orig_path, xliff_path, out_path]):
            success, errors = self.converter.import_xliff(
                orig_path, 
                xliff_path, 
                out_path,
                progress_callback=self.update_progress)
            if success:
                messagebox.showinfo("Success", "Import completed")
            else:
                messagebox.showerror("Validation Error", "\n".join(errors[:5]))

if __name__ == "__main__":
    app = ExcelXliffConverter()
    app.mainloop()