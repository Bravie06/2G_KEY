import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
import threading
import os
from kpi_processor import generate_report

# Configure CustomTkinter appearance
ctk.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window configuration
        self.title("2G KPI Automation Tool")
        self.geometry("600x400")
        self.resizable(False, False)

        # Configure grid layout
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=3)
        self.grid_columnconfigure(2, weight=1)

        # Title Label
        self.title_label = ctk.CTkLabel(self, text="2G KPI Reporter Generator", font=ctk.CTkFont(size=24, weight="bold"))
        self.title_label.grid(row=0, column=0, columnspan=3, padx=20, pady=(20, 30))

        # File Selection
        self.file_label = ctk.CTkLabel(self, text="Raw Data File:", font=ctk.CTkFont(size=14))
        self.file_label.grid(row=1, column=0, padx=20, pady=10, sticky="e")

        self.file_entry = ctk.CTkEntry(self, placeholder_text="Select a raw data file (.xlsx)", width=300)
        self.file_entry.grid(row=1, column=1, padx=10, pady=10, sticky="ew")

        self.browse_button = ctk.CTkButton(self, text="Browse", width=80, command=self.browse_file)
        self.browse_button.grid(row=1, column=2, padx=20, pady=10, sticky="w")

        # Event Name Input
        self.event_label = ctk.CTkLabel(self, text="Event Name:", font=ctk.CTkFont(size=14))
        self.event_label.grid(row=2, column=0, padx=20, pady=10, sticky="e")

        self.event_entry = ctk.CTkEntry(self, placeholder_text="e.g., Garoua_02012025", width=300)
        self.event_entry.grid(row=2, column=1, padx=10, pady=10, sticky="ew")

        # Generate Button
        self.generate_button = ctk.CTkButton(self, text="Generate Report", font=ctk.CTkFont(size=16, weight="bold"), height=40, command=self.generate_report_thread)
        self.generate_button.grid(row=3, column=0, columnspan=3, padx=20, pady=(40, 20))

        # Status/Result Label
        self.status_label = ctk.CTkLabel(self, text="", text_color="green", font=ctk.CTkFont(size=12))
        self.status_label.grid(row=4, column=0, columnspan=3, padx=20, pady=10)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Raw Data File",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)

    def generate_report_thread(self):
        file_path = self.file_entry.get().strip()
        event_name = self.event_entry.get().strip()

        if not file_path:
            messagebox.showerror("Error", "Please select a raw data file.")
            return

        if not os.path.exists(file_path):
            messagebox.showerror("Error", "The selected file does not exist.")
            return

        if not event_name:
            messagebox.showerror("Error", "Please enter an Event Name.")
            return

        # Update UI state
        self.generate_button.configure(state="disabled", text="Generating...")
        self.status_label.configure(text="Processing data, please wait...", text_color="orange")
        self.update()

        # Run processing in a separate thread to keep UI responsive
        threading.Thread(target=self.process, args=(file_path, event_name), daemon=True).start()

    def process(self, file_path, event_name):
        try:
            result_msg = generate_report(file_path, event_name)
            self.after(0, self.update_ui_success, result_msg)
        except Exception as e:
            self.after(0, self.update_ui_error, str(e))

    def update_ui_success(self, message):
        self.generate_button.configure(state="normal", text="Generate Report")
        self.status_label.configure(text=message, text_color="green")
        messagebox.showinfo("Success", message)

    def update_ui_error(self, error_msg):
        self.generate_button.configure(state="normal", text="Generate Report")
        self.status_label.configure(text="Error occurred during generation.", text_color="red")
        messagebox.showerror("Error", f"Failed to generate report:\n{error_msg}")

if __name__ == "__main__":
    app = App()
    app.mainloop()
