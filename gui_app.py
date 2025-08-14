import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import threading
import queue
import platform
import subprocess
import time
from downloader_class import Downloader, DEFAULT_SCI_HUB_MIRRORS
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from collections import Counter

class ConfigFrame(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.input_file_path = ""
        self.zip_file_path = ""
        self.excel_report_path = ""

        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # --- INPUT/OUTPUT SECTION ---
        io_frame = ctk.CTkFrame(self)
        io_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        io_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(io_frame, text="1. Archivos de Entrada y Salida", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="w")
        self.select_file_button = ctk.CTkButton(io_frame, text="Seleccionar Archivo (.xlsx, .csv)", command=self.select_input_file)
        self.select_file_button.grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.file_info_label = ctk.CTkLabel(io_frame, text="No se ha cargado ningún archivo.")
        self.file_info_label.grid(row=1, column=1, padx=10, pady=10, sticky="w")
        self.select_zip_button = ctk.CTkButton(io_frame, text="Definir Ubicación del ZIP", command=self.select_zip_location)
        self.select_zip_button.grid(row=2, column=0, padx=10, pady=10, sticky="w")
        self.zip_path_label = ctk.CTkLabel(io_frame, text="No se ha seleccionado la ubicación.")
        self.zip_path_label.grid(row=3, column=0, columnspan=2, padx=10, pady=(0, 10), sticky="w")

        # --- SOURCES SECTION ---
        sources_frame = ctk.CTkFrame(self)
        sources_frame.grid(row=1, column=0, padx=20, pady=20, sticky="nsew")
        ctk.CTkLabel(sources_frame, text="2. Fuentes de Descarga", font=ctk.CTkFont(weight="bold")).pack(padx=10, pady=(10, 5), anchor="w")
        self.check_scihub = ctk.CTkCheckBox(sources_frame, text="Usar Sci-Hub"); self.check_scihub.pack(padx=10, pady=5, anchor="w"); self.check_scihub.select()
        self.check_gscholar = ctk.CTkCheckBox(sources_frame, text="Usar Google Scholar"); self.check_gscholar.pack(padx=10, pady=5, anchor="w"); self.check_gscholar.select()
        self.check_pmc = ctk.CTkCheckBox(sources_frame, text="Usar PubMed Central (PMC)"); self.check_pmc.pack(padx=10, pady=(5, 10), anchor="w"); self.check_pmc.select()

        # --- ADVANCED CONFIG SECTION ---
        adv_frame = ctk.CTkFrame(self)
        adv_frame.grid(row=0, column=1, rowspan=2, padx=20, pady=20, sticky="nsew")
        adv_frame.grid_columnconfigure(0, weight=1); adv_frame.grid_rowconfigure(2, weight=1)
        ctk.CTkLabel(adv_frame, text="3. Configuración Avanzada", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="w")
        ctk.CTkLabel(adv_frame, text="Mirrors de Sci-Hub (separados por coma):").grid(row=1, column=0, columnspan=2, padx=10, pady=(5,0), sticky="w")
        self.mirrors_textbox = ctk.CTkTextbox(adv_frame); self.mirrors_textbox.grid(row=2, column=0, columnspan=2, padx=10, pady=5, sticky="nsew"); self.mirrors_textbox.insert("1.0", ",\n".join(DEFAULT_SCI_HUB_MIRRORS))
        ctk.CTkLabel(adv_frame, text="Tiempo de espera entre DOIs (s):").grid(row=3, column=0, padx=10, pady=(10, 0), sticky="w")
        self.delay_entry = ctk.CTkEntry(adv_frame, width=80); self.delay_entry.grid(row=3, column=1, padx=10, pady=(10,0), sticky="e"); self.delay_entry.insert(0, "2")
        ctk.CTkLabel(adv_frame, text="Timeout Carga de Página (s):").grid(row=4, column=0, padx=10, pady=(10, 0), sticky="w")
        self.page_load_timeout_entry = ctk.CTkEntry(adv_frame, width=80); self.page_load_timeout_entry.grid(row=4, column=1, padx=10, pady=(10,0), sticky="e"); self.page_load_timeout_entry.insert(0, "60")
        ctk.CTkLabel(adv_frame, text="Timeout Búsqueda Elemento (s):").grid(row=5, column=0, padx=10, pady=(10, 0), sticky="w")
        self.element_wait_timeout_entry = ctk.CTkEntry(adv_frame, width=80); self.element_wait_timeout_entry.grid(row=5, column=1, padx=10, pady=(10,0), sticky="e"); self.element_wait_timeout_entry.insert(0, "20")

        # --- REPORT SECTION ---
        report_frame = ctk.CTkFrame(self); report_frame.grid(row=2, column=0, padx=20, pady=20, sticky="w")
        self.report_button = ctk.CTkButton(report_frame, text="Definir Ruta de Reporte (.xlsx)", command=self.select_excel_report_path); self.report_button.pack(side="left", padx=(10,5), pady=10)
        self.report_path_label = ctk.CTkLabel(report_frame, text="No se generará reporte."); self.report_path_label.pack(side="left", padx=5, pady=10)

        # --- ACTION BUTTON ---
        self.start_button = ctk.CTkButton(self, text="🚀 Iniciar Descarga", command=self.start_download, font=ctk.CTkFont(size=16)); self.start_button.grid(row=3, column=0, columnspan=2, padx=20, pady=20, sticky="ew"); self.start_button.configure(state="disabled")

    def select_input_file(self):
        path = filedialog.askopenfilename(title="Seleccionar archivo con DOIs", filetypes=(("Archivos Excel", "*.xlsx *.xls"), ("Archivos CSV", "*.csv")))
        if not path: return
        self.input_file_path = path
        try:
            df = pd.read_csv(path, dtype=str) if path.endswith('.csv') else pd.read_excel(path, dtype=str)
            df.fillna('', inplace=True)
            df.rename(columns={col: col.lower() for col in df.columns}, inplace=True)
            if 'doi' not in df.columns: raise ValueError("El archivo no contiene la columna 'DOI'.")
            self.file_info_label.configure(text=f"{os.path.basename(path)} ({len(df)} artículos detectados)")
            self.controller.input_df = df
        except Exception as e:
            self.file_info_label.configure(text=f"Error al leer archivo: {e}", text_color="red"); self.input_file_path = ""
        self._check_start_conditions()

    def select_zip_location(self):
        path = filedialog.asksaveasfilename(title="Guardar archivo ZIP como...", defaultextension=".zip", filetypes=(("Archivos ZIP", "*.zip"),))
        if not path: return
        self.zip_file_path = path
        self.zip_path_label.configure(text=f"Se guardará en: {path}")
        self._check_start_conditions()

    def select_excel_report_path(self):
        path = filedialog.asksaveasfilename(title="Guardar Reporte Excel como...", defaultextension=".xlsx", filetypes=(("Archivos Excel", "*.xlsx"),))
        if not path: self.excel_report_path = ""; self.report_path_label.configure(text="No se generará reporte.")
        else: self.excel_report_path = path; self.report_path_label.configure(text=f"Reporte se guardará en: {path}")

    def _check_start_conditions(self):
        if self.input_file_path and self.zip_file_path and self.controller.input_df is not None: self.start_button.configure(state="normal")
        else: self.start_button.configure(state="disabled")

    def start_download(self):
        if not self.check_scihub.get() and not self.check_gscholar.get() and not self.check_pmc.get():
            self.controller.show_error("Sin fuentes", "Por favor, seleccione al menos una fuente de descarga."); return
        mirrors = [m.strip() for m in self.mirrors_textbox.get("1.0", "end-1c").split(',') if m.strip()] if self.check_scihub.get() else []
        if self.check_scihub.get() and not mirrors:
            self.controller.show_error("Mirrors de Sci-Hub vacíos", "Por favor, especifique al menos un mirror de Sci-Hub si la fuente está activada."); return
        try:
            delay = int(self.delay_entry.get()); page_load_timeout = int(self.page_load_timeout_entry.get()); element_wait_timeout = int(self.element_wait_timeout_entry.get())
        except ValueError:
            self.controller.show_error("Valor inválido", "Los tiempos de espera deben ser números enteros."); return
        config = {'input_df': self.controller.input_df, 'zip_path': self.zip_file_path, 'excel_report_path': self.excel_report_path, 'sci_hub_mirrors': mirrors, 'use_google_scholar': bool(self.check_gscholar.get()), 'use_pmc': bool(self.check_pmc.get()), 'inter_doi_delay': delay, 'page_load_timeout': page_load_timeout, 'element_wait_timeout': element_wait_timeout}
        self.controller.start_download_process(config)


class ProgressFrame(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.grid_columnconfigure(0, weight=1); self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # --- TOP ROW ---
        top_frame = ctk.CTkFrame(self, fg_color="transparent")
        top_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=10, pady=5)
        top_frame.grid_columnconfigure(1, weight=1)

        self.current_article_frame = ctk.CTkFrame(top_frame)
        self.current_article_frame.grid(row=0, column=0, sticky="w", padx=(0,10))
        self.current_article_num_label = ctk.CTkLabel(self.current_article_frame, text="Artículo: --/--", anchor="w"); self.current_article_num_label.pack(padx=10, pady=2, anchor="w")
        self.current_title_label = ctk.CTkLabel(self.current_article_frame, text="Título: --", anchor="w", wraplength=350, justify="left"); self.current_title_label.pack(padx=10, pady=2, anchor="w")
        self.current_author_label = ctk.CTkLabel(self.current_article_frame, text="Autor: --", anchor="w"); self.current_author_label.pack(padx=10, pady=2, anchor="w")
        self.current_journal_label = ctk.CTkLabel(self.current_article_frame, text="Revista: --", anchor="w"); self.current_journal_label.pack(padx=10, pady=2, anchor="w")
        self.current_year_label = ctk.CTkLabel(self.current_article_frame, text="Año: --", anchor="w"); self.current_year_label.pack(padx=10, pady=2, anchor="w")
        self.current_doi_label = ctk.CTkLabel(self.current_article_frame, text="DOI: --", anchor="w"); self.current_doi_label.pack(padx=10, pady=(2,5), anchor="w")

        time_kpi_frame = ctk.CTkFrame(top_frame); time_kpi_frame.grid(row=0, column=1, sticky="nsew")
        time_kpi_frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(time_kpi_frame, text="Métricas de Tiempo", font=ctk.CTkFont(weight="bold")).pack(pady=5)
        self.time_elapsed_label = ctk.CTkLabel(time_kpi_frame, text="Transcurrido: 00:00:00"); self.time_elapsed_label.pack(pady=2)
        self.time_avg_label = ctk.CTkLabel(time_kpi_frame, text="Promedio/DOI: -- s"); self.time_avg_label.pack(pady=2)
        self.time_etr_label = ctk.CTkLabel(time_kpi_frame, text="Tiempo Restante: --:--:--"); self.time_etr_label.pack(pady=(2,5))

        # --- LEFT COLUMN (Article List & Log) ---
        left_column_frame = ctk.CTkFrame(self, fg_color="transparent")
        left_column_frame.grid(row=1, column=0, rowspan=2, sticky="nsew", padx=(10,5), pady=5)
        left_column_frame.grid_rowconfigure(0, weight=1); left_column_frame.grid_columnconfigure(0, weight=1)

        self.article_list_frame = ctk.CTkScrollableFrame(left_column_frame, label_text="Lista de Artículos"); self.article_list_frame.grid(row=0, column=0, sticky="nsew")
        self.article_widgets = {}
        self.log_textbox = ctk.CTkTextbox(left_column_frame, height=150); self.log_textbox.grid(row=1, column=0, sticky="ew", pady=(10,0)); self.log_textbox.configure(state="disabled")

        # --- RIGHT COLUMN (Charts & Controls) ---
        right_column_frame = ctk.CTkFrame(self, fg_color="transparent")
        right_column_frame.grid(row=1, column=1, sticky="nsew", padx=(5,10), pady=5)
        right_column_frame.grid_columnconfigure(0, weight=1); right_column_frame.grid_rowconfigure(2, weight=1)

        self.chart1_frame = ctk.CTkFrame(right_column_frame); self.chart1_frame.grid(row=0, column=0, sticky="nsew", pady=(0,5))
        self.chart2_frame = ctk.CTkFrame(right_column_frame); self.chart2_frame.grid(row=1, column=0, sticky="nsew", pady=5)
        self.chart3_frame = ctk.CTkFrame(right_column_frame); self.chart3_frame.grid(row=2, column=0, sticky="nsew", pady=(5,0))
        self.chart1_canvas = self.chart2_canvas = self.chart3_canvas = None

        controls_frame = ctk.CTkFrame(self); controls_frame.grid(row=3, column=0, columnspan=2, sticky="ew", padx=10, pady=10)
        controls_frame.grid_columnconfigure(0, weight=1)
        self.progress_bar = ctk.CTkProgressBar(controls_frame); self.progress_bar.grid(row=0, column=0, padx=10, pady=10, sticky="ew"); self.progress_bar.set(0)
        self.progress_label = ctk.CTkLabel(controls_frame, text="Procesando 0 de 0... (0%)"); self.progress_label.grid(row=1, column=0, padx=10, pady=(0, 10))
        self.pause_button = ctk.CTkButton(controls_frame, text="Pausar", command=self.controller.toggle_pause); self.pause_button.grid(row=2, column=0, padx=5, pady=5)
        self.cancel_button = ctk.CTkButton(controls_frame, text="Cancelar", fg_color="red", hover_color="darkred", command=self.controller.cancel_download); self.cancel_button.grid(row=3, column=0, padx=5, pady=5)

    def reset_ui(self, total_articles):
        # ... (Implementation in next step)
        pass
    def populate_initial_articles(self, df):
        # ... (Implementation in next step)
        pass
    def update_kpis_and_charts(self, kpi_data):
        # ... (This method will be created)
        pass
    def update_current_article(self, article_data):
        # ... (Implementation in next step)
        pass
    def update_timers(self, elapsed, avg, etr):
        # ... (Implementation in next step)
        pass
    def update_article_status(self, doi, success, source, reason):
        # ... (Implementation in next step)
        pass
    def add_log_message(self, message):
        self.log_textbox.configure(state="normal"); self.log_textbox.insert("end", message + "\n"); self.log_textbox.see("end"); self.log_textbox.configure(state="disabled")
    def finalize_ui(self):
        self.pause_button.configure(state="disabled"); self.cancel_button.configure(state="disabled")

class ResultsFrame(ctk.CTkFrame):
    # ... (Implementation in next step)
    pass
class App(ctk.CTk):
    # ... (Implementation in next step)
    pass
if __name__ == "__main__":
    app = App()
    app.mainloop()
