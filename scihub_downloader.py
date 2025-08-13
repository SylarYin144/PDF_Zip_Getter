import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog, scrolledtext
import threading
from queue import Queue, Empty
import pandas as pd
import requests
import zipfile
import os
import re
import time
import sys
from datetime import datetime
from bs4 import BeautifulSoup
from urllib.parse import urljoin
import json
import xml.etree.ElementTree as ET
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import base64
import itertools

class ConfigurationGUI:
    def __init__(self, master):
        self.master = master
        master.title("Configuración de Sci-Hub Downloader")
        master.geometry("600x600")
        master.rowconfigure(0, weight=1)
        master.columnconfigure(0, weight=1)

        self.config = {}
        self.cancel_event = threading.Event()
        self.total_articles = 0
        self.is_processing = False

        main_frame = tk.Frame(master, padx=10, pady=10)
        main_frame.grid(row=0, column=0, sticky="nsew")
        main_frame.rowconfigure(1, weight=1)
        main_frame.columnconfigure(0, weight=1)

        self.config_frame = tk.Frame(main_frame)
        self.config_frame.grid(row=1, column=0, sticky="nsew")
        self.config_frame.columnconfigure(0, weight=1)
        self.config_frame.rowconfigure(5, weight=1)

        paths_frame = tk.LabelFrame(self.config_frame, text="Rutas de Archivos", padx=10, pady=10)
        paths_frame.grid(row=0, column=0, sticky="ew", pady=5)
        paths_frame.columnconfigure(1, weight=1)

        self.input_path = tk.StringVar()
        self.zip_path = tk.StringVar()
        self.report_path = tk.StringVar()

        tk.Label(paths_frame, text="Archivo de DOIs (Excel/CSV):").grid(row=0, column=0, sticky=tk.W, pady=2)
        tk.Entry(paths_frame, textvariable=self.input_path, width=50).grid(row=0, column=1, padx=5, sticky="ew")
        tk.Button(paths_frame, text="...", command=self.browse_input_file).grid(row=0, column=2)

        tk.Label(paths_frame, text="Guardar ZIP en:").grid(row=1, column=0, sticky=tk.W, pady=2)
        tk.Entry(paths_frame, textvariable=self.zip_path, width=50).grid(row=1, column=1, padx=5, sticky="ew")
        tk.Button(paths_frame, text="...", command=self.browse_zip_file).grid(row=1, column=2)

        tk.Label(paths_frame, text="Guardar Reporte Excel en:").grid(row=2, column=0, sticky=tk.W, pady=2)
        tk.Entry(paths_frame, textvariable=self.report_path, width=50).grid(row=2, column=1, padx=5, sticky="ew")
        tk.Button(paths_frame, text="...", command=self.browse_report_file).grid(row=2, column=2)

        delays_frame = tk.LabelFrame(self.config_frame, text="Retrasos y Tiempos (segundos)", padx=10, pady=10)
        delays_frame.grid(row=2, column=0, sticky="ew", pady=5)

        search_methods_frame = tk.LabelFrame(self.config_frame, text="Métodos de Búsqueda Adicionales", padx=10, pady=10)
        search_methods_frame.grid(row=3, column=0, sticky="ew", pady=5)
        self.use_gs_selenium = tk.BooleanVar(value=True)
        self.use_pmc_selenium = tk.BooleanVar(value=True)
        tk.Checkbutton(search_methods_frame, text="Usar Google Scholar (Navegador - Lento)", variable=self.use_gs_selenium).pack(anchor=tk.W)
        tk.Checkbutton(search_methods_frame, text="Usar PubMed Central (Navegador - Lento)", variable=self.use_pmc_selenium).pack(anchor=tk.W)

        self.inter_doi_delay = tk.StringVar(value="5")
        self.mirror_switch_delay = tk.StringVar(value="3")
        self.network_timeout = tk.StringVar(value="30")
        self.selenium_pause = tk.StringVar(value="5")
        tk.Label(delays_frame, text="Retraso entre DOIs:").grid(row=0, column=0, sticky=tk.W)
        tk.Entry(delays_frame, textvariable=self.inter_doi_delay, width=5).grid(row=0, column=1, sticky=tk.W, padx=5)
        tk.Label(delays_frame, text="Retraso al cambiar de mirror:").grid(row=0, column=2, sticky=tk.W, padx=(10,0))
        tk.Entry(delays_frame, textvariable=self.mirror_switch_delay, width=5).grid(row=0, column=3, sticky=tk.W, padx=5)
        tk.Label(delays_frame, text="Timeout de Red (s):").grid(row=1, column=0, sticky=tk.W, pady=(5,0))
        tk.Entry(delays_frame, textvariable=self.network_timeout, width=5).grid(row=1, column=1, sticky=tk.W, padx=5, pady=(5,0))
        tk.Label(delays_frame, text="Pausa Selenium (s):").grid(row=1, column=2, sticky=tk.W, padx=(10,0), pady=(5,0))
        tk.Entry(delays_frame, textvariable=self.selenium_pause, width=5).grid(row=1, column=3, sticky=tk.W, padx=5, pady=(5,0))

        strategy_frame = tk.LabelFrame(self.config_frame, text="Estrategia de Descarga", padx=10, pady=10)
        strategy_frame.grid(row=4, column=0, sticky="ew", pady=5)
        self.strategy_var = tk.StringVar(value="article_first")
        tk.Radiobutton(strategy_frame, text="Artículo por Artículo (Prueba todos los mirrors por cada artículo)", variable=self.strategy_var, value="article_first").pack(anchor=tk.W)
        tk.Radiobutton(strategy_frame, text="Mirror por Mirror (Prueba un mirror con todos los artículos, luego el siguiente)", variable=self.strategy_var, value="mirror_first").pack(anchor=tk.W)

        mirrors_frame = tk.LabelFrame(self.config_frame, text="Mirrors de Sci-Hub (separados por coma)", padx=10, pady=10)
        mirrors_frame.grid(row=5, column=0, sticky="nsew", pady=5)
        mirrors_frame.rowconfigure(0, weight=1)
        mirrors_frame.columnconfigure(0, weight=1)
        self.mirrors_text = scrolledtext.ScrolledText(mirrors_frame, wrap=tk.WORD, height=5)
        self.mirrors_text.pack(fill="both", expand=True)
        self.mirrors_text.insert(tk.END, "https://sci-hub.se/,https://sci-hub.st/")

        buttons_frame = tk.Frame(self.config_frame)
        buttons_frame.grid(row=6, column=0, sticky="ew", pady=10)
        self.start_button = tk.Button(buttons_frame, text="Iniciar Proceso", command=self.start_process, bg="green", fg="white")
        self.start_button.pack(side=tk.RIGHT, padx=5)
        self.cancel_button = tk.Button(buttons_frame, text="Cancelar", command=self.master.destroy)
        self.cancel_button.pack(side=tk.RIGHT)

        # --- Progress Frame ---
        self.progress_view_container = tk.Frame(main_frame)
        self.progress_view_container.columnconfigure(0, weight=1)
        self.progress_view_container.rowconfigure(0, weight=1)
        canvas = tk.Canvas(self.progress_view_container)
        scrollbar = tk.Scrollbar(self.progress_view_container, orient="vertical", command=canvas.yview)
        self.scrollable_progress_frame = tk.Frame(canvas)
        self.scrollable_progress_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scrollable_progress_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        progress_content_frame = self.scrollable_progress_frame
        progress_content_frame.columnconfigure(0, weight=1)

        tk.Label(progress_content_frame, text="Progreso Total (Artículos Buscados):", font="-weight bold").grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(10, 0))
        self.pbar_searched = ttk.Progressbar(progress_content_frame, length=300)
        self.pbar_searched.grid(row=1, column=0, columnspan=2, sticky=tk.EW, padx=5)
        self.pbar_searched_label_var = tk.StringVar(value="0/0 (0.00%)")
        tk.Label(progress_content_frame, textvariable=self.pbar_searched_label_var).grid(row=2, column=0, columnspan=2, sticky=tk.W, padx=5, pady=(0,10))

        tk.Label(progress_content_frame, text="Éxito (Encontrados de Buscados):", font="-weight bold").grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=(10, 0))
        self.pbar_found_of_searched = ttk.Progressbar(progress_content_frame, length=300)
        self.pbar_found_of_searched.grid(row=4, column=0, columnspan=2, sticky=tk.EW, padx=5)
        self.pbar_found_of_searched_label_var = tk.StringVar(value="0/0 (0.00%)")
        tk.Label(progress_content_frame, textvariable=self.pbar_found_of_searched_label_var).grid(row=5, column=0, columnspan=2, sticky=tk.W, padx=5, pady=(0,10))

        tk.Label(progress_content_frame, text="Éxito (Encontrados del Total):", font="-weight bold").grid(row=6, column=0, columnspan=2, sticky=tk.W, pady=(10, 0))
        self.pbar_found_of_total = ttk.Progressbar(progress_content_frame, length=300)
        self.pbar_found_of_total.grid(row=7, column=0, columnspan=2, sticky=tk.EW, padx=5)
        self.pbar_found_of_total_label_var = tk.StringVar(value="0/0 (0.00%)")
        tk.Label(progress_content_frame, textvariable=self.pbar_found_of_total_label_var).grid(row=8, column=0, columnspan=2, sticky=tk.W, padx=5, pady=(0,10))

        time_stats_frame = tk.Frame(progress_content_frame)
        time_stats_frame.grid(row=9, column=0, columnspan=2, sticky=tk.W, pady=(10,0))
        tk.Label(time_stats_frame, text="Tiempo Transcurrido:", font="-weight bold").grid(row=0, column=0, sticky=tk.W)
        self.elapsed_time_var = tk.StringVar(value="0s")
        tk.Label(time_stats_frame, textvariable=self.elapsed_time_var).grid(row=0, column=1, sticky=tk.W, padx=5)
        tk.Label(time_stats_frame, text="Tiempo Promedio/Artículo:", font="-weight bold").grid(row=1, column=0, sticky=tk.W)
        self.avg_time_var = tk.StringVar(value="0s")
        tk.Label(time_stats_frame, textvariable=self.avg_time_var).grid(row=1, column=1, sticky=tk.W, padx=5)
        tk.Label(time_stats_frame, text="Tiempo Estimado Restante:", font="-weight bold").grid(row=2, column=0, sticky=tk.W)
        self.eta_label_var = tk.StringVar(value="Calculando...")
        tk.Label(time_stats_frame, textvariable=self.eta_label_var).grid(row=2, column=1, sticky=tk.W, padx=5)

        self.current_article_frame = tk.LabelFrame(progress_content_frame, text="Procesando Artículo", padx=10, pady=10)
        self.current_article_frame.grid(row=10, column=0, columnspan=2, sticky="ew", pady=10)
        self.current_article_text = scrolledtext.ScrolledText(self.current_article_frame, height=6, wrap=tk.WORD)
        self.current_article_text.pack(fill="both", expand=True)
        self.current_article_text.insert(tk.END, "Iniciando...")
        self.current_article_text.config(state="disabled")

        self.final_status_label_var = tk.StringVar(value="Proceso en ejecución...")
        status_frame = tk.Frame(progress_content_frame)
        status_frame.grid(row=11, column=0, columnspan=2, sticky="ew")
        tk.Label(status_frame, textvariable=self.final_status_label_var, pady=20, font="-weight bold").pack(side=tk.LEFT)
        self.working_indicator_var = tk.StringVar(value="")
        tk.Label(status_frame, textvariable=self.working_indicator_var, pady=20, font="-weight bold").pack(side=tk.LEFT, padx=5)

        self.stop_button = tk.Button(progress_content_frame, text="DETENER PROCESO", command=self.cancel_process, bg="red", fg="white", font="-weight bold")
        self.stop_button.grid(row=12, column=0, columnspan=2, pady=10, sticky="ew")

        self.source_stats_frame = tk.LabelFrame(progress_content_frame, text="Estadísticas de Origen", padx=10, pady=10)
        self.source_stats_frame.grid(row=13, column=0, columnspan=2, sticky="ew", pady=10)
        self.source_stats_labels = {}

    def process_queue(self):
        self.master.config(cursor="")
        try:
            message = self.progress_queue.get_nowait()
            if message['type'] == 'total':
                self.total_articles = message['value']
                self.pbar_searched['maximum'] = self.total_articles
                self.pbar_found_of_total['maximum'] = self.total_articles
            elif message['type'] == 'progress':
                searched = message['searched']
                found = message['found']
                if self.total_articles > 0:
                    self.pbar_searched['value'] = searched
                    searched_perc = (searched / self.total_articles) * 100
                    self.pbar_searched_label_var.set(f"{searched}/{self.total_articles} ({searched_perc:.2f}%)")
                    if searched > 0:
                        self.pbar_found_of_searched['maximum'] = searched
                        self.pbar_found_of_searched['value'] = found
                        found_of_searched_perc = (found / searched) * 100
                        self.pbar_found_of_searched_label_var.set(f"{found}/{searched} ({found_of_searched_perc:.2f}%)")
                    failed = searched - found
                    max_possible_success = self.total_articles - failed
                    self.pbar_found_of_total['maximum'] = max_possible_success if max_possible_success > 0 else 1
                    self.pbar_found_of_total['value'] = found
                    found_of_total_perc = (found / self.total_articles) * 100
                    max_perc = (max_possible_success / self.total_articles) * 100 if self.total_articles > 0 else 0
                    self.pbar_found_of_total_label_var.set(f"{found}/{self.total_articles} ({found_of_total_perc:.2f}%) (Máx. posible: {max_perc:.2f}%)")
            elif message['type'] == 'done':
                self.is_processing = False
                self.final_status_label_var.set(message['message'])
                self.stop_button.config(state="disabled")
                return
            elif message['type'] == 'current_article':
                self.current_article_text.config(state="normal")
                self.current_article_text.delete("1.0", tk.END)
                self.current_article_text.insert(tk.END, message['value'])
                self.current_article_text.config(state="disabled")
            elif message['type'] == 'time_update':
                self.eta_label_var.set(message['value']['eta'])
                self.avg_time_var.set(message['value']['avg'])
            elif message['type'] == 'source_stat':
                stats = message['stats']
                total_found = sum(stats.values())
                for source, count in stats.items():
                    percentage = (count / total_found) * 100 if total_found > 0 else 0
                    text = f"{source}: {count} ({percentage:.2f}%)"
                    if source not in self.source_stats_labels:
                        self.source_stats_labels[source] = tk.Label(self.source_stats_frame, text=text)
                        self.source_stats_labels[source].pack(anchor=tk.W)
                    else:
                        self.source_stats_labels[source].config(text=text)
            elif message['type'] == 'error':
                self.is_processing = False
                self.final_status_label_var.set(f"ERROR: {message['message']}")
                self.stop_button.config(state="disabled")
                return
        except Empty: pass
        if self.is_processing:
            self.master.after(100, self.process_queue)

    def animate_working_indicator(self, counter=0):
        if not self.is_processing:
            self.working_indicator_var.set("")
            return
        dots = "." * (counter % 4)
        self.working_indicator_var.set(f"Procesando{dots}")
        self.master.after(500, lambda: self.animate_working_indicator(counter + 1))

    def update_elapsed_time(self):
        if not self.is_processing: return
        elapsed_seconds = time.time() - self.process_start_time
        mins, secs = divmod(elapsed_seconds, 60)
        hours, mins = divmod(mins, 60)
        self.elapsed_time_var.set(f"{int(hours)}h {int(mins)}m {int(secs)}s" if hours > 0 else f"{int(mins)}m {int(secs)}s" if mins > 0 else f"{int(secs)}s")
        self.master.after(1000, self.update_elapsed_time)

    def browse_input_file(self):
        path = filedialog.askopenfilename(title="Seleccionar archivo con DOIs", filetypes=(("Excel & CSV", "*.xlsx *.xls *.csv"), ("Archivos Excel", "*.xlsx *.xls"), ("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*")))
        if path: self.input_path.set(path)

    def browse_zip_file(self):
        path = filedialog.asksaveasfilename(title="Guardar archivo ZIP como...", defaultextension=".zip", filetypes=(("Archivos ZIP", "*.zip"),))
        if path: self.zip_path.set(path)

    def browse_report_file(self):
        path = filedialog.asksaveasfilename(title="Guardar Reporte Excel como...", defaultextension=".xlsx", filetypes=(("Archivos Excel", "*.xlsx"),))
        if path: self.report_path.set(path)

    def start_process(self):
        input_file = self.input_path.get()
        zip_file = self.zip_path.get()
        if not input_file or not zip_file:
            messagebox.showerror("Error", "Debe especificar la ruta del archivo de entrada y del archivo ZIP de salida.")
            return
        try:
            inter_doi = int(self.inter_doi_delay.get())
            mirror_switch = int(self.mirror_switch_delay.get())
            network_timeout_val = int(self.network_timeout.get())
            selenium_pause_val = int(self.selenium_pause.get())
            if inter_doi < 0 or mirror_switch < 0 or network_timeout_val < 0 or selenium_pause_val < 0: raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Los valores de tiempo deben ser números enteros no negativos.")
            return
        mirrors = [m.strip() for m in self.mirrors_text.get("1.0", tk.END).split(',') if m.strip()]
        self.config = {
            "input_file_path": input_file, "zip_path": zip_file, "excel_report_path": self.report_path.get(),
            "user_inter_doi_delay": inter_doi, "user_mirror_switch_delay": mirror_switch, "user_defined_mirrors": mirrors,
            "use_gs_selenium": self.use_gs_selenium.get(), "use_pmc_selenium": self.use_pmc_selenium.get(),
            "network_timeout": network_timeout_val, "selenium_pause": selenium_pause_val, "download_strategy": self.strategy_var.get(),
        }
        self.cancel_event.clear()
        self.config_frame.grid_remove()
        self.progress_view_container.grid(row=1, column=0, sticky="nsew")
        self.progress_queue = Queue()
        self.worker_thread = threading.Thread(target=download_pdfs_from_file, args=(self.config, self.progress_queue, self.cancel_event))
        self.worker_thread.start()
        self.is_processing = True
        self.process_start_time = time.time()
        self.master.after(100, self.process_queue)
        self.update_elapsed_time()
        self.animate_working_indicator()

    def cancel_process(self):
        if messagebox.askyesno("Confirmar", "¿Está seguro de que desea detener el proceso?"):
            self.cancel_event.set()

# --- Constants and Helpers ---
STANDARD_USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.127 Safari/537.36'

def clean_filename(title): return re.sub(r'[\\/*?:"<>|]', '_', title)
def get_general_source_name(source_string):
    source_lower = source_string.lower()
    if "sci-hub" in source_lower: return "Sci-Hub"
    if "google scholar" in source_lower: return "Google Scholar"
    if "pmc" in source_lower or "pubmed" in source_lower: return "PubMed Central"
    try:
        domain_match = re.search(r'https?://(?:www\.)?([a-zA-Z0-9.-]+)(?:/|$)', source_string)
        if domain_match: return domain_match.group(1)
    except Exception: pass
    return "Otro"

def write_excel_report(excel_path, successful_data, failed_data, all_logs, original_columns, base_scihub_url, queue):
    if not excel_path: return
    try:
        ob_cols = ['DOI', 'Title', 'Successful_Mirror'] + [c for c in original_columns if c not in ['DOI', 'Title', 'Successful_Mirror']] + ['SciHub_Link']
        fa_cols = ['DOI', 'Title'] + [c for c in original_columns if c not in ['DOI', 'Title', 'Failure_Reason', 'Detailed_Status']] + ['Failure_Reason', 'Detailed_Status', 'SciHub_Link']
        ti_cols = ['DOI', 'Title'] + [c for c in original_columns if c not in ['DOI', 'Title', 'Successful_Mirror', 'Start_Time', 'End_Time', 'Duration_Seconds', 'Detailed_Status', 'Failure_Reason']] + ['Successful_Mirror', 'Start_Time', 'End_Time', 'Duration_Seconds', 'Detailed_Status', 'Failure_Reason', 'SciHub_Link']
        def create_ordered_df(data, cols):
            df_ = pd.DataFrame(data) if data else pd.DataFrame()
            if not df_.empty:
                df_['SciHub_Link'] = df_.apply(lambda r: f"{base_scihub_url}{r.get('DOI', r.get('doi', ''))}" if pd.notna(r.get('DOI', r.get('doi', ''))) else '', axis=1)
            for c in cols:
                if c not in df_.columns: df_[c] = pd.NA
            all_data_keys = set()
            if data: all_data_keys.update(k for item in data for k in item.keys())
            final_cols_ordered = [c for c in cols if c in all_data_keys or c == 'SciHub_Link']
            final_cols_ordered.extend([k for k in all_data_keys if k not in final_cols_ordered and k != 'SciHub_Link'])
            return df_.reindex(columns=final_cols_ordered)
        df_obtenidos = create_ordered_df(successful_data, ob_cols)
        df_fallidos = create_ordered_df(failed_data, fa_cols)
        df_tiempos = create_ordered_df(all_logs, ti_cols)
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            df_fallidos.to_excel(writer, sheet_name='Fallidos', index=False)
            df_obtenidos.to_excel(writer, sheet_name='Obtenidos', index=False)
            df_tiempos.to_excel(writer, sheet_name='Tiempos', index=False)
    except Exception as e:
        print(f"Error guardando reporte Excel: {e}")
        if queue: queue.put({'type': 'error', 'message': f"Error guardando Excel: {e}"})

def print_to_console(message, orig_stdout):
    print(message, file=orig_stdout)

def _get_display_text_for_article(row_data, current_num, total_articles, prefix=""):
    article_progress_str = f"{prefix}Artículo: {current_num}/{total_articles} ({(current_num/total_articles)*100:.2f}%)"
    title_str = f"Título: {row_data.get('Title', 'N/A')}"
    doi_str = f"DOI: {row_data.get('DOI', row_data.get('doi', 'N/A'))}"
    return f"{article_progress_str}\n{title_str}\n{doi_str}"

def _try_download_methods(context, original_row_data, mirrors_to_try):
    pass # Placeholder

def _process_article(context, original_row_data, zf, index, total_articles_in_run):
    pass # Placeholder

def _strategy_article_first(context):
    pass

def _strategy_mirror_first(context):
    pass

def download_pdfs_from_file(config, queue, cancel_event):
    driver = None
    original_stdout = sys.stdout
    temp_pdf_paths = []

    try:
        os.environ['WDM_LOG_LEVEL'] = '0'
        options = webdriver.ChromeOptions()
        options.add_argument('--headless'); options.add_argument('--disable-gpu'); options.add_argument('--no-sandbox'); options.add_argument('--disable-dev-shm-usage'); options.add_argument('--log-level=3')
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
    except Exception as e:
        queue.put({'type': 'error', 'message': f"Error al inicializar WebDriver: {e}"})
        driver = None

    try:
        df = None
        file_extension = os.path.splitext(config["input_file_path"])[1].lower()
        if file_extension in ['.xlsx', '.xls']: df = pd.read_excel(config["input_file_path"])
        elif file_extension == '.csv': df = pd.read_csv(config["input_file_path"])
        else: raise ValueError(f"Tipo de archivo no soportado: {file_extension}")
        original_input_columns = [col for col in df.columns if col not in ['DOI', 'Title']]
    except Exception as e:
        queue.put({'type': 'error', 'message': f"Error al leer archivo: {e}"}); return

    user_defined_mirrors = []
    for mirror_url in config["user_defined_mirrors"]:
        if not mirror_url.startswith(("http://", "https://")): mirror_url = "https://" + mirror_url
        if not mirror_url.endswith('/'): mirror_url += '/'
        user_defined_mirrors.append(mirror_url)

    context = {
        "config": config, "queue": queue, "cancel_event": cancel_event, "driver": driver, "df": df,
        "session": requests.Session(), "zip_path": config["zip_path"], "excel_report_path": config["excel_report_path"],
        "original_stdout": original_stdout, "temp_pdf_paths": temp_pdf_paths, "all_articles_log": [],
        "successful_articles_data": [], "failed_articles_data": [], "total_articles": len(df), "successful_downloads": 0,
        "total_downloaded_size_bytes": 0, "source_stats": {}, "process_start_time": time.time(),
        "user_defined_mirrors": user_defined_mirrors, "original_input_columns": original_input_columns
    }
    context['session'].headers.update({'User-Agent': STANDARD_USER_AGENT})

    queue.put({'type': 'total', 'value': context['total_articles']})
    if config["excel_report_path"]:
        write_excel_report(config["excel_report_path"], [], [], [], df.columns, user_defined_mirrors[0] if user_defined_mirrors else "N/A", queue)

    download_strategy = config.get("download_strategy", "article_first")
    print_to_console(f"Iniciando proceso con estrategia: {download_strategy}", original_stdout)

    if download_strategy == 'mirror_first':
        _strategy_mirror_first(context)
    else:
        _strategy_article_first(context)

    # CORRECTLY PLACED FINAL REPORTING LOGIC
    successful_dois = {str(item.get('DOI', item.get('doi', ''))).strip() for item in context['successful_articles_data']}
    doi_col_name = 'DOI' if 'DOI' in df.columns else 'doi'
    df[doi_col_name] = df[doi_col_name].astype(str).str.strip()
    final_failed_df = df[~df[doi_col_name].isin(successful_dois)]
    final_failed_articles_data = final_failed_df.to_dict('records')

    for item in final_failed_articles_data:
        item['Failure_Reason'] = item.get('Failure_Reason', 'No descargado (proceso cancelado o no alcanzado)')
        item['Detailed_Status'] = item.get('Detailed_Status', 'Not Attempted or Cancelled')

    write_excel_report(config["excel_report_path"], context['successful_articles_data'], final_failed_articles_data, context['all_articles_log'], df.columns, user_defined_mirrors[0] if user_defined_mirrors else "N/A", queue)

    summary_message = (f"Proceso completado.\n\nDescargas exitosas: {context['successful_downloads']}\nDescargas fallidas: {len(final_failed_articles_data)}\n" f"Tamaño total PDFs: {context['total_downloaded_size_bytes'] / (1024 * 1024):.2f} MB")
    print_to_console("\n" + "="*50, original_stdout); print_to_console(summary_message, original_stdout); print_to_console("="*50, original_stdout)
    queue.put({'type': 'done', 'message': f'Completado: {context["successful_downloads"]}/{context["total_articles"]} descargados.'})

    finally:
        if driver: driver.quit()
        if sys.stdout != original_stdout: sys.stdout = original_stdout
        for temp_path in temp_pdf_paths:
            try: os.remove(temp_path)
            except OSError: pass
        temp_dir_to_check = "temp_scihub_pdfs"
        if os.path.exists(temp_dir_to_check) and not os.listdir(temp_dir_to_check):
            try: os.rmdir(temp_dir_to_check)
            except OSError: pass
        print_to_console("--- Limpieza Finalizada ---", original_stdout)

if __name__ == "__main__":
    root = tk.Tk()
    app = ConfigurationGUI(root)
    root.mainloop()
    print("\nScript finalizado.")
