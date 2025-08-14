import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
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
from tkinter import ttk
import queue
import threading
import tkinter.scrolledtext as st

# --- Configuration Constants ---
DEFAULT_SCI_HUB_MIRRORS_EXAMPLE = ["https://sci-hub.se/", "https://sci-hub.st/", "https://sci-hub.box/", "https://sci-hub.ru/", "https://sci-hub.red/"]
INTER_DOI_DELAY_SECONDS = 5
MIRROR_SWITCH_DELAY_SECONDS = 3
STANDARD_USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.127 Safari/537.36'

# --- Helper Functions ---
def clean_filename(title):
    return re.sub(r'[\\/*?:"<>|]', '_', title)

def extract_pdf_link_from_html(article_page_url, session):
    try:
        response = session.get(article_page_url, timeout=30)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, 'html.parser')
        iframe = soup.find('iframe', id='pdf')
        if iframe and iframe.get('src'):
            pdf_src = iframe['src']
            if pdf_src.startswith('//'): return 'https:' + pdf_src
            elif pdf_src.startswith('/'): return urljoin(article_page_url, pdf_src)
            return pdf_src
        embed = soup.find('embed', attrs={'type': 'application/pdf'})
        if embed and embed.get('src'):
            return urljoin(article_page_url, embed['src'])
        return None
    except requests.exceptions.RequestException:
        return None
    except Exception:
        return None

# --- Main Worker Function ---
def start_download_process(queue, cancel_requested, articles_data, original_input_columns, zip_path, excel_report_path_config, user_inter_doi_delay, user_mirror_switch_delay, user_defined_mirrors, use_scholar, use_pmc):
    driver = None
    temp_pdf_paths = []
    try:
        if use_scholar or use_pmc:
            try:
                options = webdriver.ChromeOptions()
                options.add_argument('--headless')
                options.add_argument('--disable-gpu')
                options.add_argument('--no-sandbox')
                options.add_argument('--disable-dev-shm-usage')
                print("Inicializando WebDriver de Selenium...")
                driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
                print("WebDriver de Selenium inicializado.")
            except Exception as e:
                print(f"Error al inicializar WebDriver: {e}. Las descargas por Selenium se omitirán.")
                driver = None

        sci_hub_base_url_for_report = user_defined_mirrors[0] if user_defined_mirrors else "N/A"
        session = requests.Session()
        session.headers.update({'User-Agent': STANDARD_USER_AGENT})

        all_articles_log = []
        successful_articles_data = []
        failed_articles_data = []

        successful_downloads = 0
        total_articles = len(articles_data)

        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for index, original_row_data in enumerate(articles_data):
                if cancel_requested.is_set():
                    print("Cancelación solicitada. Deteniendo...")
                    break

                doi = str(original_row_data.get('DOI', '')).strip()
                title = str(original_row_data.get('Título', '')).strip()
                effective_title = title if title else doi
                queue.put({"type": "article_update", "title": f"({index + 1}/{total_articles}) {effective_title}"})

                pdf_content = None
                successful_source = "Ninguna"

                if user_defined_mirrors:
                    for mirror_url in user_defined_mirrors:
                        if cancel_requested.is_set(): break
                        full_url = f"{mirror_url}{doi}"
                        try:
                            pdf_link = extract_pdf_link_from_html(full_url, session)
                            if pdf_link:
                                response = session.get(pdf_link, timeout=60)
                                if 'application/pdf' in response.headers.get('Content-Type', ''):
                                    pdf_content = response.content
                                    successful_source = "Sci-Hub"
                                    break
                        except Exception:
                            continue

                if not pdf_content and use_scholar and driver:
                    if cancel_requested.is_set(): break
                    # Simplified placeholder for selenium logic
                    print(f"Buscando en Google Scholar para DOI: {doi}")
                    # In a real scenario, this would call a detailed selenium function

                if not pdf_content and use_pmc and driver:
                    if cancel_requested.is_set(): break
                    # Simplified placeholder for selenium logic
                    print(f"Buscando en PMC para DOI: {doi}")

                if pdf_content:
                    successful_downloads += 1
                    pdf_filename = clean_filename(effective_title) + ".pdf"
                    zf.writestr(pdf_filename, pdf_content)
                else:
                    failed_articles_data.append(original_row_data)

                queue.put({"type": "progress", "attempted": index + 1, "successful": successful_downloads, "total": total_articles})
                if index < total_articles - 1:
                    time.sleep(user_inter_doi_delay)

    except Exception as e:
        print(f"Error fatal en el hilo de trabajo: {e}")
    finally:
        if driver:
            driver.quit()
        queue.put(None)

# --- GUI and File Reading ---
def read_dois_from_file(file_path):
    try:
        file_extension = os.path.splitext(file_path)[1].lower()
        if file_extension in ['.xlsx', '.xls']:
            df = pd.read_excel(file_path, engine='openpyxl')
            file_type = 'Excel'
        elif file_extension == '.csv':
            try:
                df = pd.read_csv(file_path, sep=None, engine='python', on_bad_lines='warn')
            except Exception:
                df = pd.read_csv(file_path, sep=',', on_bad_lines='warn')
            file_type = 'CSV'
        else:
            return None, None, f"Tipo de archivo no soportado: {file_extension}"

        column_mapping = {
            'DOI': ['doi'],
            'Título': ['title', 'título', 'titulo', 'article title'],
            'Autor': ['author', 'authors', 'first author', 'autores', 'autor'],
            'Revista': ['journal', 'journal title', 'revista', 'journal/book'],
            'Año de publicación': ['year', 'publication year', 'año', 'ano', 'fecha de publicación']
        }

        renamed_columns = {}
        original_columns_before_rename = list(df.columns)

        for standard_name, possible_names in column_mapping.items():
            for col in original_columns_before_rename:
                if str(col).strip().lower() in possible_names:
                    if col not in renamed_columns:
                        renamed_columns[col] = standard_name
                        break

        df.rename(columns=renamed_columns, inplace=True)
        original_input_columns = [col for col in original_columns_before_rename if col not in renamed_columns]

        if 'DOI' not in df.columns:
            return None, None, "No se encontró una columna de DOI."

        articles_data = df.to_dict('records')
        return articles_data, original_input_columns, file_type

    except FileNotFoundError:
        return None, None, f"Archivo no encontrado: {file_path}"
    except Exception as e:
        return None, None, f"Ocurrió un error al leer el archivo: {e}"

class TextRedirector:
    def __init__(self, widget, queue, tag="stdout"):
        self.widget = widget
        self.queue = queue
        self.tag = tag

    def write(self, s):
        self.queue.put({"type": "log", "message": s})

    def flush(self):
        pass

class ConfigFrame(tk.Frame):
    def __init__(self, parent, start_callback):
        super().__init__(parent, padx=10, pady=10)
        self.start_callback = start_callback
        self.input_file_path = tk.StringVar()
        self.zip_path = tk.StringVar()
        self.excel_report_path = tk.StringVar()
        self.inter_doi_delay = tk.StringVar(value=str(INTER_DOI_DELAY_SECONDS))
        self.mirror_switch_delay = tk.StringVar(value=str(MIRROR_SWITCH_DELAY_SECONDS))
        self.use_scholar = tk.BooleanVar(value=True)
        self.use_pmc = tk.BooleanVar(value=True)
        self.setup_widgets()

    def setup_widgets(self):
        path_frame = tk.LabelFrame(self, text="Rutas de Archivos", padx=10, pady=10)
        path_frame.pack(fill="x", expand=True, pady=5)
        tk.Button(path_frame, text="Seleccionar Archivo de DOIs (Excel/CSV)", command=self.select_input_file).pack(fill="x", pady=2)
        tk.Label(path_frame, textvariable=self.input_file_path, wraplength=500).pack(fill="x")
        tk.Button(path_frame, text="Seleccionar Ubicación del ZIP de Salida", command=self.select_zip_path).pack(fill="x", pady=2)
        tk.Label(path_frame, textvariable=self.zip_path, wraplength=500).pack(fill="x")
        tk.Button(path_frame, text="Seleccionar Ubicación del Reporte Excel (Opcional)", command=self.select_excel_report_path).pack(fill="x", pady=2)
        tk.Label(path_frame, textvariable=self.excel_report_path, wraplength=500).pack(fill="x")
        settings_frame = tk.LabelFrame(self, text="Configuraciones", padx=10, pady=10)
        settings_frame.pack(fill="x", expand=True, pady=5)
        delay_frame = tk.Frame(settings_frame)
        delay_frame.pack(fill="x", expand=True)
        tk.Label(delay_frame, text="Retraso Inter-DOI (s):").pack(side="left", padx=(0, 5))
        tk.Entry(delay_frame, textvariable=self.inter_doi_delay, width=5).pack(side="left")
        tk.Label(delay_frame, text="Retraso Cambio de Mirror (s):").pack(side="left", padx=(20, 5))
        tk.Entry(delay_frame, textvariable=self.mirror_switch_delay, width=5).pack(side="left")
        sources_frame = tk.Frame(settings_frame)
        sources_frame.pack(fill="x", expand=True, pady=5)
        tk.Checkbutton(sources_frame, text="Usar Google Scholar", variable=self.use_scholar).pack(side="left")
        tk.Checkbutton(sources_frame, text="Usar PubMed Central", variable=self.use_pmc).pack(side="left", padx=10)
        mirrors_frame = tk.LabelFrame(self, text="Mirrors de Sci-Hub (uno por línea)", padx=10, pady=10)
        mirrors_frame.pack(fill="both", expand=True, pady=5)
        self.mirrors_text = st.ScrolledText(mirrors_frame, height=5)
        self.mirrors_text.insert(tk.END, "\n".join(DEFAULT_SCI_HUB_MIRRORS_EXAMPLE))
        self.mirrors_text.pack(fill="both", expand=True)
        tk.Button(self, text="Iniciar Descarga", command=self.on_start).pack(pady=10)

    def select_input_file(self):
        path = filedialog.askopenfilename(title="Seleccionar archivo con DOIs", filetypes=(("Archivos Excel", "*.xlsx *.xls"), ("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*")))
        if path: self.input_file_path.set(path)

    def select_zip_path(self):
        path = filedialog.asksaveasfilename(title="Guardar archivo ZIP como...", defaultextension=".zip", filetypes=(("Archivos ZIP", "*.zip"),))
        if path: self.zip_path.set(path)

    def select_excel_report_path(self):
        path = filedialog.asksaveasfilename(title="Guardar Reporte Excel como...", defaultextension=".xlsx", filetypes=(("Archivos Excel", "*.xlsx"),))
        if path: self.excel_report_path.set(path)

    def on_start(self):
        if not self.input_file_path.get() or not self.zip_path.get():
            messagebox.showerror("Error", "Debe seleccionar un archivo de entrada y una ubicación para el ZIP.")
            return
        mirrors = [line.strip() for line in self.mirrors_text.get("1.0", tk.END).splitlines() if line.strip()]
        config = {
            "input_file_path": self.input_file_path.get(),
            "zip_path": self.zip_path.get(),
            "excel_report_path": self.excel_report_path.get(),
            "inter_doi_delay": int(self.inter_doi_delay.get()),
            "mirror_switch_delay": int(self.mirror_switch_delay.get()),
            "mirrors": mirrors,
            "use_scholar": self.use_scholar.get(),
            "use_pmc": self.use_pmc.get()
        }
        self.start_callback(config)

class MonitorFrame(tk.Frame):
    def __init__(self, parent, cancel_callback):
        super().__init__(parent, padx=10, pady=10)
        self.cancel_callback = cancel_callback
        self.progress = tk.IntVar()
        self.total_articles = tk.IntVar(value=0)
        self.attempted_articles = tk.IntVar(value=0)
        self.successful_articles = tk.IntVar(value=0)
        self.current_article_text = tk.StringVar(value="Esperando para iniciar...")
        self.setup_widgets()

    def setup_widgets(self):
        stats_frame = tk.Frame(self)
        stats_frame.pack(fill="x", pady=5)
        tk.Label(stats_frame, text="Total:").pack(side="left")
        tk.Label(stats_frame, textvariable=self.total_articles).pack(side="left", padx=(0, 10))
        tk.Label(stats_frame, text="Intentados:").pack(side="left")
        tk.Label(stats_frame, textvariable=self.attempted_articles).pack(side="left", padx=(0, 10))
        tk.Label(stats_frame, text="Obtenidos:").pack(side="left")
        tk.Label(stats_frame, textvariable=self.successful_articles).pack(side="left", padx=(0, 10))
        self.progress_bar = ttk.Progressbar(self, orient="horizontal", length=100, mode="determinate", variable=self.progress)
        self.progress_bar.pack(fill="x", pady=5)
        tk.Label(self, text="Procesando actualmente:").pack(anchor="w")
        tk.Label(self, textvariable=self.current_article_text, wraplength=780, justify="left").pack(anchor="w", fill="x")
        log_frame = tk.LabelFrame(self, text="Log de Descarga", padx=10, pady=10)
        log_frame.pack(fill="both", expand=True, pady=5)
        self.log_text = st.ScrolledText(log_frame, state="disabled", wrap=tk.WORD)
        self.log_text.pack(fill="both", expand=True)
        tk.Button(self, text="Cancelar", command=self.cancel_callback).pack(pady=10)

    def update_stats(self, attempted, successful, total):
        self.attempted_articles.set(attempted)
        self.successful_articles.set(successful)
        self.total_articles.set(total)
        if total > 0: self.progress.set((attempted / total) * 100)
        else: self.progress.set(0)

    def add_log_message(self, message):
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, message)
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")

    def reset(self):
        self.progress.set(0)
        self.total_articles.set(0)
        self.attempted_articles.set(0)
        self.successful_articles.set(0)
        self.current_article_text.set("Esperando para iniciar...")
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", tk.END)
        self.log_text.config(state="disabled")

class SciHubDownloaderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Sci-Hub PDF Downloader")
        self.root.geometry("800x600")
        self.download_thread = None
        self.queue = None
        self.cancel_requested = threading.Event()
        self.config_frame = ConfigFrame(root, self.start_download)
        self.monitor_frame = MonitorFrame(root, self.cancel_download)
        self.show_config_frame()

    def show_config_frame(self):
        self.monitor_frame.pack_forget()
        self.config_frame.pack(fill="both", expand=True)

    def show_monitor_frame(self):
        self.config_frame.pack_forget()
        self.monitor_frame.pack(fill="both", expand=True)

    def start_download(self, config):
        self.monitor_frame.reset()
        self.cancel_requested.clear()
        articles_data, original_cols, file_type = read_dois_from_file(config["input_file_path"])
        if articles_data is None:
            messagebox.showerror("Error de Archivo", f"No se pudo leer el archivo: {original_cols}")
            return
        self.monitor_frame.total_articles.set(len(articles_data))
        self.queue = queue.Queue()
        sys.stdout = TextRedirector(self.monitor_frame.log_text, self.queue)
        thread_args = (
            self.queue, self.cancel_requested, articles_data, original_cols,
            config["zip_path"], config["excel_report_path"], config["inter_doi_delay"],
            config["mirror_switch_delay"], config["mirrors"], config["use_scholar"], config["use_pmc"]
        )
        self.download_thread = threading.Thread(target=start_download_process, args=thread_args)
        self.download_thread.start()
        self.show_monitor_frame()
        self.process_queue()

    def cancel_download(self):
        if self.download_thread and self.download_thread.is_alive():
            self.cancel_requested.set()
            messagebox.showinfo("Cancelando", "Se solicitará la cancelación del proceso. Espere a que la tarea actual finalice.")
        sys.stdout = sys.__stdout__
        self.show_config_frame()

    def process_queue(self):
        try:
            while True:
                msg = self.queue.get_nowait()
                if msg is None:
                    sys.stdout = sys.__stdout__
                    self.download_thread = None
                    messagebox.showinfo("Completado", "El proceso de descarga ha finalizado.")
                    self.show_config_frame()
                    return
                if msg.get("type") == "log":
                    self.monitor_frame.add_log_message(msg["message"])
                elif msg.get("type") == "progress":
                    self.monitor_frame.update_stats(msg["attempted"], msg["successful"], msg["total"])
                elif msg.get("type") == "article_update":
                    self.monitor_frame.current_article_text.set(msg["title"])
        except queue.Empty:
            pass
        if self.download_thread and self.download_thread.is_alive():
            self.root.after(100, self.process_queue)
        else:
            sys.stdout = sys.__stdout__
            self.download_thread = None

if __name__ == "__main__":
    root = tk.Tk()
    app = SciHubDownloaderGUI(root)
    root.mainloop()
