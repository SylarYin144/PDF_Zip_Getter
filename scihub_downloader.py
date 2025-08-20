import tkinter as tk
from tkinter import ttk, filedialog, messagebox
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
import threading
import queue
import itertools
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# --- Configuration Constants ---
DEFAULT_SCI_HUB_MIRRORS_EXAMPLE = ["https://sci-hub.se/", "https://sci-hub.st/", "https://sci-hub.box/", "https://sci-hub.ru/", "https://sci-hub.red/"]
INTER_DOI_DELAY_SECONDS = 5
MIRROR_SWITCH_DELAY_SECONDS = 3
STANDARD_USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.127 Safari/537.36'

class SciHubDownloaderApp:
    def __init__(self):
        # --- GUI Elements (initialized in create_gui) ---
        self.root = None
        self.progress_window = None
        self.progress_bar = None
        self.stats_label = None
        self.log_text = None
        self.charts = {}
        self.pause_resume_button = None

        # --- Tkinter Variables (initialized in create_gui) ---
        self.tk_vars = {}

        # --- Default Configuration (standard Python types) ---
        self.config = {
            'input_file_path': "",
            'zip_path': "",
            'excel_report_path': "",
            'use_scihub': True,
            'use_pmc': True,
            'use_google_scholar': True,
            'inter_doi_delay': INTER_DOI_DELAY_SECONDS,
            'mirror_switch_delay': MIRROR_SWITCH_DELAY_SECONDS,
            'scihub_mirrors': ", ".join(DEFAULT_SCI_HUB_MIRRORS_EXAMPLE)
        }

        # --- Threading and Queue ---
        self.download_thread = None
        self.message_queue = queue.Queue()
        self.is_paused = threading.Event()
        self.is_cancelled = threading.Event()
        self.driver = None

    def get_config(self, key):
        """Helper to get config value from the correct source (tk_vars or self.config)."""
        if self.root: # GUI is active
            return self.tk_vars[key].get()
        else: # Headless/test mode
            return self.config[key]

    def create_gui(self):
        """Creates the entire GUI. Must be called to run in GUI mode."""
        self.root = tk.Tk()
        self.root.title("Sci-Hub Downloader")

        # --- Initialize Tkinter variables ---
        self.tk_vars['input_file_path'] = tk.StringVar(value=self.config['input_file_path'])
        self.tk_vars['zip_path'] = tk.StringVar(value=self.config['zip_path'])
        self.tk_vars['excel_report_path'] = tk.StringVar(value=self.config['excel_report_path'])
        self.tk_vars['use_scihub'] = tk.BooleanVar(value=self.config['use_scihub'])
        self.tk_vars['use_pmc'] = tk.BooleanVar(value=self.config['use_pmc'])
        self.tk_vars['use_google_scholar'] = tk.BooleanVar(value=self.config['use_google_scholar'])
        self.tk_vars['inter_doi_delay'] = tk.IntVar(value=self.config['inter_doi_delay'])
        self.tk_vars['mirror_switch_delay'] = tk.IntVar(value=self.config['mirror_switch_delay'])
        self.tk_vars['scihub_mirrors'] = tk.StringVar(value=self.config['scihub_mirrors'])

        self.create_config_window()
        return self.root
        
    def create_config_window(self):
        """Creates the initial configuration window."""
        config_frame = ttk.Frame(self.root, padding="20")
        config_frame.pack(fill=tk.BOTH, expand=True)
        config_frame.columnconfigure(1, weight=1)

        paths_frame = ttk.LabelFrame(config_frame, text="Rutas de Archivos", padding="10")
        paths_frame.grid(row=0, column=0, columnspan=3, sticky=tk.EW, pady=5)
        paths_frame.columnconfigure(1, weight=1)
        ttk.Label(paths_frame, text="Archivo de Entrada (Excel/CSV):").grid(row=0, column=0, sticky=tk.W, pady=5, padx=5)
        ttk.Entry(paths_frame, textvariable=self.tk_vars['input_file_path'], width=60).grid(row=0, column=1, sticky=tk.EW, padx=5)
        ttk.Button(paths_frame, text="Examinar...", command=self.browse_input_file).grid(row=0, column=2, padx=5)
        ttk.Label(paths_frame, text="Guardar ZIP en:").grid(row=1, column=0, sticky=tk.W, pady=5, padx=5)
        ttk.Entry(paths_frame, textvariable=self.tk_vars['zip_path'], width=60).grid(row=1, column=1, sticky=tk.EW, padx=5)
        ttk.Button(paths_frame, text="Examinar...", command=self.browse_zip_path).grid(row=1, column=2, padx=5)
        ttk.Label(paths_frame, text="Guardar Reporte en (Opcional):").grid(row=2, column=0, sticky=tk.W, pady=5, padx=5)
        ttk.Entry(paths_frame, textvariable=self.tk_vars['excel_report_path'], width=60).grid(row=2, column=1, sticky=tk.EW, padx=5)
        ttk.Button(paths_frame, text="Examinar...", command=self.browse_report_path).grid(row=2, column=2, padx=5)

        sources_frame = ttk.LabelFrame(config_frame, text="Fuentes de Descarga", padding="10")
        sources_frame.grid(row=1, column=0, columnspan=3, sticky=tk.EW, pady=10)
        ttk.Checkbutton(sources_frame, text="Sci-Hub", variable=self.tk_vars['use_scihub']).pack(side=tk.LEFT, padx=15, pady=5)
        ttk.Checkbutton(sources_frame, text="PubMed Central (PMC)", variable=self.tk_vars['use_pmc']).pack(side=tk.LEFT, padx=15, pady=5)
        ttk.Checkbutton(sources_frame, text="Google Scholar", variable=self.tk_vars['use_google_scholar']).pack(side=tk.LEFT, padx=15, pady=5)

        delays_frame = ttk.LabelFrame(config_frame, text="Retrasos (segundos)", padding="10")
        delays_frame.grid(row=2, column=0, columnspan=3, sticky=tk.EW, pady=5)
        ttk.Label(delays_frame, text="Entre cada DOI:").pack(side=tk.LEFT, padx=5)
        ttk.Entry(delays_frame, textvariable=self.tk_vars['inter_doi_delay'], width=5).pack(side=tk.LEFT, padx=5)
        ttk.Label(delays_frame, text="Al cambiar de Mirror:").pack(side=tk.LEFT, padx=15)
        ttk.Entry(delays_frame, textvariable=self.tk_vars['mirror_switch_delay'], width=5).pack(side=tk.LEFT, padx=5)

        mirrors_frame = ttk.LabelFrame(config_frame, text="Mirrors de Sci-Hub (separados por coma)", padding="10")
        mirrors_frame.grid(row=3, column=0, columnspan=3, sticky=tk.EW, pady=5)
        mirrors_frame.columnconfigure(0, weight=1)
        ttk.Entry(mirrors_frame, textvariable=self.tk_vars['scihub_mirrors']).grid(row=0, column=0, sticky=tk.EW)

        ttk.Button(config_frame, text="Iniciar Descarga", command=self.start_download_process).grid(row=4, column=0, columnspan=3, pady=20)

    def browse_input_file(self):
        path = filedialog.askopenfilename(title="Seleccionar archivo con DOIs", filetypes=(("Archivos Excel", "*.xlsx *.xls"), ("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*")))
        if path: self.tk_vars['input_file_path'].set(path)

    def browse_zip_path(self):
        path = filedialog.asksaveasfilename(title="Guardar archivo ZIP como...", defaultextension=".zip", filetypes=(("Archivos ZIP", "*.zip"),))
        if path: self.tk_vars['zip_path'].set(path)

    def browse_report_path(self):
        path = filedialog.asksaveasfilename(title="Guardar Reporte Excel como...", defaultextension=".xlsx", filetypes=(("Archivos Excel", "*.xlsx"),))
        if path: self.tk_vars['excel_report_path'].set(path)

    def start_download_process(self):
        if not self.get_config('input_file_path') or not self.get_config('zip_path'):
            if self.root: messagebox.showerror("Error", "Debe especificar el archivo de entrada y la ruta del ZIP de salida.")
            else: print("ERROR: Input file or zip path not set.")
            return
        if not (self.get_config('use_scihub') or self.get_config('use_pmc') or self.get_config('use_google_scholar')):
            if self.root: messagebox.showerror("Error", "Debe seleccionar al menos una fuente de descarga.")
            else: print("ERROR: No download source selected.")
            return

        if self.root: self.create_progress_window()
        self.is_cancelled.clear()
        self.is_paused.clear()

        self.download_thread = threading.Thread(target=self.download_worker, daemon=True)
        self.download_thread.start()

        if self.root: self.root.after(100, self.process_queue)

    def create_progress_window(self):
        self.progress_window = tk.Toplevel(self.root)
        self.progress_window.title("Progreso de la Descarga")
        self.progress_window.geometry("900x700")
        self.progress_window.protocol("WM_DELETE_WINDOW", self.cancel_download)

        main_frame = ttk.Frame(self.progress_window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill=tk.X, pady=5)
        self.stats_label = ttk.Label(top_frame, text="Buscados: 0/0 | Obtenidos: 0 | Fallidos: 0", font=("Helvetica", 10))
        self.stats_label.pack(side=tk.LEFT, padx=5)
        self.progress_bar = ttk.Progressbar(top_frame, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill=tk.X, expand=True, padx=5)

        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        bottom_frame.columnconfigure(0, weight=2); bottom_frame.columnconfigure(1, weight=1); bottom_frame.rowconfigure(0, weight=1)
        log_frame = ttk.LabelFrame(bottom_frame, text="Registro de Actividad", padding="5")
        log_frame.grid(row=0, column=0, sticky="nsew", padx=5)
        log_frame.rowconfigure(0, weight=1); log_frame.columnconfigure(0, weight=1)
        self.log_text = tk.Text(log_frame, wrap=tk.WORD, state='disabled', font=("Courier New", 9))
        log_scroll = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text['yscrollcommand'] = log_scroll.set
        self.log_text.grid(row=0, column=0, sticky="nsew")
        log_scroll.grid(row=0, column=1, sticky="ns")
        charts_frame = ttk.Frame(bottom_frame)
        charts_frame.grid(row=0, column=1, sticky="nsew", padx=5)
        charts_frame.rowconfigure(0, weight=1); charts_frame.rowconfigure(1, weight=1); charts_frame.rowconfigure(2, weight=1); charts_frame.columnconfigure(0, weight=1)
        self.create_charts(charts_frame)

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10, side=tk.BOTTOM)
        self.pause_resume_button = ttk.Button(button_frame, text="Pausar", command=self.toggle_pause)
        self.pause_resume_button.pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancelar", command=self.cancel_download).pack(side=tk.LEFT, padx=5)

    def create_charts(self, parent_frame):
        self.charts['status'] = self.create_pie_chart(parent_frame, "Estado General", 0)
        self.charts['comparison'] = self.create_pie_chart(parent_frame, "Obtenidos vs. Fallidos", 1)
        self.charts['source'] = self.create_pie_chart(parent_frame, "Fuentes de Éxito", 2)

    def create_pie_chart(self, parent, title, row):
        fig = Figure(figsize=(3, 2.3), dpi=80)
        ax = fig.add_subplot(111)
        ax.set_title(title, fontsize=10)
        fig.tight_layout(pad=1.5)
        canvas = FigureCanvasTkAgg(fig, master=parent)
        canvas.get_tk_widget().grid(row=row, column=0, sticky="nsew")
        canvas.draw()
        return {'fig': fig, 'ax': ax, 'canvas': canvas}

    def update_charts(self, stats):
        if not self.root or not self.progress_window or not self.progress_window.winfo_exists(): return
        status_ax = self.charts['status']['ax']
        status_ax.clear()
        status_labels = 'Obtenidos', 'Fallidos', 'Pendientes'
        status_sizes = [stats['obtained'], stats['failed'], stats['pending']]
        status_ax.pie(status_sizes, labels=status_labels, autopct='%1.1f%%', startangle=90, textprops={'fontsize': 8})
        status_ax.axis('equal'); status_ax.set_title("Estado General", fontsize=10)
        self.charts['status']['canvas'].draw()
        comp_ax = self.charts['comparison']['ax']
        comp_ax.clear()
        if (stats['obtained'] + stats['failed']) > 0:
            comp_ax.pie([stats['obtained'], stats['failed']], labels=['Obtenidos', 'Fallidos'], autopct='%1.1f%%', colors=['#4CAF50', '#F44336'], textprops={'fontsize': 8})
        comp_ax.set_title("Obtenidos vs. Fallidos", fontsize=10)
        self.charts['comparison']['canvas'].draw()
        source_ax = self.charts['source']['ax']
        source_ax.clear()
        if stats['sources']:
            source_ax.pie(list(stats['sources'].values()), labels=list(stats['sources'].keys()), autopct='%1.1f%%', textprops={'fontsize': 8})
        source_ax.set_title("Fuentes de Éxito", fontsize=10)
        self.charts['source']['canvas'].draw()

    def process_queue(self):
        try:
            while True:
                message = self.message_queue.get_nowait()
                if not self.progress_window or not self.progress_window.winfo_exists(): continue
                msg_type, data = message.get('type'), message.get('data')
                if msg_type == 'log':
                    self.log_text.configure(state='normal')
                    self.log_text.insert(tk.END, data + '\n'); self.log_text.see(tk.END)
                    self.log_text.configure(state='disabled')
                elif msg_type == 'progress': self.progress_bar['value'] = data
                elif msg_type == 'stats':
                    self.stats_label['text'] = f"Buscados: {data['current']}/{data['total']} | Obtenidos: {data['obtained']} | Fallidos: {data['failed']}"
                    self.update_charts(data)
                elif msg_type == 'finish':
                    messagebox.showinfo("Proceso Terminado", data)
                    self.progress_window.destroy()
                elif msg_type == 'error':
                    messagebox.showerror("Error Crítico", data)
                    if self.progress_window: self.progress_window.destroy()
        except queue.Empty: pass
        finally:
            if self.download_thread and (self.download_thread.is_alive() or not self.message_queue.empty()):
                self.root.after(100, self.process_queue)

    def toggle_pause(self):
        if self.is_paused.is_set():
            self.is_paused.clear()
            self.pause_resume_button['text'] = "Pausar"
            self.log_message("...Reanudando descarga.")
        else:
            self.is_paused.set()
            self.pause_resume_button['text'] = "Reanudar"
            self.log_message("Descarga pausada.")

    def cancel_download(self):
        do_cancel = True
        if self.root: # Only show messagebox if in GUI mode
            do_cancel = messagebox.askyesno("Cancelar", "¿Está seguro?")

        if do_cancel:
            self.is_cancelled.set()
            self.is_paused.clear()
            self.log_message("Cancelando descarga...")

    def log_message(self, msg):
        self.message_queue.put({'type': 'log', 'data': str(msg)})

    def check_for_pause_or_cancel(self):
        if self.is_cancelled.is_set(): return True
        while self.is_paused.is_set():
            if self.is_cancelled.is_set(): return True
            time.sleep(0.5)
        return False

    def get_pdf_content_via_js(self, driver, pdf_url):
        script = """
        const callback = arguments[arguments.length - 1];
        fetch(arguments[0])
            .then(response => response.ok ? response.blob() : Promise.reject(new Error(response.statusText)))
            .then(blob => {
                const reader = new FileReader();
                reader.onloadend = () => {
                    const base64Marker = ';base64,';
                    callback(reader.result.substring(reader.result.indexOf(base64Marker) + base64Marker.length));
                };
                reader.onerror = (err) => callback({error: 'FileReader error: ' + err.toString()});
                reader.readAsDataURL(blob);
            })
            .catch(error => callback({error: 'JS Fetch error: ' + error.toString()}));
        """
        try:
            driver.set_script_timeout(90)
            result = driver.execute_async_script(script, pdf_url)
            if isinstance(result, dict) and 'error' in result:
                self.log_message(f"JS Helper Error for {pdf_url}: {result['error']}")
                return None
            return base64.b64decode(result) if result else None
        except Exception as e:
            self.log_message(f"JS Helper Exception for {pdf_url}: {e}")
            return None

    def clean_filename(self, title):
        return re.sub(r'[\\/*?:"<>|]', '_', title)

    def extract_pdf_link_from_html(self, article_page_url, session):
        try:
            response = session.get(article_page_url, timeout=30)
            response.raise_for_status()
            soup = BeautifulSoup(response.content, 'html.parser')
            iframe = soup.find('iframe', id='pdf')
            if iframe and iframe.get('src'):
                pdf_src = iframe['src']
                return urljoin(article_page_url, pdf_src) if pdf_src.startswith('/') else ('https:' + pdf_src if pdf_src.startswith('//') else pdf_src)
            embed = soup.find('embed', attrs={'type': 'application/pdf'})
            if embed and embed.get('src'): return urljoin(article_page_url, embed['src'])
            return None
        except requests.exceptions.RequestException as e:
            self.log_message(f"Error fetching HTML for PDF link extraction from {article_page_url}: {e}")
            return None

    def download_from_scihub(self, doi, session, mirrors):
        for mirror_base_url in mirrors:
            if self.check_for_pause_or_cancel(): return None, "CANCELADO"
            full_url = f"{mirror_base_url}{doi}"
            self.log_message(f"Intentando con Sci-Hub mirror: {mirror_base_url}")
            try:
                pdf_download_url = self.extract_pdf_link_from_html(full_url, session)
                if pdf_download_url:
                    response = session.get(pdf_download_url, timeout=60)
                    response.raise_for_status()
                    if 'application/pdf' in response.headers.get('Content-Type', '').lower():
                        return response.content, f"OBTENIDO (Sci-Hub Iframe/Embed - {mirror_base_url})"
                response = session.get(full_url, timeout=30)
                response.raise_for_status()
                if 'application/pdf' in response.headers.get('Content-Type', '').lower():
                    return response.content, f"OBTENIDO (Sci-Hub Direct - {mirror_base_url})"
            except requests.exceptions.RequestException as e:
                self.log_message(f"Fallo en mirror {mirror_base_url}: {e}")
                if self.check_for_pause_or_cancel(): return None, "CANCELADO"
                time.sleep(self.get_config('mirror_switch_delay'))
        return None, "FALLO - No se pudo descargar desde Sci-Hub"

    def download_from_pmc(self, doi, session):
        self.log_message(f"Intentando con PubMed Central para DOI: {doi}")
        if not self.driver: return None, "FALLO - Driver no inicializado"
        return self.download_with_selenium_pmc(self.driver, doi, "")

    def download_from_google_scholar(self, doi, session):
        self.log_message(f"Intentando con Google Scholar para DOI: {doi}")
        if not self.driver: return None, "FALLO - Driver no inicializado"
        return self.download_with_selenium_google_scholar(self.driver, doi, "")

    def download_with_selenium_pmc(self, driver, doi, title):
        search_url = f"https://www.ncbi.nlm.nih.gov/pmc/?term={doi}"
        try:
            driver.get(search_url)
            article_link = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.rprt .title a")))
            article_url = article_link.get_attribute('href')
            driver.get(article_url)
            pdf_link = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "a.format-pdf")))
            pdf_url = pdf_link.get_attribute('href')
            pdf_content = self.get_pdf_content_via_js(driver, pdf_url)
            return (pdf_content, f"OBTENIDO (PMC Selenium - {pdf_url})") if pdf_content else (None, "FALLO - No se pudo obtener contenido PDF desde PMC")
        except (TimeoutException, NoSuchElementException) as e:
            self.log_message(f"Error en Selenium (PMC) para DOI {doi}: {e}")
            return None, "FALLO - No se encontró el PDF en PMC (Selenium)"

    def download_with_selenium_google_scholar(self, driver, doi, title):
        scholar_url = f"https://scholar.google.com/scholar?hl=en&q={doi}"
        try:
            driver.get(scholar_url)
            pdf_link_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//a[contains(@href, '.pdf')]")))
            pdf_url = pdf_link_element.get_attribute('href')
            pdf_content = self.get_pdf_content_via_js(driver, pdf_url)
            return (pdf_content, f"OBTENIDO (Google Scholar Selenium - {pdf_url})") if pdf_content else (None, "FALLO - No se pudo obtener contenido PDF desde Google Scholar")
        except (TimeoutException, NoSuchElementException) as e:
            self.log_message(f"Error en Selenium (Google Scholar) para DOI {doi}: {e}")
            return None, "FALLO - No se encontró el PDF en Google Scholar (Selenium)"

    def download_worker(self):
        try:
            self.log_message("Inicializando WebDriver...")
            try:
                options = webdriver.ChromeOptions(); options.add_argument('--headless'); options.add_argument('--disable-gpu'); options.add_argument('--no-sandbox'); options.add_argument('--disable-dev-shm-usage')
                self.driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
                self.log_message("WebDriver inicializado.")
            except Exception as e:
                self.log_message(f"ALERTA: No se pudo inicializar Selenium. Las descargas de PMC y Google Scholar no estarán disponibles. Error: {e}")
                self.driver = None

            input_path = self.get_config('input_file_path')
            self.log_message(f"Leyendo archivo: {input_path}")
            try:
                df = pd.read_csv(input_path) if input_path.endswith('.csv') else pd.read_excel(input_path)
            except Exception as e: raise Exception(f"No se pudo leer el archivo de entrada: {e}")

            total_articles = len(df)
            stats = {'total': total_articles, 'current': 0, 'obtained': 0, 'failed': 0, 'pending': total_articles, 'sources': {}}
            session = requests.Session(); session.headers.update({'User-Agent': STANDARD_USER_AGENT})
            user_mirrors = [m.strip() for m in self.get_config('scihub_mirrors').split(',') if m.strip()]

            with zipfile.ZipFile(self.get_config('zip_path'), 'w', zipfile.ZIP_DEFLATED) as zf:
                for index, row in df.iterrows():
                    if self.check_for_pause_or_cancel(): break
                    doi = str(row.get('DOI', '')).strip()
                    title = str(row.get('Title', '')).strip()
                    effective_title = title if title else doi
                    self.log_message(f"--- Procesando {index + 1}/{total_articles}: {effective_title} ---")

                    if not doi:
                        self.log_message("DOI vacío, saltando."); stats['failed'] += 1; stats['pending'] -= 1; continue
                    
                    pdf_content, status_msg = None, ""
                    if self.get_config('use_scihub'): pdf_content, status_msg = self.download_from_scihub(doi, session, user_mirrors)
                    if not pdf_content and self.get_config('use_pmc'):
                        if self.check_for_pause_or_cancel(): break
                        pdf_content, status_msg = self.download_from_pmc(doi, session)
                    if not pdf_content and self.get_config('use_google_scholar'):
                        if self.check_for_pause_or_cancel(): break
                        pdf_content, status_msg = self.download_from_google_scholar(doi, session)
                    
                    stats['current'] = index + 1; stats['pending'] -= 1
                    if pdf_content:
                        stats['obtained'] += 1
                        source_name = re.search(r"\((.*?)\)", status_msg).group(1) if re.search(r"\((.*?)\)", status_msg) else "Desconocido"
                        stats['sources'][source_name] = stats['sources'].get(source_name, 0) + 1
                        zf.writestr(self.clean_filename(effective_title)[:150] + ".pdf", pdf_content)
                        self.log_message(f"ÉXITO: {status_msg}")
                    else:
                        stats['failed'] += 1
                        self.log_message(f"FALLO: No se pudo descargar DOI: {doi}")

                    self.message_queue.put({'type': 'progress', 'data': (stats['current'] / total_articles) * 100})
                    self.message_queue.put({'type': 'stats', 'data': stats.copy()})
                    if self.check_for_pause_or_cancel(): break
                    time.sleep(self.get_config('inter_doi_delay'))

            finish_msg = "Cancelado." if self.is_cancelled.is_set() else f"Completado. Obtenidos: {stats['obtained']}, Fallidos: {stats['failed']}."
            self.message_queue.put({'type': 'finish', 'data': finish_msg})
        except Exception as e:
            import traceback
            self.log_message(f"Error fatal en worker: {e}\n{traceback.format_exc()}")
            self.message_queue.put({'type': 'error', 'data': f"Error crítico: {e}"})
        finally:
            if self.driver: self.driver.quit(); self.log_message("WebDriver cerrado.")

if __name__ == "__main__":
    app = SciHubDownloaderApp()
    root = app.create_gui()
    root.mainloop()
