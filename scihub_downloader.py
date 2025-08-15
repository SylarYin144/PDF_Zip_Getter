import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import requests
import zipfile
import os
import re
import time
import sys
import threading
import queue
from datetime import datetime, timedelta
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
import itertools
import base64

# --- Helper Functions ---
STANDARD_USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.127 Safari/537.36'

def clean_filename(title):
    return re.sub(r'[\\/*?:"<>|]', '_', title)

def get_pdf_content_via_js(driver, pdf_url):
    script = "const callback=arguments[arguments.length-1];fetch(arguments[0]).then(r=>{if(!r.ok)throw new Error('Network response was not ok: '+r.status+' '+r.statusText);return r.blob()}).then(b=>{if('application/pdf'!==b.type&&!arguments[0].toLowerCase().endsWith('.pdf')&&'application/octet-stream'!==b.type)throw new Error('Content-Type is not application/pdf or octet-stream: '+b.type);const r=new FileReader;r.onloadend=()=>{const t=';base64,',e=r.result.substring(r.result.indexOf(t)+t.length);callback(e)},r.onerror=t=>{callback({error:'FileReader error: '+t.toString()})},r.readAsDataURL(b)}).catch(t=>{callback({error:'JS Fetch error: '+t.toString()})});"
    try:
        driver.set_script_timeout(90)
        result = driver.execute_async_script(script, pdf_url)
        if isinstance(result, dict) and 'error' in result:
            print(f"JS Fetch Helper Error: {result['error']}")
            return None
        if result: return base64.b64decode(result)
        return None
    except Exception as e:
        print(f"JS Fetch Helper Exception: {e}")
        return None

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

def download_with_selenium_google_scholar(driver, doi, title):
    print(f"Buscando en Google Scholar (Selenium) por DOI: {doi}")
    scholar_url = f"https://scholar.google.com/scholar?hl=en&q={doi}"
    try:
        driver.get(scholar_url)
        # Simplified logic: find a direct PDF link
        pdf_links = driver.find_elements(By.XPATH, "//a[contains(@href, '.pdf')]")
        if not pdf_links:
            return None, "FALLO - No se encontró link .pdf en Google Scholar"

        pdf_url = pdf_links[0].get_attribute('href')
        print(f"Intentando descargar desde URL: {pdf_url}")
        pdf_content = get_pdf_content_via_js(driver, pdf_url)
        if pdf_content:
            return pdf_content, f"OBTENIDO (Google Scholar Selenium)"
        return None, "FALLO - No se pudo obtener contenido desde Google Scholar"
    except Exception as e:
        print(f"Error en Selenium/Google Scholar: {e}")
        return None, "FALLO - Excepción en Google Scholar"

def download_with_selenium_pmc(driver, doi, title):
    print(f"Buscando en PubMed Central (Selenium) por DOI: {doi}")
    search_url = f"https://www.ncbi.nlm.nih.gov/pmc/?term={doi}"
    try:
        driver.get(search_url)
        # Find link to article page
        article_link = driver.find_element(By.CSS_SELECTOR, "div.rprt .title a")
        article_url = article_link.get_attribute('href')
        driver.get(article_url)
        # Find PDF link on article page
        pdf_link = driver.find_element(By.XPATH, "//a[contains(@href, '.pdf') and contains(., 'PDF')]")
        pdf_url = pdf_link.get_attribute('href')
        print(f"Intentando descargar desde URL: {pdf_url}")
        pdf_content = get_pdf_content_via_js(driver, pdf_url)
        if pdf_content:
            return pdf_content, "OBTENIDO (PMC Selenium)"
        return None, "FALLO - No se pudo obtener contenido desde PMC"
    except Exception as e:
        print(f"Error en Selenium/PMC: {e}")
        return None, "FALLO - Excepción en PMC"

# --- Main Application ---

class TextRedirector:
    def __init__(self, q): self.queue = q
    def write(self, str_): self.queue.put({'type': 'log', 'data': str_})
    def flush(self): pass

class SciHubDownloaderApp:
    DEFAULT_SCI_HUB_MIRRORS = ["https://sci-hub.se/", "https://sci-hub.st/", "https://sci-hub.ru/", "https://sci-hub.red/"]
    DEFAULT_INTER_DOI_DELAY = 5
    DEFAULT_MIRROR_SWITCH_DELAY = 3

    def __init__(self, root):
        self.root = root
        self.root.title("Sci-Hub Downloader Pro")
        self.root.geometry("900x750")
        self.root.minsize(700, 600)
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)

        self.ui_queue = queue.Queue()
        self.download_thread = None
        self.df_articles = None

        # UI Variables
        self.input_file_path = tk.StringVar()
        self.article_count_str = tk.StringVar(value="Detectados: 0 artículos")
        self.zip_output_path = tk.StringVar()
        self.report_output_path = tk.StringVar()
        self.use_scihub = tk.BooleanVar(value=True)
        self.use_google_scholar = tk.BooleanVar(value=True)
        self.use_pmc = tk.BooleanVar(value=False)
        self.inter_doi_delay = tk.IntVar(value=self.DEFAULT_INTER_DOI_DELAY)
        self.mirror_switch_delay = tk.IntVar(value=self.DEFAULT_MIRROR_SWITCH_DELAY)

        # Progress UI Variables
        self.progress_article_str = tk.StringVar(value="Artículo: 0/0 (0.0%)")
        self.progress_title_str = tk.StringVar(value="Título: -")
        self.progress_author_str = tk.StringVar(value="Autor: -")
        self.progress_journal_str = tk.StringVar(value="Revista: -")
        self.progress_year_str = tk.StringVar(value="Año: -")
        self.progress_doi_str = tk.StringVar(value="DOI: -")
        self.time_elapsed_str = tk.StringVar(value="Transcurrido: 0s")
        self.time_avg_doi_str = tk.StringVar(value="Promedio/DOI: 0s")
        self.time_remaining_str = tk.StringVar(value="Tiempo Restante: -")

        self._create_frames()
        self._create_config_widgets()
        self._create_progress_widgets()

        self.config_frame.pack(fill=tk.BOTH, expand=True)

    def _create_frames(self):
        self.config_frame = ttk.Frame(self.root, padding="10")
        self.progress_frame = ttk.Frame(self.root, padding="10")

    def _create_config_widgets(self):
        self.config_frame.columnconfigure(0, weight=1)
        files_frame = ttk.LabelFrame(self.config_frame, text="Archivos de Entrada y Salida", padding="10")
        files_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        files_frame.columnconfigure(1, weight=1)
        ttk.Label(files_frame, text="Seleccionar Archivo (.xlsx, .csv):").grid(row=0, column=0, sticky="w", pady=2)
        ttk.Entry(files_frame, textvariable=self.input_file_path, state="readonly").grid(row=0, column=1, sticky="ew", padx=5)
        ttk.Button(files_frame, text="Seleccionar...", command=self._select_input_file).grid(row=0, column=2, sticky="e")
        ttk.Label(files_frame, textvariable=self.article_count_str).grid(row=1, column=1, sticky="w", padx=5)
        ttk.Label(files_frame, text="Definir Ubicación del ZIP:").grid(row=2, column=0, sticky="w", pady=2)
        ttk.Entry(files_frame, textvariable=self.zip_output_path, state="readonly").grid(row=2, column=1, sticky="ew", padx=5)
        ttk.Button(files_frame, text="Guardar en...", command=self._select_zip_output).grid(row=2, column=2, sticky="e")
        sources_frame = ttk.LabelFrame(self.config_frame, text="Fuentes de Descarga y Reporte", padding="10")
        sources_frame.grid(row=1, column=0, sticky="ew", padx=5, pady=5)
        sources_frame.columnconfigure(1, weight=1)
        ttk.Checkbutton(sources_frame, text="Usar Sci-Hub", variable=self.use_scihub).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(sources_frame, text="Usar Google Scholar (respaldo)", variable=self.use_google_scholar).grid(row=1, column=0, sticky="w")
        ttk.Checkbutton(sources_frame, text="Usar PubMed Central (PMC) (respaldo)", variable=self.use_pmc).grid(row=2, column=0, sticky="w")
        ttk.Label(sources_frame, text="Definir Ruta de Reporte (.xlsx):").grid(row=3, column=0, sticky="w", pady=(10, 2))
        ttk.Entry(sources_frame, textvariable=self.report_output_path, state="readonly").grid(row=3, column=1, sticky="ew", padx=5)
        ttk.Button(sources_frame, text="Guardar como...", command=self._select_report_output).grid(row=3, column=2, sticky="e")
        advanced_frame = ttk.LabelFrame(self.config_frame, text="Configuración Avanzada", padding="10")
        advanced_frame.grid(row=2, column=0, sticky="ew", padx=5, pady=5)
        advanced_frame.columnconfigure(0, weight=1)
        ttk.Label(advanced_frame, text="Mirrors de Sci-Hub (uno por línea):").grid(row=0, column=0, sticky="w")
        self.mirrors_text = scrolledtext.ScrolledText(advanced_frame, height=5, wrap=tk.WORD)
        self.mirrors_text.grid(row=1, column=0, sticky="ew", pady=5)
        self.mirrors_text.insert(tk.END, "\n".join(self.DEFAULT_SCI_HUB_MIRRORS))
        timeouts_frame = ttk.Frame(advanced_frame)
        timeouts_frame.grid(row=2, column=0, sticky="ew", pady=5)
        ttk.Label(timeouts_frame, text="Espera entre búsquedas (s):").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Spinbox(timeouts_frame, from_=0, to=60, textvariable=self.inter_doi_delay, width=5).pack(side=tk.LEFT, padx=5)
        ttk.Label(timeouts_frame, text="Espera al cambiar de mirror (s):").pack(side=tk.LEFT, padx=(15, 5))
        ttk.Spinbox(timeouts_frame, from_=0, to=60, textvariable=self.mirror_switch_delay, width=5).pack(side=tk.LEFT, padx=5)
        style = ttk.Style()
        style.configure("Big.TButton", font=("Helvetica", 12, "bold"))
        self.start_button = ttk.Button(self.config_frame, text="Iniciar Descarga", style="Big.TButton", command=self._start_download_process)
        self.start_button.grid(row=3, column=0, pady=20, ipady=5)

    def _create_progress_widgets(self):
        self.progress_frame.columnconfigure(0, weight=1)
        self.progress_frame.rowconfigure(1, weight=1)
        top_info_frame = ttk.Frame(self.progress_frame)
        top_info_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        top_info_frame.columnconfigure(1, weight=1)
        current_article_frame = ttk.LabelFrame(top_info_frame, text="Procesando Artículo", padding="10")
        current_article_frame.grid(row=0, column=0, sticky="ns", padx=(0, 10))
        ttk.Label(current_article_frame, textvariable=self.progress_article_str, font=("Helvetica", 10, "bold")).pack(anchor="w")
        ttk.Label(current_article_frame, textvariable=self.progress_title_str).pack(anchor="w")
        ttk.Label(current_article_frame, textvariable=self.progress_author_str).pack(anchor="w")
        ttk.Label(current_article_frame, textvariable=self.progress_journal_str).pack(anchor="w")
        ttk.Label(current_article_frame, textvariable=self.progress_year_str).pack(anchor="w")
        ttk.Label(current_article_frame, textvariable=self.progress_doi_str).pack(anchor="w")
        metrics_frame = ttk.Frame(top_info_frame)
        metrics_frame.grid(row=0, column=1, sticky="nsew")
        metrics_frame.columnconfigure(0, weight=1)
        self.progress_bar = ttk.Progressbar(metrics_frame, orient="horizontal", mode="determinate")
        self.progress_bar.grid(row=0, column=0, columnspan=3, sticky="ew", pady=5)
        ttk.Label(metrics_frame, textvariable=self.time_elapsed_str).grid(row=1, column=0, sticky="w")
        ttk.Label(metrics_frame, textvariable=self.time_avg_doi_str).grid(row=1, column=1, sticky="w")
        ttk.Label(metrics_frame, textvariable=self.time_remaining_str).grid(row=1, column=2, sticky="w")
        list_frame = ttk.LabelFrame(self.progress_frame, text="Lista de Artículos", padding="10")
        list_frame.grid(row=1, column=0, sticky="nsew", pady=5)
        list_frame.rowconfigure(0, weight=1)
        list_frame.columnconfigure(0, weight=1)
        cols = ("#", "DOI", "Título", "Estado")
        self.article_treeview = ttk.Treeview(list_frame, columns=cols, show="headings", height=10)
        for col in cols: self.article_treeview.heading(col, text=col)
        self.article_treeview.column("#", width=50, anchor="center")
        self.article_treeview.column("DOI", width=200)
        self.article_treeview.column("Título", width=400)
        self.article_treeview.column("Estado", width=100, anchor="center")
        self.article_treeview.tag_configure('fallido', background='#FFDDDD')
        self.article_treeview.tag_configure('obtenido', background='#DDFFDD')
        vsb = ttk.Scrollbar(list_frame, orient="vertical", command=self.article_treeview.yview)
        self.article_treeview.configure(yscrollcommand=vsb.set)
        self.article_treeview.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        log_frame = ttk.LabelFrame(self.progress_frame, text="Log de Actividad", padding="10")
        log_frame.grid(row=2, column=0, sticky="ew", pady=5)
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)
        self.log_text_widget = scrolledtext.ScrolledText(log_frame, height=8, wrap=tk.WORD, state='disabled')
        self.log_text_widget.grid(row=0, column=0, sticky="nsew")
        self.back_button = ttk.Button(self.progress_frame, text="< Volver a Configuración", command=self._show_config_view)
        self.back_button.grid(row=3, column=0, sticky="w", pady=10)

    def _select_input_file(self):
        path = filedialog.askopenfilename(filetypes=(("Excel/CSV", "*.xlsx;*.xls;*.csv"), ("Todos", "*.*")))
        if path:
            self.input_file_path.set(path)
            try:
                self.df_articles = pd.read_excel(path) if path.lower().endswith(('.xlsx', '.xls')) else pd.read_csv(path)
                self.article_count_str.set(f"Detectados: {len(self.df_articles)} artículos")
            except Exception as e:
                messagebox.showerror("Error de Lectura", f"No se pudo leer el archivo:\n{e}")
                self.article_count_str.set("Error al leer el archivo")
                self.df_articles = None

    def _select_zip_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".zip", filetypes=(("Archivos ZIP", "*.zip"),))
        if path: self.zip_output_path.set(path)

    def _select_report_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=(("Archivos Excel", "*.xlsx"),), initialfile="SciHub_Reporte.xlsx")
        if path: self.report_output_path.set(path)

    def _start_download_process(self):
        config = {
            "input_file": self.input_file_path.get(),
            "zip_output": self.zip_output_path.get(),
            "report_output": self.report_output_path.get(),
            "use_scihub": self.use_scihub.get(),
            "use_google_scholar": self.use_google_scholar.get(),
            "use_pmc": self.use_pmc.get(),
            "mirrors": [m.strip() for m in self.mirrors_text.get("1.0", tk.END).strip().split("\n") if m.strip()],
            "inter_doi_delay": self.inter_doi_delay.get(),
            "mirror_switch_delay": self.mirror_switch_delay.get(),
        }
        if not config["input_file"] or not config["zip_output"]:
            messagebox.showwarning("Faltan Datos", "Por favor, especifique el archivo de entrada y la ubicación del ZIP.")
            return
        if self.df_articles is None:
            messagebox.showerror("Error de Archivo", "No se ha cargado correctamente la lista de artículos.")
            return

        self.config_frame.pack_forget()
        self.progress_frame.pack(fill=tk.BOTH, expand=True)
        self._populate_initial_treeview()

        self.download_thread = threading.Thread(target=self._download_worker, args=(config, self.ui_queue))
        self.download_thread.start()
        self.root.after(100, self._process_queue)

    def _populate_initial_treeview(self):
        for i in self.article_treeview.get_children():
            self.article_treeview.delete(i)
        for index, row in self.df_articles.iterrows():
            doi = str(row.get('DOI', row.get('doi', 'N/A'))).strip()
            title = str(row.get('Title', row.get('title', 'Sin Título'))).strip()
            self.article_treeview.insert("", "end", iid=index, values=(index + 1, doi, title, "Pendiente"))

    def _process_queue(self):
        try:
            while True:
                msg = self.ui_queue.get_nowait()
                msg_type, data = msg.get('type'), msg.get('data')
                if msg_type == 'log':
                    self.log_text_widget.configure(state='normal')
                    self.log_text_widget.insert(tk.END, data)
                    self.log_text_widget.see(tk.END)
                    self.log_text_widget.configure(state='disabled')
                elif msg_type == 'progress':
                    self.progress_bar['value'] = data['percentage']
                    self.progress_article_str.set(f"Artículo: {data['current']}/{data['total']} ({data['percentage']:.1f}%)")
                    self.time_elapsed_str.set(f"Transcurrido: {timedelta(seconds=int(data['elapsed']))}")
                    self.time_avg_doi_str.set(f"Promedio/DOI: {data['avg_time']:.1f}s")
                    self.time_remaining_str.set(f"Tiempo Restante: {timedelta(seconds=int(data['remaining']))}")
                elif msg_type == 'current_article':
                    self.progress_title_str.set(f"Título: {data.get('title', 'N/A')[:80]}")
                    self.progress_author_str.set(f"Autor: {data.get('author', 'N/A')}")
                    self.progress_journal_str.set(f"Revista: {data.get('journal', 'N/A')}")
                    self.progress_year_str.set(f"Año: {data.get('year', 'N/A')}")
                    self.progress_doi_str.set(f"DOI: {data.get('doi', 'N/A')}")
                elif msg_type == 'article_status':
                    self.article_treeview.item(data['row_id'], values=(data['row_id'] + 1, data['doi'], data['title'], data['status']), tags=(data['tag'],))
                elif msg_type == 'finished':
                    messagebox.showinfo("Proceso Finalizado", data)
                    self._show_config_view()
                    return
        except queue.Empty:
            if self.download_thread and self.download_thread.is_alive():
                self.root.after(100, self._process_queue)

    def _download_worker(self, config, q):
        sys.stdout = TextRedirector(q)
        start_time_total = time.time()
        driver = None
        all_articles_log, successful_articles_data, failed_articles_data = [], [], []
        try:
            if config['use_google_scholar'] or config['use_pmc']:
                try:
                    print("Inicializando WebDriver de Selenium...")
                    options = webdriver.ChromeOptions(); options.add_argument('--headless'); options.add_argument('--disable-gpu'); options.add_argument('--no-sandbox'); options.add_argument('--disable-dev-shm-usage')
                    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
                    print("WebDriver inicializado.")
                except Exception as e:
                    print(f"ALERTA: No se pudo inicializar WebDriver: {e}\nLas descargas con Selenium se omitirán.")
                    driver = None

            session = requests.Session(); session.headers.update({'User-Agent': STANDARD_USER_AGENT})
            total_articles = len(self.df_articles)
            successful_downloads = 0

            with zipfile.ZipFile(config['zip_output'], 'w', zipfile.ZIP_DEFLATED) as zf:
                for index, row in self.df_articles.iterrows():
                    current_num = index + 1
                    elapsed = time.time() - start_time_total
                    avg_time = elapsed / current_num if current_num > 0 else 0
                    q.put({'type': 'progress', 'data': {'current': current_num, 'total': total_articles, 'percentage': (current_num / total_articles) * 100, 'elapsed': elapsed, 'avg_time': avg_time, 'remaining': (total_articles - current_num) * avg_time}})
                    
                    original_row_data = row.to_dict()
                    doi = str(original_row_data.get('DOI', '')).strip()
                    title = str(original_row_data.get('Title', doi or 'N/A')).strip()
                    
                    q.put({'type': 'current_article', 'data': {'title': title, 'doi': doi, 'author': original_row_data.get('First Author', 'N/A'), 'journal': original_row_data.get('Journal/Book', 'N/A'),'year': original_row_data.get('Publication Year', 'N/A')}})
                    
                    pdf_content, status_detail = None, "DOI no provisto"
                    if doi:
                        status_detail = "Falló en todas las fuentes"
                        if config['use_scihub']:
                            for mirror in config['mirrors']:
                                print(f"Intentando mirror Sci-Hub: {mirror}")
                                pdf_content = extract_pdf_link_from_html(f"{mirror}{doi}", session)
                                if pdf_content: status_detail = f"OBTENIDO ({mirror})"; break
                                time.sleep(config['mirror_switch_delay'])
                        if not pdf_content and config['use_google_scholar'] and driver:
                            pdf_content, status_detail = download_with_selenium_google_scholar(driver, doi, title)
                        if not pdf_content and config['use_pmc'] and driver:
                            pdf_content, status_detail = download_with_selenium_pmc(driver, doi, title)
                    
                    status, tag = ("Obtenido", "obtenido") if pdf_content else ("Fallido", "fallido")
                    if pdf_content:
                        successful_downloads += 1
                        zf.writestr(f"{clean_filename(title)}.pdf", pdf_content)
                    
                    q.put({'type': 'article_status', 'data': {'row_id': index, 'doi': doi, 'title': title, 'status': status, 'tag': tag}})
                    time.sleep(config['inter_doi_delay'])

            summary_msg = f"Proceso completado.\n\nDescargas exitosas: {successful_downloads}/{total_articles}"
            q.put({'type': 'finished', 'data': summary_msg})
        except Exception as e:
            print(f"\nERROR CRÍTICO: {e}")
            import traceback; traceback.print_exc()
            q.put({'type': 'finished', 'data': f"El proceso ha fallado con un error crítico:\n{e}"})
        finally:
            if driver: print("Cerrando WebDriver..."); driver.quit()
            sys.stdout = sys.__stdout__

    def _show_config_view(self):
        self.progress_frame.pack_forget()
        self.config_frame.pack(fill=tk.BOTH, expand=True)

    def _on_closing(self):
        if self.download_thread and self.download_thread.is_alive():
            if messagebox.askokcancel("Salir", "El proceso de descarga está en curso. ¿Realmente desea salir?"):
                self.root.destroy()
        else:
            self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = SciHubDownloaderApp(root)
    root.mainloop()
