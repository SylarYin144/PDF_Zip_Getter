import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import pandas as pd
import requests
import zipfile
import os
import re
import time
import sys
import threading
import queue
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
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# --- Helper Functions (Directly from user, adapted for GUI) ---
STANDARD_USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.127 Safari/537.36'

def clean_filename(title):
    return re.sub(r'[\\/*?:"<>|]', '_', title)

def get_pdf_content_via_js(driver, pdf_url, queue):
    script = """
    const callback = arguments[arguments.length - 1];
    fetch(arguments[0])
        .then(response => response.blob())
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
            queue.put({"type": "log", "text": f"      - JS Fetch Helper Error: {result['error']}"})
            return None
        if result: return base64.b64decode(result)
        return None
    except Exception as e:
        queue.put({"type": "log", "text": f"      - JS Fetch Helper Exception: {e}"})
        return None

def download_from_scihub(doi, session, queue, mirrors, timeout, switch_delay):
    for mirror_url in mirrors:
        try:
            full_url = f"{mirror_url}{doi}"
            queue.put({"type": "log", "text": f"   - Intentando con mirror Sci-Hub: {mirror_url}"})

            html_response = session.get(full_url, timeout=timeout)
            html_response.raise_for_status()
            soup = BeautifulSoup(html_response.content, 'html.parser')
            iframe = soup.find('iframe', id='pdf')
            pdf_link = None
            if iframe and iframe.get('src'):
                pdf_src = iframe['src']
                if pdf_src.startswith('//'): pdf_link = 'https:' + pdf_src
                else: pdf_link = urljoin(full_url, pdf_src)

            if pdf_link:
                pdf_response = session.get(pdf_link, timeout=timeout)
                pdf_response.raise_for_status()
                if 'application/pdf' in pdf_response.headers.get('Content-Type', '').lower():
                    return pdf_response.content, f"Sci-Hub ({mirror_url})"

            if 'application/pdf' in html_response.headers.get('Content-Type', '').lower():
                return html_response.content, f"Sci-Hub ({mirror_url})"

            queue.put({"type": "log", "text": f"      - No se encontró PDF en {mirror_url}."})
            time.sleep(switch_delay)
        except requests.exceptions.RequestException as e:
            queue.put({"type": "log", "text": f"      - Error de red con mirror {mirror_url}: {e}."})
            time.sleep(switch_delay)
            continue
    return None, "FALLO - Sci-Hub"

def download_with_selenium_google_scholar(driver, doi, title, queue, timeout):
    queue.put({"type": "log", "text": f"   - Buscando en Google Scholar (Selenium) por: {doi or title}"})
    scholar_url = f"https://scholar.google.com/scholar?hl=en&q={doi or title}"
    try:
        driver.set_page_load_timeout(timeout)
        driver.get(scholar_url)
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, "gs_res_ccl_mid")))
        pdf_links_elements = driver.find_elements(By.XPATH, "//a[contains(@href, '.pdf')] | //div[@class='gs_ggs']//a")
        queue.put({"type": "log", "text": f"      - Se encontraron {len(pdf_links_elements)} enlaces potenciales."})
        for link_el in pdf_links_elements:
            pdf_url = link_el.get_attribute('href')
            if not pdf_url or 'javascript' in pdf_url: continue
            queue.put({"type": "log", "text": f"      - Intentando descargar desde: {pdf_url[:70]}..."})
            try:
                pdf_content = get_pdf_content_via_js(driver, pdf_url, queue)
                if pdf_content and len(pdf_content) > 1024:
                    return pdf_content, "Google Scholar (Selenium)"
            except Exception as e:
                queue.put({"type": "log", "text": f"         - Falló la descarga desde {pdf_url[:70]}: {e}"})
                continue
        return None, "FALLO - Google Scholar (Selenium)"
    except Exception as e:
        queue.put({"type": "log", "text": f"      - Error durante la búsqueda en Google Scholar (Selenium): {e}"})
        return None, "FALLO - Google Scholar (Selenium)"

def download_from_pmc(doi, session, queue, timeout):
    queue.put({"type": "log", "text": f"   - Buscando en PMC por DOI: {doi}"})
    try:
        id_conv_url = f"https://www.ncbi.nlm.nih.gov/pmc/utils/idconv/v1.0/?ids={doi}&format=json"
        response_id_conv = session.get(id_conv_url, timeout=timeout)
        response_id_conv.raise_for_status()
        data = response_id_conv.json()
        pmcid = data.get("records", [{}])[0].get("pmcid")
        if not pmcid: return None, "FALLO - No se encontró PMCID"

        efetch_url = f"https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?db=pmc&id={pmcid}&rettype=xml"
        response_efetch = session.get(efetch_url, timeout=timeout)
        response_efetch.raise_for_status()
        root = ET.fromstring(response_efetch.content)
        pdf_filename = None
        for element in root.iterfind(".//*[@content-type='pdf']"):
            pdf_filename = element.get('{http://www.w3.org/1999/xlink}href')
            if pdf_filename: break
        if not pdf_filename:
            for element in root.iterfind(".//*[@pub-id-type='pmc-pdf']"):
                pdf_filename = element.text
                if pdf_filename: break

        if not pdf_filename: return None, f"FALLO - No se encontró nombre de PDF en XML para {pmcid}"

        pdf_url = f"https://www.ncbi.nlm.nih.gov/pmc/articles/{pmcid}/pdf/{pdf_filename}"
        queue.put({"type": "log", "text": f"      - Enlace PDF encontrado en PMC: {pdf_url}"})
        pdf_response = session.get(pdf_url, timeout=timeout)
        pdf_response.raise_for_status()
        if 'application/pdf' in pdf_response.headers.get('Content-Type', ''):
            return pdf_response.content, f"PMC ({pmcid})"

        return None, "FALLO - El enlace de PMC no era un PDF"
    except Exception as e:
        queue.put({"type": "log", "text": f"      - Error durante la búsqueda en PMC: {e}"})
        return None, "FALLO - PMC"

def download_with_selenium_pmc(driver, doi, title, queue, timeout):
    queue.put({"type": "log", "text": f"   - Buscando en PMC (Selenium) por: {doi or title}"})
    search_url = f"https://www.ncbi.nlm.nih.gov/pmc/?term={doi or title}"
    try:
        driver.set_page_load_timeout(timeout)
        driver.get(search_url)
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CLASS_NAME, "rprt")))

        possible_article_links = driver.find_elements(By.CSS_SELECTOR, "div.rprt .title a")
        if not possible_article_links: return None, "FALLO - No se encontró artículo en búsqueda PMC"

        article_url = possible_article_links[0].get_attribute('href')
        queue.put({"type": "log", "text": f"      - Navegando a página de artículo: {article_url[:70]}..."})
        driver.get(article_url)

        pdf_link_elements = WebDriverWait(driver, 15).until(EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@class, 'format-pdf')] | //a[contains(text(), 'Download PDF')]")))
        if not pdf_link_elements: return None, "FALLO - No se encontró enlace PDF en página de artículo PMC"

        pdf_url = pdf_link_elements[0].get_attribute('href')
        queue.put({"type": "log", "text": f"      - Intentando descargar desde: {pdf_url[:70]}..."})
        pdf_content = get_pdf_content_via_js(driver, pdf_url, queue)
        if pdf_content and len(pdf_content) > 1024:
            return pdf_content, "PMC (Selenium)"

        return None, "FALLO - PMC (Selenium)"
    except Exception as e:
        queue.put({"type": "log", "text": f"      - Error durante la búsqueda en PMC (Selenium): {e}"})
        return None, "FALLO - PMC (Selenium)"

class SciHubDownloaderApp:
    def __init__(self, master):
        self.master = master
        self.progress_window = None
        self.queue = queue.Queue()
        self.download_thread = None
        self.cancel_event = threading.Event()
        self.pause_event = threading.Event()
        self.stats_counters = {"total": 0, "procesados": 0, "obtenidos": 0, "fallidos": 0, "fuentes": {}}
        master.title("Configuración de Descarga")
        master.geometry("600x500")

        self.main_frame = ttk.Frame(master, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(self.main_frame, text="Archivo de entrada (Excel):").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.input_file_var = tk.StringVar()
        ttk.Entry(self.main_frame, textvariable=self.input_file_var, width=50).grid(row=0, column=1, sticky=tk.EW)
        ttk.Button(self.main_frame, text="Examinar...", command=self.browse_input_file).grid(row=0, column=2, padx=5)

        ttk.Label(self.main_frame, text="Guardar ZIP en:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.zip_output_var = tk.StringVar()
        ttk.Entry(self.main_frame, textvariable=self.zip_output_var, width=50).grid(row=1, column=1, sticky=tk.EW)
        ttk.Button(self.main_frame, text="Examinar...", command=self.browse_zip_folder).grid(row=1, column=2, padx=5)

        ttk.Label(self.main_frame, text="Guardar Reporte en:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.report_output_var = tk.StringVar()
        ttk.Entry(self.main_frame, textvariable=self.report_output_var, width=50).grid(row=2, column=1, sticky=tk.EW)
        ttk.Button(self.main_frame, text="Examinar...", command=self.browse_report_folder).grid(row=2, column=2, padx=5)

        sources_frame = ttk.LabelFrame(self.main_frame, text="Fuentes de descarga", padding="10")
        sources_frame.grid(row=3, column=0, columnspan=3, sticky=tk.EW, pady=10)
        self.sources_vars = { "sci-hub": tk.BooleanVar(value=True), "pubmed": tk.BooleanVar(value=True), "google-scholar": tk.BooleanVar(value=True) }
        ttk.Checkbutton(sources_frame, text="Sci-Hub", variable=self.sources_vars["sci-hub"]).pack(side=tk.LEFT, padx=5)
        ttk.Checkbutton(sources_frame, text="PubMed Central", variable=self.sources_vars["pubmed"]).pack(side=tk.LEFT, padx=5)
        ttk.Checkbutton(sources_frame, text="Google Scholar (Selenium)", variable=self.sources_vars["google-scholar"]).pack(side=tk.LEFT, padx=5)

        timeouts_frame = ttk.LabelFrame(self.main_frame, text="Configuración de Timeouts (segundos)", padding="10")
        timeouts_frame.grid(row=4, column=0, columnspan=3, sticky=tk.EW, pady=10)
        ttk.Label(timeouts_frame, text="General:").grid(row=0, column=0, padx=5)
        self.general_timeout_var = tk.StringVar(value="45")
        ttk.Entry(timeouts_frame, textvariable=self.general_timeout_var, width=5).grid(row=0, column=1)
        ttk.Label(timeouts_frame, text="Cambio Mirror:").grid(row=0, column=2, padx=5)
        self.switch_timeout_var = tk.StringVar(value="5")
        ttk.Entry(timeouts_frame, textvariable=self.switch_timeout_var, width=5).grid(row=0, column=3)

        self.start_button = ttk.Button(self.main_frame, text="Iniciar Descarga", command=self.start_download)
        self.start_button.grid(row=5, column=0, columnspan=3, pady=20)
        self.main_frame.columnconfigure(1, weight=1)

    def browse_input_file(self): self.input_file_var.set(filedialog.askopenfilename(filetypes=(("Excel", "*.xlsx *.xls"),("All files", "*.*"))))
    def browse_zip_folder(self): self.zip_output_var.set(filedialog.askdirectory())
    def browse_report_folder(self): self.report_output_var.set(filedialog.askdirectory())

    def start_download(self):
        if not all([self.input_file_var.get(), self.zip_output_var.get(), self.report_output_var.get()]):
            messagebox.showerror("Error", "Por favor, complete todos los campos."); return

        self.cancel_event.clear(); self.pause_event.clear()
        self.stats_counters = {"total": 0, "procesados": 0, "obtenidos": 0, "fallidos": 0, "fuentes": {}}
        self.start_button.config(state=tk.DISABLED); self.master.withdraw(); self._create_progress_ui()

        try: timeouts = { "general": int(self.general_timeout_var.get()), "switch": int(self.switch_timeout_var.get()) }
        except ValueError: messagebox.showerror("Error", "Los timeouts deben ser números."); self.close_progress_window(); return

        self.download_thread = threading.Thread(
            target=self._execute_download,
            args=(self.input_file_var.get(), self.zip_output_var.get(), self.report_output_var.get(), dict(self.sources_vars), timeouts),
            daemon=True
        )
        self.download_thread.start()
        self.master.after(100, self._update_progress_ui)

    def _create_progress_ui(self):
        self.progress_window = tk.Toplevel(self.master)
        self.progress_window.title("Progreso de Descarga")
        self.progress_window.geometry("900x700")
        self.progress_window.protocol("WM_DELETE_WINDOW", self.cancel_download)

        main_frame = ttk.Frame(self.progress_window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.columnconfigure(1, weight=1)

        left_panel = ttk.Frame(main_frame, padding="10")
        left_panel.grid(row=0, column=0, sticky="ns")

        ttk.Label(left_panel, text="Progreso General:").pack(anchor=tk.W)
        self.progress_bar = ttk.Progressbar(left_panel, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill=tk.X, pady=5, anchor=tk.W)
        self.progress_label = ttk.Label(left_panel, text="Iniciando...")
        self.progress_label.pack(anchor=tk.W, pady=5)
        self.stats_label = ttk.Label(left_panel, text="")
        self.stats_label.pack(anchor=tk.W, pady=5)

        ttk.Label(left_panel, text="Registro:").pack(anchor=tk.W, pady=(10,0))
        log_frame = ttk.Frame(left_panel)
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.log_text = tk.Text(log_frame, height=15, width=60, state=tk.DISABLED, wrap=tk.WORD)
        log_scroll = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text['yscrollcommand'] = log_scroll.set
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        control_frame = ttk.Frame(left_panel)
        control_frame.pack(pady=10)
        self.pause_resume_button = ttk.Button(control_frame, text="Pausar", command=self.toggle_pause_resume)
        self.pause_resume_button.pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="Cancelar", command=self.cancel_download).pack(side=tk.LEFT, padx=5)

        right_panel = ttk.Frame(main_frame)
        right_panel.grid(row=0, column=1, sticky="nsew", padx=10)
        self.fig = Figure(figsize=(5, 6), dpi=100); self.fig.subplots_adjust(hspace=0.5)
        self.ax1 = self.fig.add_subplot(3, 1, 1); self.ax2 = self.fig.add_subplot(3, 1, 2); self.ax3 = self.fig.add_subplot(3, 1, 3)
        self.chart_canvas = FigureCanvasTkAgg(self.fig, master=right_panel)
        self.chart_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        self._update_charts()

    def _update_charts(self):
        stats = self.stats_counters; self.ax1.clear(); self.ax2.clear(); self.ax3.clear()
        self.fig.suptitle("Estadísticas de Descarga", fontsize=14)
        self.ax1.set_title('Estado General'); self.ax2.set_title('Obtenidos vs. Fallidos'); self.ax3.set_title('Fuentes de Éxito')

        # Chart 1
        sizes1 = [stats['obtenidos'], stats['fallidos'], stats["total"] - stats["procesados"]]
        if sum(sizes1) > 0:
            self.ax1.pie(sizes1, labels=['Obtenidos', 'Fallidos', 'Pendientes'], autopct='%1.1f%%', startangle=90, colors=['#4CAF50', '#F44336', '#9E9E9E'])
        else:
            self.ax1.text(0.5, 0.5, 'Esperando...', ha='center', va='center')

        # Chart 2
        sizes2 = [stats['obtenidos'], stats['fallidos']]
        if sum(sizes2) > 0:
            self.ax2.pie(sizes2, labels=['Obtenidos', 'Fallidos'], autopct='%1.1f%%', startangle=90, colors=['#4CAF50', '#F44336'])
        else:
            self.ax2.text(0.5, 0.5, 'Esperando...', ha='center', va='center')

        # Chart 3
        if stats['fuentes']:
            self.ax3.pie(list(stats['fuentes'].values()), labels=list(stats['fuentes'].keys()), autopct='%1.1f%%', startangle=90)
        else:
            self.ax3.text(0.5, 0.5, 'Esperando...', ha='center', va='center')

        self.chart_canvas.draw()

    def _update_progress_ui(self):
        try:
            while True:
                msg = self.queue.get_nowait()
                if msg["type"] == "progress": self.progress_bar["value"] = msg["value"]; self.progress_label.config(text=f"Procesando {msg['current']} de {msg['total']}...")
                elif msg["type"] == "stats":
                    self.stats_counters = msg["stats"]
                    self.stats_label.config(text=f"Buscados: {self.stats_counters['procesados']}/{self.stats_counters['total']} | Obtenidos: {self.stats_counters['obtenidos']} | Fallidos: {self.stats_counters['fallidos']}")
                    self._update_charts()
                elif msg["type"] == "log": self.log_text.config(state=tk.NORMAL); self.log_text.insert(tk.END, msg["text"] + "\n"); self.log_text.see(tk.END); self.log_text.config(state=tk.DISABLED)
                elif msg["type"] == "done": messagebox.showinfo("Finalizado", "El proceso ha terminado."); self.close_progress_window(); return
                elif msg["type"] == "error": messagebox.showerror("Error", msg["text"]); self.close_progress_window(); return
                elif msg["type"] == "cancelled": messagebox.showinfo("Cancelado", "La descarga ha sido cancelada."); self.close_progress_window(); return
        except queue.Empty: pass
        if self.download_thread.is_alive(): self.master.after(100, self._update_progress_ui)

    def _execute_download(self, input_file, zip_dir, report_dir, sources_vars, timeouts):
        driver = None; session = requests.Session(); session.headers.update({'User-Agent': STANDARD_USER_AGENT})
        obtenidos_list, fallidos_list = [], []
        zip_filename = os.path.join(zip_dir, f"Descarga_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip")

        try:
            with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zf:
                selected_sources = {k: v.get() for k, v in sources_vars.items()}
                self.queue.put({"type": "log", "text": f"Fuentes activas: {[k for k,v in selected_sources.items() if v]}"})

                if selected_sources["google-scholar"]:
                    try:
                        self.queue.put({"type": "log", "text": "Inicializando WebDriver..."})
                        options = webdriver.ChromeOptions(); options.add_argument('--headless'); options.add_argument("--disable-gpu")
                        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
                    except Exception as e: self.queue.put({"type": "error", "text": f"Error WebDriver: {e}"}); return

                df = pd.read_excel(input_file).fillna('')
                self.stats_counters["total"] = len(df); self.queue.put({"type": "stats", "stats": self.stats_counters})
                self.queue.put({"type": "log", "text": f"Se encontraron {len(df)} artículos."})

                for index, row in df.iterrows():
                    if self.cancel_event.is_set(): break
                    while self.pause_event.is_set(): time.sleep(0.1)
                    self.stats_counters["procesados"] += 1
                    doi = str(row.get('DOI', '')).strip(); title = str(row.get('Title', '')).strip()
                    self.queue.put({"type": "log", "text": f"--- Procesando {index+1}/{len(df)}: {title or doi} ---"})
                    pdf_content, source_used = self._attempt_download(doi, title, selected_sources, session, driver, timeouts)
                    if pdf_content:
                        self.stats_counters["obtenidos"] += 1; self.stats_counters["fuentes"][source_used] = self.stats_counters["fuentes"].get(source_used, 0) + 1
                        zf.writestr(clean_filename(title or doi) + '.pdf', pdf_content)
                        new_row = row.to_dict(); new_row.update({'Estado': "Obtenido", 'Fuente': source_used}); obtenidos_list.append(new_row)
                    else: fallidos_list.append(row.to_dict())
                    self.queue.put({"type": "progress", "value": (index + 1) * 100 / len(df), "current": index + 1, "total": len(df)})
                    self.queue.put({"type": "stats", "stats": self.stats_counters})

                if fallidos_list and not self.cancel_event.is_set():
                    self.queue.put({"type": "log", "text": "\n--- Iniciando Fase de Reintento ---"})
                    failed_to_retry = list(fallidos_list); fallidos_list.clear()
                    for item in failed_to_retry:
                        if self.cancel_event.is_set(): break
                        doi = str(item.get('DOI', '')).strip(); title = str(item.get('Title', '')).strip()
                        self.queue.put({"type": "log", "text": f"--- Reintentando: {title or doi} ---"})
                        pdf_content, source_used = self._attempt_download(doi, title, selected_sources, session, driver, timeouts)
                        if pdf_content:
                            self.stats_counters["obtenidos"] += 1; self.stats_counters["fuentes"][source_used] = self.stats_counters["fuentes"].get(source_used, 0) + 1
                            zf.writestr(clean_filename(title or doi) + '.pdf', pdf_content)
                            new_row = item; new_row.update({'Estado': "Obtenido (Reintento)", 'Fuente': source_used}); obtenidos_list.append(new_row)
                        else: fallidos_list.append(item)
                        self.queue.put({"type": "stats", "stats": self.stats_counters})

                if not self.cancel_event.is_set():
                    self.stats_counters["fallidos"] = len(fallidos_list); self.queue.put({"type": "stats", "stats": self.stats_counters})
                    self.queue.put({"type": "log", "text": "\nGenerando reporte Excel..."})
                    report_filename = os.path.join(report_dir, f"Reporte_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
                    with pd.ExcelWriter(report_filename, engine='openpyxl') as writer:
                        pd.DataFrame(fallidos_list).to_excel(writer, sheet_name='Fallidos', index=False)
                        pd.DataFrame(obtenidos_list).to_excel(writer, sheet_name='Obtenidos', index=False)
                    self.queue.put({"type": "log", "text": f"Reporte guardado."}); self.queue.put({"type": "done"})
                else: self.queue.put({"type": "cancelled"})
        except Exception as e:
            import traceback; self.queue.put({"type": "error", "text": f"Error fatal: {e}\n{traceback.format_exc()}"})
        finally:
            if driver: driver.quit()
            session.close()
            self.queue.put({"type": "log", "text": "Recursos liberados."})

    def _attempt_download(self, doi, title, selected_sources, session, driver, timeouts):
        pdf_content, source_used = None, "Ninguna"
        general_timeout = timeouts["general"]

        if selected_sources.get("sci-hub") and doi:
            mirrors = ["https://sci-hub.se/", "https://sci-hub.st/", "https://sci-hub.ru/"]
            pdf_content, source_used = download_from_scihub(doi, session, self.queue, mirrors, general_timeout, timeouts["switch"])

        if not pdf_content and selected_sources.get("pubmed") and doi:
            pdf_content, source_used = download_from_pmc(doi, session, self.queue, general_timeout)

        if not pdf_content and selected_sources.get("google-scholar") and driver:
            pdf_content, source_used = download_with_selenium_google_scholar(driver, doi, title, self.queue, general_timeout)
            if not pdf_content and selected_sources.get("pubmed") and driver:
                 pdf_content, source_used = download_with_selenium_pmc(driver, doi, title, self.queue, general_timeout)

        return pdf_content, source_used

    def toggle_pause_resume(self):
        if self.pause_event.is_set(): self.pause_event.clear(); self.pause_resume_button.config(text="Pausar"); self.queue.put({"type": "log", "text": "--- Reanudado ---"})
        else: self.pause_event.set(); self.pause_resume_button.config(text="Reanudar"); self.queue.put({"type": "log", "text": "--- Pausado ---"})

    def cancel_download(self):
        if messagebox.askyesno("Confirmar", "¿Cancelar la descarga?"):
            self.cancel_event.set()
            if self.pause_event.is_set(): self.pause_event.clear()
            self.pause_resume_button.config(state=tk.DISABLED)

    def close_progress_window(self):
        if self.progress_window: self.progress_window.destroy(); self.progress_window = None
        self.master.deiconify(); self.start_button.config(state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk()
    app = SciHubDownloaderApp(root)
    root.mainloop()
