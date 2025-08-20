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
from urllib.parse import urljoin, quote_plus
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import base64
import csv
import tempfile

# --- Helper Functions ---
STANDARD_USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.127 Safari/537.36'

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
            if pdf_src.startswith('/'): return urljoin(article_page_url, pdf_src)
            return pdf_src
        embed = soup.find('embed', attrs={'type': 'application/pdf'})
        if embed and embed.get('src'): return urljoin(article_page_url, embed['src'])
        return None
    except Exception as e:
        print(f"Error extrayendo link de HTML: {e}")
        return None

def download_from_google_scholar(query, session):
    print(f"Buscando en Google Scholar por: '{query}'")
    scholar_url = f"https://scholar.google.com/scholar?hl=en&q={quote_plus(query)}"
    try:
        headers = {'User-Agent': STANDARD_USER_AGENT, 'Accept-Language': 'en-US,en;q=0.9'}
        response = session.get(scholar_url, headers=headers, timeout=30)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, 'html.parser')

        potential_links = []
        for link_tag in soup.find_all('a', href=True):
            href = link_tag.get('href', '')
            if href.lower().endswith('.pdf') and 'scholar.google.com' not in href:
                potential_links.append(href)

        print(f"Encontrados {len(potential_links)} links PDF potenciales en Google Scholar.")
        for pdf_url in potential_links:
            try:
                print(f"Intentando descargar desde: {pdf_url}")
                pdf_response = session.get(pdf_url, headers=headers, timeout=60, allow_redirects=True)
                if 'application/pdf' in pdf_response.headers.get('Content-Type', '').lower() and len(pdf_response.content) > 1000:
                    print(f"Descarga exitosa desde: {pdf_url}")
                    return pdf_response.content, "Google Scholar"
            except Exception as e:
                print(f"Intento fallido para {pdf_url}: {e}")
        return None, "FALLO - No se encontró PDF válido en Google Scholar"
    except Exception as e:
        print(f"Error en Google Scholar: {e}")
        return None, "FALLO - Excepción en Google Scholar"

def download_from_pmc(query, session):
    print(f"Buscando en PubMed Central por: '{query}'")
    try:
        pmcid = None
        if re.match(r'10.\d{4,9}/[-._;()/:A-Z0-9]+$', query, re.IGNORECASE):
            id_conv_url = f"https://www.ncbi.nlm.nih.gov/pmc/utils/idconv/v1.0/?ids={query}&format=json"
            response_id = session.get(id_conv_url, timeout=20)
            pmcid = response_id.json().get("records", [{}])[0].get("pmcid")

        if not pmcid:
            search_url = f"https://www.ncbi.nlm.nih.gov/pmc/?term={quote_plus(query)}"
            response_search = session.get(search_url, timeout=30)
            soup_search = BeautifulSoup(response_search.content, 'html.parser')
            article_link_tag = soup_search.find('div', class_='rprt').find('a')
            if not article_link_tag: return None, "FALLO - No se encontró artículo en PMC"
            article_url = urljoin(search_url, article_link_tag['href'])
        else:
            article_url = f"https://www.ncbi.nlm.nih.gov/pmc/articles/{pmcid}/"

        response_html = session.get(article_url, timeout=30)
        soup = BeautifulSoup(response_html.content, 'html.parser')

        for link_tag in soup.select('div.format-menu a[href$=".pdf"], .pdf-btn a[href$=".pdf"]'):
            pdf_url = urljoin(article_url, link_tag['href'])
            print(f"Intentando descargar desde URL: {pdf_url}")
            pdf_response = session.get(pdf_url, timeout=60)
            if 'application/pdf' in pdf_response.headers.get('Content-Type', '') and len(pdf_response.content) > 1000:
                return pdf_response.content, "PMC"
        return None, "FALLO - No se encontró link PDF en PMC"
    except Exception as e:
        print(f"Error en PMC: {e}")
        return None, "FALLO - Excepción en PMC"

class TextRedirector:
    def __init__(self, q): self.queue = q
    def write(self, str_): self.queue.put({'type': 'log', 'data': str_})
    def flush(self): pass

class SciHubDownloaderApp:
    def __init__(self, root):
        self.root = root; self.root.title("Sci-Hub Downloader Pro"); self.root.geometry("1200x850"); self.root.minsize(800, 700)
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
        self.ui_queue = queue.Queue(); self.download_thread = None; self.df_articles = None
        self.pause_event = threading.Event(); self.cancel_event = threading.Event()
        self.input_file_path = tk.StringVar(); self.article_count_str = tk.StringVar(value="Detectados: 0 artículos"); self.zip_output_path = tk.StringVar(); self.report_output_path = tk.StringVar()
        self.use_scihub = tk.BooleanVar(value=True); self.use_google_scholar = tk.BooleanVar(value=True); self.use_pmc = tk.BooleanVar(value=True)
        self.inter_doi_delay = tk.IntVar(value=5); self.mirror_switch_delay = tk.IntVar(value=3)
        self.progress_article_str = tk.StringVar(value="Artículo: 0/0 (0.0%)"); self.progress_title_str = tk.StringVar(value="Título: -"); self.progress_author_str = tk.StringVar(value="Autor: -")
        self.progress_journal_str = tk.StringVar(value="Revista: -"); self.progress_year_str = tk.StringVar(value="Año: -"); self.progress_doi_str = tk.StringVar(value="DOI: -")
        self.stats_text_str = tk.StringVar(value="Buscados: 0/0 | Obtenidos: 0 | Fallidos: 0 | Pendientes: 0")
        self._create_frames(); self._create_config_widgets(); self._create_progress_widgets(); self.config_frame.pack(fill=tk.BOTH, expand=True)

    def _create_frames(self): self.config_frame = ttk.Frame(self.root, padding="10"); self.progress_frame = ttk.Frame(self.root, padding="10")
    def _create_config_widgets(self):
        self.config_frame.columnconfigure(0, weight=1)
        files_frame = ttk.LabelFrame(self.config_frame, text="Archivos de Entrada y Salida", padding="10"); files_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5); files_frame.columnconfigure(1, weight=1)
        ttk.Label(files_frame, text="Seleccionar Archivo (.xlsx, .csv):").grid(row=0, column=0, sticky="w", pady=2); ttk.Entry(files_frame, textvariable=self.input_file_path, state="readonly").grid(row=0, column=1, sticky="ew", padx=5); ttk.Button(files_frame, text="Seleccionar...", command=self._select_input_file).grid(row=0, column=2, sticky="e")
        ttk.Label(files_frame, textvariable=self.article_count_str).grid(row=1, column=1, sticky="w", padx=5); ttk.Label(files_frame, text="Definir Ubicación del ZIP:").grid(row=2, column=0, sticky="w", pady=2); ttk.Entry(files_frame, textvariable=self.zip_output_path, state="readonly").grid(row=2, column=1, sticky="ew", padx=5); ttk.Button(files_frame, text="Guardar en...", command=self._select_zip_output).grid(row=2, column=2, sticky="e")
        sources_frame = ttk.LabelFrame(self.config_frame, text="Fuentes de Descarga y Reporte", padding="10"); sources_frame.grid(row=1, column=0, sticky="ew", padx=5, pady=5); sources_frame.columnconfigure(1, weight=1)
        ttk.Checkbutton(sources_frame, text="Usar Sci-Hub", variable=self.use_scihub).grid(row=0, column=0, sticky="w"); ttk.Checkbutton(sources_frame, text="Usar PubMed Central (respaldo)", variable=self.use_pmc).grid(row=1, column=0, sticky="w"); ttk.Checkbutton(sources_frame, text="Usar Google Scholar (respaldo)", variable=self.use_google_scholar).grid(row=2, column=0, sticky="w")
        ttk.Label(sources_frame, text="Definir Ruta de Reporte (.xlsx):").grid(row=3, column=0, sticky="w", pady=(10, 2)); ttk.Entry(sources_frame, textvariable=self.report_output_path, state="readonly").grid(row=3, column=1, sticky="ew", padx=5); ttk.Button(sources_frame, text="Guardar como...", command=self._select_report_output).grid(row=3, column=2, sticky="e")
        advanced_frame = ttk.LabelFrame(self.config_frame, text="Configuración Avanzada", padding="10"); advanced_frame.grid(row=2, column=0, sticky="ew", padx=5, pady=5); advanced_frame.columnconfigure(0, weight=1)
        ttk.Label(advanced_frame, text="Mirrors de Sci-Hub (uno por línea):").grid(row=0, column=0, sticky="w"); self.mirrors_text = scrolledtext.ScrolledText(advanced_frame, height=5, wrap=tk.WORD); self.mirrors_text.grid(row=1, column=0, sticky="ew", pady=5); self.mirrors_text.insert(tk.END, "\n".join(["https://sci-hub.se/", "https://sci-hub.st/", "https://sci-hub.ru/", "https://sci-hub.red/"]))
        timeouts_frame = ttk.Frame(advanced_frame); timeouts_frame.grid(row=2, column=0, sticky="w", pady=5)
        ttk.Label(timeouts_frame, text="Espera entre búsquedas (s):").pack(side=tk.LEFT, padx=(0, 5)); ttk.Spinbox(timeouts_frame, from_=0, to=60, textvariable=self.inter_doi_delay, width=5).pack(side=tk.LEFT, padx=5); ttk.Label(timeouts_frame, text="Espera al cambiar de mirror (s):").pack(side=tk.LEFT, padx=(15, 5)); ttk.Spinbox(timeouts_frame, from_=0, to=60, textvariable=self.mirror_switch_delay, width=5).pack(side=tk.LEFT, padx=5)
        style = ttk.Style(); style.configure("Big.TButton", font=("Helvetica", 12, "bold")); self.start_button = ttk.Button(self.config_frame, text="Iniciar Descarga", style="Big.TButton", command=self._start_download_process); self.start_button.grid(row=3, column=0, pady=20, ipady=5)
    def _create_progress_widgets(self):
        self.progress_frame.columnconfigure(0, weight=1); self.progress_frame.columnconfigure(1, weight=1); self.progress_frame.rowconfigure(0, weight=1)
        left_column_frame = ttk.Frame(self.progress_frame); left_column_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10)); left_column_frame.columnconfigure(0, weight=1); left_column_frame.rowconfigure(1, weight=1)
        top_info_frame = ttk.Frame(left_column_frame); top_info_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10)); top_info_frame.columnconfigure(1, weight=1)
        current_article_frame = ttk.LabelFrame(top_info_frame, text="Procesando Artículo", padding="10"); current_article_frame.grid(row=0, column=0, sticky="ns", padx=(0, 10))
        ttk.Label(current_article_frame, textvariable=self.progress_article_str, font=("Helvetica", 10, "bold")).pack(anchor="w"); ttk.Label(current_article_frame, textvariable=self.progress_title_str).pack(anchor="w"); ttk.Label(current_article_frame, textvariable=self.progress_author_str).pack(anchor="w"); ttk.Label(current_article_frame, textvariable=self.progress_journal_str).pack(anchor="w"); ttk.Label(current_article_frame, textvariable=self.progress_year_str).pack(anchor="w"); ttk.Label(current_article_frame, textvariable=self.progress_doi_str).pack(anchor="w")
        metrics_frame = ttk.Frame(top_info_frame); metrics_frame.grid(row=0, column=1, sticky="nsew"); metrics_frame.columnconfigure(0, weight=1)
        self.progress_bar = ttk.Progressbar(metrics_frame, orient="horizontal", mode="determinate"); self.progress_bar.grid(row=0, column=0, columnspan=3, sticky="ew", pady=5)
        self.stats_label = ttk.Label(metrics_frame, textvariable=self.stats_text_str, justify=tk.CENTER); self.stats_label.grid(row=1, column=0, columnspan=3, sticky="ew", pady=5)
        list_frame = ttk.LabelFrame(left_column_frame, text="Lista de Artículos", padding="10"); list_frame.grid(row=1, column=0, sticky="nsew", pady=5); list_frame.rowconfigure(0, weight=1); list_frame.columnconfigure(0, weight=1)
        cols = ("#", "DOI", "Título", "Estado"); self.article_treeview = ttk.Treeview(list_frame, columns=cols, show="headings", height=10);
        for col in cols: self.article_treeview.heading(col, text=col)
        self.article_treeview.column("#", width=50, anchor="center"); self.article_treeview.column("DOI", width=200); self.article_treeview.column("Título", width=400); self.article_treeview.column("Estado", width=100, anchor="center")
        self.article_treeview.tag_configure('fallido', background='#FFDDDD'); self.article_treeview.tag_configure('obtenido', background='#DDFFDD')
        vsb = ttk.Scrollbar(list_frame, orient="vertical", command=self.article_treeview.yview); self.article_treeview.configure(yscrollcommand=vsb.set); self.article_treeview.grid(row=0, column=0, sticky="nsew"); vsb.grid(row=0, column=1, sticky="ns")
        log_frame = ttk.LabelFrame(left_column_frame, text="Log de Actividad", padding="10"); log_frame.grid(row=2, column=0, sticky="ew", pady=5); log_frame.rowconfigure(0, weight=1); log_frame.columnconfigure(0, weight=1)
        self.log_text_widget = scrolledtext.ScrolledText(log_frame, height=8, wrap=tk.WORD, state='disabled'); self.log_text_widget.grid(row=0, column=0, sticky="nsew")
        right_column_frame = ttk.Frame(self.progress_frame); right_column_frame.grid(row=0, column=1, sticky="nsew"); right_column_frame.columnconfigure(0, weight=1); right_column_frame.rowconfigure(0, weight=1)
        charts_frame = ttk.LabelFrame(right_column_frame, text="Estadísticas", padding="10"); charts_frame.pack(fill=tk.BOTH, expand=True); charts_frame.columnconfigure(0, weight=1)
        self.fig1 = Figure(figsize=(4, 2.5), dpi=100); self.ax1 = self.fig1.add_subplot(111); self.canvas1 = FigureCanvasTkAgg(self.fig1, master=charts_frame); self.canvas1.get_tk_widget().pack(pady=5, fill=tk.X)
        self.fig2 = Figure(figsize=(4, 2.5), dpi=100); self.ax2 = self.fig2.add_subplot(111); self.canvas2 = FigureCanvasTkAgg(self.fig2, master=charts_frame); self.canvas2.get_tk_widget().pack(pady=5, fill=tk.X)
        self.fig3 = Figure(figsize=(4, 2.5), dpi=100); self.ax3 = self.fig3.add_subplot(111); self.canvas3 = FigureCanvasTkAgg(self.fig3, master=charts_frame); self.canvas3.get_tk_widget().pack(pady=5, fill=tk.X)
        controls_frame = ttk.Frame(right_column_frame); controls_frame.pack(fill=tk.X, pady=10)
        self.pause_button = ttk.Button(controls_frame, text="Pausar", command=self._toggle_pause); self.pause_button.pack(side=tk.LEFT, expand=True, padx=5)
        self.cancel_button = ttk.Button(controls_frame, text="Cancelar", command=self._cancel_download); self.cancel_button.pack(side=tk.LEFT, expand=True, padx=5)
        self._update_pie_charts()
    def _select_input_file(self):
        path = filedialog.askopenfilename(filetypes=(("Excel/CSV", "*.xlsx;*.xls;*.csv"), ("Todos", "*.*")))
        if path: self.input_file_path.set(path);
        try: self.df_articles = pd.read_excel(path) if path.lower().endswith(('.xlsx', '.xls')) else pd.read_csv(path); self.article_count_str.set(f"Detectados: {len(self.df_articles)} artículos")
        except Exception as e: messagebox.showerror("Error de Lectura", f"No se pudo leer el archivo:\n{e}"); self.df_articles = None
    def _select_zip_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".zip", filetypes=(("Archivos ZIP", "*.zip"),))
        if path: self.zip_output_path.set(path)
    def _select_report_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=(("Archivos Excel", "*.xlsx"),), initialfile="SciHub_Reporte.xlsx")
        if path: self.report_output_path.set(path)
    def _start_download_process(self):
        config = {"input_file": self.input_file_path.get(), "zip_output": self.zip_output_path.get(), "report_output": self.report_output_path.get(), "use_scihub": self.use_scihub.get(), "use_google_scholar": self.use_google_scholar.get(), "use_pmc": self.use_pmc.get(), "mirrors": [m.strip() for m in self.mirrors_text.get("1.0", tk.END).strip().split("\n") if m.strip()], "inter_doi_delay": self.inter_doi_delay.get(), "mirror_switch_delay": self.mirror_switch_delay.get()}
        if not config["input_file"] or not config["zip_output"]: messagebox.showwarning("Faltan Datos", "Especifique el archivo de entrada y la ubicación del ZIP."); return
        if self.df_articles is None: messagebox.showerror("Error de Archivo", "No se ha cargado la lista de artículos."); return
        self.pause_event.clear(); self.cancel_event.clear(); self.pause_button.config(text="Pausar")
        self.config_frame.pack_forget(); self.progress_frame.pack(fill=tk.BOTH, expand=True); self._populate_initial_treeview()
        self.download_thread = threading.Thread(target=self._download_worker, args=(config, self.ui_queue)); self.download_thread.start()
        self.root.after(100, self._process_queue)
    def _populate_initial_treeview(self):
        for i in self.article_treeview.get_children(): self.article_treeview.delete(i)
        for index, row in self.df_articles.iterrows():
            doi = str(row.get('DOI', 'N/A')).strip(); title = str(row.get('Title', 'Sin Título')).strip()
            self.article_treeview.insert("", "end", iid=index, values=(index + 1, doi, title, "Pendiente"))
    def _process_queue(self):
        try:
            while True:
                msg = self.ui_queue.get_nowait(); msg_type, data = msg.get('type'), msg.get('data')
                if msg_type == 'log': self.log_text_widget.configure(state='normal'); self.log_text_widget.insert(tk.END, data); self.log_text_widget.see(tk.END); self.log_text_widget.configure(state='disabled')
                elif msg_type == 'progress': self.progress_bar['value'] = data['percentage']; self.progress_article_str.set(f"Artículo: {data['current']}/{data['total']} ({data['percentage']:.1f}%)"); self.stats_text_str.set(f"Buscados: {data['stats']['successful'] + data['stats']['failed']}/{data['total']} | Obtenidos: {data['stats']['successful']} | Fallidos: {data['stats']['failed']} | Pendientes: {data['stats']['pending']}")
                elif msg_type == 'current_article': self.progress_title_str.set(f"Título: {data.get('title', 'N/A')[:80]}"); self.progress_author_str.set(f"Autor: {data.get('author', 'N/A')}"); self.progress_journal_str.set(f"Revista: {data.get('journal', 'N/A')}"); self.progress_year_str.set(f"Año: {data.get('year', 'N/A')}"); self.progress_doi_str.set(f"DOI: {data.get('doi', 'N/A')}")
                elif msg_type == 'article_status': self.article_treeview.item(data['row_id'], values=(data['row_id'] + 1, data['doi'], data['title'], data['status']), tags=(data['tag'],)); self._update_pie_charts(data['stats'])
                elif msg_type == 'finished': messagebox.showinfo("Proceso Finalizado", data); self._show_config_view(); return
        except queue.Empty:
            if self.download_thread and self.download_thread.is_alive(): self.root.after(100, self._process_queue)
    def _update_pie_charts(self, stats=None):
        if stats is None: stats = {'successful': 0, 'failed': 0, 'pending': 1, 'sources': {}}
        self.ax1.clear(); labels1=['Obtenidos','Fallidos','Pendientes']; sizes1=[stats['successful'],stats['failed'],stats['pending']]; colors1=['#8FBC8F','#F08080','#D3D3D3']
        self.ax1.pie(sizes1, labels=labels1, colors=colors1, autopct='%1.1f%%', startangle=90, wedgeprops={'edgecolor':'white'}); self.ax1.set_title('Estado General'); self.ax1.axis('equal'); self.canvas1.draw()
        self.ax2.clear(); labels2=['Obtenidos','Fallidos']; sizes2=[stats['successful'],stats['failed']] if (stats['successful']+stats['failed'])>0 else [1,0]; colors2=['#8FBC8F','#F08080']
        self.ax2.pie(sizes2, labels=labels2, colors=colors2, autopct='%1.1f%%', startangle=90, wedgeprops={'edgecolor':'white'}); self.ax2.set_title('Resultados'); self.ax2.axis('equal'); self.canvas2.draw()
        self.ax3.clear(); source_labels=list(stats['sources'].keys()); source_sizes=list(stats['sources'].values())
        if not source_sizes: source_labels,source_sizes=['N/A'],[1]
        self.ax3.pie(source_sizes, labels=source_labels, autopct='%1.1f%%', startangle=90, wedgeprops={'edgecolor':'white'}); self.ax3.set_title('Fuentes Exitosas'); self.ax3.axis('equal'); self.canvas3.draw()
    def _toggle_pause(self):
        if self.pause_event.is_set(): self.pause_event.clear(); self.pause_button.config(text="Pausar")
        else: self.pause_event.set(); self.pause_button.config(text="Reanudar")
    def _cancel_download(self):
        if messagebox.askokcancel("Cancelar", "Esto detendrá el proceso de descarga actual. ¿Está seguro?"): self.cancel_event.set()
    def _download_worker(self, config, q):
        sys.stdout = TextRedirector(q); start_time_total = time.time(); stats = {'successful': 0, 'failed': 0, 'pending': len(self.df_articles), 'sources': {}}
        temp_dir = tempfile.gettempdir(); success_csv_path = os.path.join(temp_dir, f"scihub_success_{os.getpid()}.csv"); fail_csv_path = os.path.join(temp_dir, f"scihub_fail_{os.getpid()}.csv")
        f_success, f_fail = None, None
        try:
            f_success = open(success_csv_path, 'w', newline='', encoding='utf-8'); f_fail = open(fail_csv_path, 'w', newline='', encoding='utf-8')
            success_writer = csv.writer(f_success); fail_writer = csv.writer(f_fail)
            header = list(self.df_articles.columns); success_writer.writerow(header + ["source"]); fail_writer.writerow(header + ["reason"])
            session = requests.Session(); session.headers.update({'User-Agent': STANDARD_USER_AGENT}); total_articles = len(self.df_articles)
            with zipfile.ZipFile(config['zip_output'], 'w', zipfile.ZIP_DEFLATED) as zf:
                for index, row in self.df_articles.iterrows():
                    if self.cancel_event.is_set(): print("\nDescarga cancelada."); break
                    while self.pause_event.is_set(): time.sleep(1)
                    current_num = index + 1
                    q.put({'type': 'progress', 'data': {'current': current_num, 'total': total_articles, 'percentage': (current_num / total_articles) * 100, 'stats': stats}})
                    original_row_data = row.to_dict(); doi = str(original_row_data.get('DOI', '')).strip(); title = str(original_row_data.get('Title', doi or 'N/A')).strip()
                    q.put({'type': 'current_article', 'data': {'title': title, 'doi': doi, 'author': original_row_data.get('First Author', 'N/A'), 'journal': original_row_data.get('Journal/Book', 'N/A'),'year': original_row_data.get('Publication Year', 'N/A')}})
                    query_term = doi if (doi and doi.lower() != 'nan') else title
                    pdf_content, source, reason = None, "N/A", "Término de búsqueda no válido"
                    if query_term:
                        reason = "Falló en todas las fuentes"
                        if config['use_scihub']:
                            for mirror in config['mirrors']:
                                if self.cancel_event.is_set(): break
                                print(f"Intentando Sci-Hub ({query_term}) en: {mirror}"); pdf_content = extract_pdf_link_from_html(f"{mirror}{query_term}", session)
                                if pdf_content and len(pdf_content) > 1000: source = "Sci-Hub"; break
                                else: pdf_content = None
                                time.sleep(config['mirror_switch_delay'])
                        if not pdf_content and config['use_pmc']: pdf_content, source = download_from_pmc(query_term, session)
                        if not pdf_content and config['use_google_scholar']: pdf_content, source = download_from_google_scholar(query_term, session)
                    stats['pending'] -= 1; row_values = list(original_row_data.values())
                    if pdf_content:
                        stats['successful'] += 1; stats['sources'][source] = stats['sources'].get(source, 0) + 1
                        status, tag = ("Obtenido", "obtenido"); zf.writestr(f"{clean_filename(title)}.pdf", pdf_content); success_writer.writerow(row_values + [source])
                    else:
                        stats['failed'] += 1; status, tag = ("Fallido", "fallido"); fail_writer.writerow(row_values + [reason])
                    q.put({'type': 'article_status', 'data': {'row_id': index, 'doi': doi, 'title': title, 'status': status, 'tag': tag, 'stats': stats}})
                    time.sleep(config['inter_doi_delay'])
            if config.get('report_output'):
                print("Generando reporte Excel desde archivos CSV temporales...")
                try:
                    if f_success: f_success.close(); f_success = None
                    if f_fail: f_fail.close(); f_fail = None
                    df_obtenidos = pd.read_csv(success_csv_path); df_fallidos = pd.read_csv(fail_csv_path)
                    with pd.ExcelWriter(config['report_output'], engine='openpyxl') as writer:
                        df_fallidos.to_excel(writer, sheet_name='Fallidos', index=False)
                        df_obtenidos.to_excel(writer, sheet_name='Obtenidos', index=False)
                    print(f"Reporte Excel guardado en: {config['report_output']}")
                except Exception as e: print(f"Error al generar el reporte Excel: {e}")
            summary_msg = f"Proceso completado.\n\nDescargas exitosas: {stats['successful']}/{total_articles}"
            if self.cancel_event.is_set(): summary_msg = f"Proceso cancelado.\n\nDescargas completadas: {stats['successful']}/{total_articles}"
            q.put({'type': 'finished', 'data': summary_msg})
        except Exception as e:
            print(f"\nERROR CRÍTICO: {e}"); import traceback; traceback.print_exc()
            q.put({'type': 'finished', 'data': f"El proceso ha fallado con un error crítico:\n{e}"})
        finally:
            if f_success: f_success.close();
            if f_fail: f_fail.close()
            print("Limpiando archivos temporales...")
            try:
                if os.path.exists(success_csv_path): os.remove(success_csv_path); print(f"Eliminado: {success_csv_path}")
                if os.path.exists(fail_csv_path): os.remove(fail_csv_path); print(f"Eliminado: {fail_csv_path}")
            except Exception as e: print(f"Error durante la limpieza de archivos temporales: {e}")
            sys.stdout = sys.__stdout__
    def _show_config_view(self): self.progress_frame.pack_forget(); self.config_frame.pack(fill=tk.BOTH, expand=True)
    def _on_closing(self):
        if self.download_thread and self.download_thread.is_alive():
            if messagebox.askokcancel("Salir", "La descarga está en curso. ¿Desea salir?"): self.cancel_event.set(); self.root.destroy()
        else: self.root.destroy()
if __name__ == "__main__":
    root = tk.Tk(); app = SciHubDownloaderApp(root); root.mainloop()
