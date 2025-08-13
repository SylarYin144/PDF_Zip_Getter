import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog, scrolledtext
import threading
from queue import Queue, Empty

class ConfigurationGUI:
    def __init__(self, master):
        self.master = master
        master.title("Configuración de Sci-Hub Downloader")
        master.geometry("600x450")

        self.config = {}
        self.cancelled = True
        self.total_articles = 0

        # Frame principal
        main_frame = tk.Frame(master, padx=10, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- Frame de Configuración (visible al inicio) ---
        self.config_frame = tk.Frame(main_frame)
        self.config_frame.pack(fill=tk.BOTH, expand=True)

        paths_frame = tk.LabelFrame(self.config_frame, text="Rutas de Archivos", padx=10, pady=10)
        paths_frame.pack(fill=tk.X, pady=5)

        self.input_path = tk.StringVar()
        self.zip_path = tk.StringVar()
        self.report_path = tk.StringVar()

        tk.Label(paths_frame, text="Archivo de DOIs (Excel/CSV):").grid(row=0, column=0, sticky=tk.W, pady=2)
        tk.Entry(paths_frame, textvariable=self.input_path, width=50).grid(row=0, column=1, padx=5)
        tk.Button(paths_frame, text="...", command=self.browse_input_file).grid(row=0, column=2)

        tk.Label(paths_frame, text="Guardar ZIP en:").grid(row=1, column=0, sticky=tk.W, pady=2)
        tk.Entry(paths_frame, textvariable=self.zip_path, width=50).grid(row=1, column=1, padx=5)
        tk.Button(paths_frame, text="...", command=self.browse_zip_file).grid(row=1, column=2)

        tk.Label(paths_frame, text="Guardar Reporte Excel en:").grid(row=2, column=0, sticky=tk.W, pady=2)
        tk.Entry(paths_frame, textvariable=self.report_path, width=50).grid(row=2, column=1, padx=5)
        tk.Button(paths_frame, text="...", command=self.browse_report_file).grid(row=2, column=2)

        delays_frame = tk.LabelFrame(self.config_frame, text="Retrasos (segundos)", padx=10, pady=10)
        delays_frame.pack(fill=tk.X, pady=5)

        self.inter_doi_delay = tk.StringVar(value=str(INTER_DOI_DELAY_SECONDS))
        self.mirror_switch_delay = tk.StringVar(value=str(MIRROR_SWITCH_DELAY_SECONDS))

        tk.Label(delays_frame, text="Retraso entre DOIs:").pack(side=tk.LEFT, padx=5)
        tk.Entry(delays_frame, textvariable=self.inter_doi_delay, width=5).pack(side=tk.LEFT)
        tk.Label(delays_frame, text="Retraso al cambiar de mirror:").pack(side=tk.LEFT, padx=(20, 5))
        tk.Entry(delays_frame, textvariable=self.mirror_switch_delay, width=5).pack(side=tk.LEFT)

        mirrors_frame = tk.LabelFrame(self.config_frame, text="Mirrors de Sci-Hub (separados por coma)", padx=10, pady=10)
        mirrors_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.mirrors_text = scrolledtext.ScrolledText(mirrors_frame, wrap=tk.WORD, height=5)
        self.mirrors_text.pack(fill=tk.BOTH, expand=True)
        self.mirrors_text.insert(tk.END, ",".join(DEFAULT_SCI_HUB_MIRRORS_EXAMPLE))

        buttons_frame = tk.Frame(self.config_frame)
        buttons_frame.pack(fill=tk.X, pady=10)

        self.start_button = tk.Button(buttons_frame, text="Iniciar Proceso", command=self.start_process, bg="green", fg="white")
        self.start_button.pack(side=tk.RIGHT, padx=5)
        self.cancel_button = tk.Button(buttons_frame, text="Cancelar", command=self.cancel_process)
        self.cancel_button.pack(side=tk.RIGHT)

        # --- Frame de Progreso (inicialmente oculto) ---
        self.progress_frame = tk.Frame(main_frame)

        tk.Label(self.progress_frame, text="Progreso Total (Artículos Buscados):", font="-weight bold").grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(10, 0))
        self.pbar_searched = ttk.Progressbar(self.progress_frame, length=300)
        self.pbar_searched.grid(row=1, column=0, sticky=tk.EW, padx=5, pady=(0, 10))
        self.pbar_searched_label_var = tk.StringVar(value="0/0 (0.00%)")
        tk.Label(self.progress_frame, textvariable=self.pbar_searched_label_var).grid(row=1, column=1, sticky=tk.W, padx=5)

        tk.Label(self.progress_frame, text="Éxito (Encontrados de Buscados):", font="-weight bold").grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=(10, 0))
        self.pbar_found_of_searched = ttk.Progressbar(self.progress_frame, length=300)
        self.pbar_found_of_searched.grid(row=3, column=0, sticky=tk.EW, padx=5, pady=(0, 10))
        self.pbar_found_of_searched_label_var = tk.StringVar(value="0/0 (0.00%)")
        tk.Label(self.progress_frame, textvariable=self.pbar_found_of_searched_label_var).grid(row=3, column=1, sticky=tk.W, padx=5)

        tk.Label(self.progress_frame, text="Éxito (Encontrados del Total):", font="-weight bold").grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=(10, 0))
        self.pbar_found_of_total = ttk.Progressbar(self.progress_frame, length=300)
        self.pbar_found_of_total.grid(row=5, column=0, sticky=tk.EW, padx=5, pady=(0, 10))
        self.pbar_found_of_total_label_var = tk.StringVar(value="0/0 (0.00%)")
        tk.Label(self.progress_frame, textvariable=self.pbar_found_of_total_label_var).grid(row=5, column=1, sticky=tk.W, padx=5)

        tk.Label(self.progress_frame, text="Tiempo Estimado Restante:", font="-weight bold").grid(row=6, column=0, columnspan=2, sticky=tk.W, pady=(10,0))
        self.eta_label_var = tk.StringVar(value="Calculando...")
        tk.Label(self.progress_frame, textvariable=self.eta_label_var).grid(row=7, column=0, columnspan=2, sticky=tk.W, padx=5)

        self.current_article_frame = tk.LabelFrame(self.progress_frame, text="Procesando Artículo", padx=10, pady=10)
        self.current_article_frame.grid(row=8, column=0, columnspan=2, sticky="ew", pady=10)
        self.current_article_label_var = tk.StringVar(value="Iniciando...")
        tk.Label(self.current_article_frame, textvariable=self.current_article_label_var, justify=tk.LEFT).pack(anchor=tk.W)

        self.final_status_label_var = tk.StringVar(value="Proceso en ejecución...")
        tk.Label(self.progress_frame, textvariable=self.final_status_label_var, pady=20, font="-weight bold").grid(row=9, column=0, columnspan=2, sticky=tk.W)

        self.source_stats_frame = tk.LabelFrame(self.progress_frame, text="Estadísticas de Origen", padx=10, pady=10)
        self.source_stats_frame.grid(row=10, column=0, columnspan=2, sticky="ew", pady=10)
        self.source_stats_labels = {}

        self.progress_frame.columnconfigure(0, weight=1)


    def process_queue(self):
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
                    # Bar 1
                    self.pbar_searched['value'] = searched
                    searched_perc = (searched / self.total_articles) * 100
                    self.pbar_searched_label_var.set(f"{searched}/{self.total_articles} ({searched_perc:.2f}%)")

                    # Bar 2
                    if searched > 0:
                        self.pbar_found_of_searched['maximum'] = searched
                        self.pbar_found_of_searched['value'] = found
                        found_of_searched_perc = (found / searched) * 100
                        self.pbar_found_of_searched_label_var.set(f"{found}/{searched} ({found_of_searched_perc:.2f}%)")

                    # Bar 3
                    self.pbar_found_of_total['value'] = found
                    found_of_total_perc = (found / self.total_articles) * 100
                    self.pbar_found_of_total_label_var.set(f"{found}/{self.total_articles} ({found_of_total_perc:.2f}%)")

            elif message['type'] == 'done':
                self.final_status_label_var.set(message['message'])
                self.cancel_button.config(text="Cerrar", state=tk.NORMAL)
                return # Stop polling

            elif message['type'] == 'current_article':
                self.current_article_label_var.set(message['value'])

            elif message['type'] == 'time_update':
                self.eta_label_var.set(message['value'])

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
                self.final_status_label_var.set(f"ERROR: {message['message']}")
                self.cancel_button.config(text="Cerrar", state=tk.NORMAL)
                return # Stop polling

        except Empty:
            pass

        self.master.after(100, self.process_queue)


    def browse_input_file(self):
        path = filedialog.askopenfilename(title="Seleccionar archivo con DOIs", filetypes=(("Archivos Excel", "*.xlsx *.xls"), ("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*")))
        if path:
            self.input_path.set(path)

    def browse_zip_file(self):
        path = filedialog.asksaveasfilename(title="Guardar archivo ZIP como...", defaultextension=".zip", filetypes=(("Archivos ZIP", "*.zip"),))
        if path:
            self.zip_path.set(path)

    def browse_report_file(self):
        path = filedialog.asksaveasfilename(title="Guardar Reporte Excel como...", defaultextension=".xlsx", filetypes=(("Archivos Excel", "*.xlsx"),))
        if path:
            self.report_path.set(path)

    def start_process(self):
        input_file = self.input_path.get()
        zip_file = self.zip_path.get()

        if not input_file or not zip_file:
            messagebox.showerror("Error", "Debe especificar la ruta del archivo de entrada y del archivo ZIP de salida.")
            return

        try:
            inter_doi = int(self.inter_doi_delay.get())
            mirror_switch = int(self.mirror_switch_delay.get())
            if inter_doi < 0 or mirror_switch < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Los valores de retraso deben ser números enteros no negativos.")
            return

        mirrors = [m.strip() for m in self.mirrors_text.get("1.0", tk.END).split(',') if m.strip()]
        if not mirrors:
            messagebox.showerror("Error", "Debe especificar al menos un mirror de Sci-Hub.")
            return

        self.config = {
            "input_file_path": input_file,
            "zip_path": zip_file,
            "excel_report_path": self.report_path.get(),
            "user_inter_doi_delay": inter_doi,
            "user_mirror_switch_delay": mirror_switch,
            "user_defined_mirrors": mirrors,
        }
        self.cancelled = False

        # Switch to the progress view
        self.config_frame.pack_forget()
        self.progress_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Create queue for thread communication
        self.progress_queue = Queue()

        # Start worker thread
        self.worker_thread = threading.Thread(
            target=download_pdfs_from_file,
            args=(self.config, self.progress_queue)
        )
        self.worker_thread.start()

        # Start polling the queue
        self.master.after(100, self.process_queue)

    def cancel_process(self):
        self.cancelled = True
        self.master.destroy()


# import tkinter.scrolledtext as st # GUI logging disabled
import pandas as pd
import requests
import zipfile
import os
import re 
import time
import sys 
import os
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

# --- Configuration Constants (Primarily for defaults now) ---
DEFAULT_SCI_HUB_MIRRORS_EXAMPLE = ["https://sci-hub.se/", "https://sci-hub.st/", "https://sci-hub.box/", "https://sci-hub.ru/", "https://sci-hub.red/"]
INTER_DOI_DELAY_SECONDS = 5 
MIRROR_SWITCH_DELAY_SECONDS = 3
STANDARD_USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.127 Safari/537.36'
# --- End Configuration Constants ---

import base64 # For JS PDF Fetch helper

# Helper function for fetching PDF content via JavaScript
def get_pdf_content_via_js(driver, pdf_url):
    script = """
    const callback = arguments[arguments.length - 1];
    fetch(arguments[0])
        .then(response => {
            if (!response.ok) {
                throw new Error('Network response was not ok: ' + response.status + ' ' + response.statusText);
            }
            return response.blob();
        })
        .then(blob => {
            if (blob.type !== 'application/pdf') {
                // Allow for cases where Content-Type might be octet-stream but it's a PDF
                // This is a common issue with some servers.
                // We might rely on the URL ending with .pdf or other heuristics if needed,
                // but for now, a strict check on blob.type might be too restrictive.
                // Let's log it but proceed if the blob has size.
                console.warn('JS Fetch: Content-Type is ' + blob.type + ' for URL ' + arguments[0] + '. Proceeding if blob has size.');
                if (blob.type !== 'application/pdf' && !arguments[0].toLowerCase().endsWith('.pdf') && blob.type !== 'application/octet-stream') {
                     // If not PDF, not ending with .pdf, and not octet-stream, then it's likely not a PDF.
                     throw new Error('Content-Type is not application/pdf or octet-stream: ' + blob.type);
                }
            }
            const reader = new FileReader();
            reader.onloadend = () => {
                // reader.result is data:application/pdf;base64,xxxxx
                // We only want the xxxxx part
                const base64Marker = ';base64,';
                const base64Data = reader.result.substring(reader.result.indexOf(base64Marker) + base64Marker.length);
                callback(base64Data);
            };
            reader.onerror = (err) => {
                console.error('FileReader error:', err);
                callback({error: 'FileReader error: ' + err.toString()});
            };
            reader.readAsDataURL(blob);
        })
        .catch(error => {
            console.error('JS Fetch error:', error);
            callback({error: 'JS Fetch error: ' + error.toString()});
        });
    """
    try:
        # Increased async script timeout
        driver.set_script_timeout(90) # seconds for the async script to complete
        result = driver.execute_async_script(script, pdf_url)
        if isinstance(result, dict) and 'error' in result:
            # print(f"JS Fetch Helper: Error reported from JS for {pdf_url}: {result['error']}") # Reduced noise
            return None
        if result:
            return base64.b64decode(result)
        # print(f"JS Fetch Helper: No result or empty result from JS for {pdf_url}") # Reduced noise
        return None
    except Exception as e:
        # print(f"JS Fetch Helper: Exception during execute_async_script for {pdf_url}: {e}") # Reduced noise
        return None

def format_and_log_article_status(original_row_data, doi, title, current_article_num, total_articles, 
                                  successful_downloads_count_for_stats,
                                  mirror_attempts_details, 
                                  overall_doi_status, current_user_inter_doi_delay, is_retry=False, failed_articles_data_len=0):
    try:
        # --- Calculations ---
        buscados_percentage = (current_article_num / total_articles) * 100 if total_articles > 0 else 0
        obtenidos_percentage = (successful_downloads_count_for_stats / total_articles) * 100 if total_articles > 0 else 0
        parcial_percentage = (successful_downloads_count_for_stats / current_article_num) * 100 if current_article_num > 0 else 0

        log_lines = []

        # --- Initial Info Block ---
        retry_prefix = "[REINTENTO] " if is_retry else ""
        log_lines.append(f"{retry_prefix}Artículo: {current_article_num}/{total_articles} ({buscados_percentage:.2f}%)")
        log_lines.append(f"Título: {title if title else 'N/A'}")

        # Simplified author and journal info extraction for concise logging
        author_val = original_row_data.get('First Author', 'N/A')
        journal_title_val = original_row_data.get('Journal/Book', 'N/A')
        pub_year_val = original_row_data.get('Publication Year', 'N/A')
        
        log_lines.append(f"First Author: {author_val}")
        log_lines.append(f"Journal/Book: {journal_title_val}")
        log_lines.append(f"Publication Year: {pub_year_val}")
        log_lines.append(f"DOI: {doi}")
        log_lines.append("")

        # --- Mirror Attempts ---
        for i, attempt in enumerate(mirror_attempts_details):
            mirror_url, status, reason = attempt
            try:
                domain_match = re.search(r'https?://(?:www\.)?([a-zA-Z0-9.-]+)(?:/|$)', mirror_url)
                mirror_short_name = domain_match.group(1) if domain_match else mirror_url[-20:]
            except Exception:
                mirror_short_name = mirror_url[-20:]
            log_lines.append(f"Intento Mirror {i+1} ({mirror_short_name}): {status}. {reason if reason else ''}".strip())
        log_lines.append("")

        # --- Status and Totals ---
        if overall_doi_status.startswith("OBTENIDO"):
            log_lines.append("## OBTENIDO ##")
        else:
            log_lines.append("## FALTANTE ##")
        log_lines.append("")
        
        log_lines.append(f"Total descargados parcial: {successful_downloads_count_for_stats}/{current_article_num}\t({parcial_percentage:.2f}%)")
        log_lines.append(f"Total Descargados:         {successful_downloads_count_for_stats}/{total_articles}\t({obtenidos_percentage:.2f}%)")
        log_lines.append("")

        # --- Waiting Message ---
        # Simplified logic: print if not the absolute last item being processed.
        is_last_main_loop_item_with_no_retries = (current_article_num == total_articles) and (not is_retry) and (failed_articles_data_len == 0)
        is_last_retry_item = is_retry and (failed_articles_data_len <= 1)

        if not is_last_main_loop_item_with_no_retries and not is_last_retry_item:
            log_lines.append(f"Esperando {current_user_inter_doi_delay} segundos…")

        # --- Final Print ---
        formatted_message = "\n".join(log_lines)
        print(f"\n{formatted_message}")

    except Exception as e:
        print(f"\nError al formatear log para DOI {doi}: {e}")
        print(f"Fallback Log: DOI: {doi}, Status: {overall_doi_status}\n")

def clean_filename(title):
    return re.sub(r'[\\/*?:"<>|]', '_', title)

def extract_pdf_link_from_html(article_page_url, session):
    try:
        response = session.get(article_page_url, timeout=30); response.raise_for_status()
        soup = BeautifulSoup(response.content, 'html.parser')
        iframe = soup.find('iframe', id='pdf')
        if iframe and iframe.get('src'):
            pdf_src = iframe['src']
            if pdf_src.startswith('//'): return 'https:' + pdf_src
            elif pdf_src.startswith('/'): return urljoin(article_page_url, pdf_src)
            return pdf_src
        embed = soup.find('embed', attrs={'type': 'application/pdf'})
        if embed and embed.get('src'):
            pdf_src = embed['src']
            return urljoin(article_page_url, pdf_src)
        return None
    except requests.exceptions.RequestException:
        return None
    except Exception:
        return None

def download_from_google_scholar_old(doi, title, session):
    scholar_url = f"https://scholar.google.com/scholar?q={doi}"
    try:
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'}
        response = session.get(scholar_url, headers=headers, timeout=30)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, 'html.parser')
        potential_links = []
        for link_tag in soup.find_all('a', href=True):
            href = link_tag['href']
            link_text = link_tag.get_text().lower()
            if href.lower().endswith('.pdf') or '[pdf]' in link_text or 'pdf' in link_text:
                if not href.startswith('http'):
                    href = urljoin(scholar_url, href)
                if 'pdf' in href.lower() and not any(x in href.lower() for x in ['view', 'download=false', 'scholar.google']):
                     potential_links.append(href)
                elif href.lower().endswith('.pdf'):
                    potential_links.append(href)
        unique_potential_links = []
        for plink in potential_links:
            if plink not in unique_potential_links:
                unique_potential_links.append(plink)
        for pdf_url in unique_potential_links:
            try:
                head_response = session.head(pdf_url, headers=headers, timeout=20, allow_redirects=True)
                head_response.raise_for_status()
                content_type = head_response.headers.get('Content-Type', '').lower()
                if 'application/pdf' in content_type:
                    pdf_response = session.get(pdf_url, headers=headers, timeout=60, stream=True)
                    pdf_response.raise_for_status()
                    get_content_type = pdf_response.headers.get('Content-Type', '').lower()
                    if 'application/pdf' in get_content_type:
                        pdf_content = pdf_response.content
                        domain_match = re.search(r'https?://(?:www\.)?([a-zA-Z0-9.-]+)(?:/|$)', pdf_url)
                        source_domain = domain_match.group(1) if domain_match else "Unknown Domain"
                        return pdf_content, f"OBTENIDO (Google Scholar via {source_domain})"
            except requests.exceptions.HTTPError:
                pass
            except requests.exceptions.Timeout:
                pass
            except requests.exceptions.RequestException:
                pass
            except Exception:
                pass
        return None, f"FALLO - No PDF en Google Scholar ({scholar_url})"
    except requests.exceptions.RequestException:
        return None, f"FALLO - Error búsqueda Google Scholar ({scholar_url})"
    except Exception:
        return None, f"FALLO - Error inesperado Google Scholar ({scholar_url})"

def download_from_google_scholar(doi, title, session):
    scholar_url = f"https://scholar.google.com/scholar?hl=en&q={doi}"
    try:
        headers = {
            'User-Agent': STANDARD_USER_AGENT,
            'Accept-Language': 'en-US,en;q=0.9',
            'Accept-Encoding': 'gzip, deflate, br',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
            'Connection': 'keep-alive'
        }
        response = session.get(scholar_url, headers=headers, timeout=30)
        response.raise_for_status()
        content_type_initial = response.headers.get('Content-Type', '').lower()
        if 'application/pdf' in content_type_initial:
            if len(response.content) > 1000:
                 return response.content, f"OBTENIDO (Google Scholar Direct Response - {scholar_url})"
        soup = BeautifulSoup(response.content, 'html.parser')
        potential_links = []
        for link_tag in soup.find_all('a', href=True):
            href = link_tag['href']
            link_text = link_tag.get_text(strip=True).lower()
            is_pdf_link = False
            if href.lower().endswith('.pdf'):
                is_pdf_link = True
            elif '[pdf]' in link_text or 'pdf' in link_text or link_tag.find(lambda tag: tag.name == 'span' and 'pdf' in tag.get_text(strip=True).lower()):
                 is_pdf_link = True
            if is_pdf_link:
                if not href.startswith('http'):
                    href = urljoin(scholar_url, href)
                if 'scholar.google.com' in href.lower() and not href.lower().endswith('.pdf'):
                    continue
                if any(x in href.lower() for x in [' Morales', ' Privacy', ' Terms', ' Sign in', ' Settings', ' My Citations', ' Profiles', ' cited by', ' related articles', ' versions', ' web search', 'javascript:void(0)']):
                    continue
                if href.endswith("#"):
                    continue
                potential_links.append(href)
        for result_div in soup.find_all('div', class_='gs_ri'):
            title_link_tag = result_div.find('h3', class_='gs_rt').find('a', href=True) if result_div.find('h3', class_='gs_rt') else None
            pdf_div = result_div.find_next_sibling('div', class_='gs_ggs')
            if pdf_div:
                pdf_link_tag = pdf_div.find('a', href=True)
                if pdf_link_tag and pdf_link_tag['href'].lower().endswith('.pdf'):
                    href = pdf_link_tag['href']
                    if not href.startswith('http'): href = urljoin(scholar_url, href)
                    potential_links.append(href)
            if title_link_tag and title_link_tag['href'].lower().endswith('.pdf'):
                 href = title_link_tag['href']
                 if not href.startswith('http'): href = urljoin(scholar_url, href)
                 potential_links.append(href)
        unique_potential_links = []
        for plink in potential_links:
            if plink not in unique_potential_links:
                unique_potential_links.append(plink)
        for pdf_url in unique_potential_links:
            try:
                head_response = session.head(pdf_url, headers=headers, timeout=20, allow_redirects=True)
                head_response.raise_for_status()
                content_type = head_response.headers.get('Content-Type', '').lower()
                if 'application/pdf' in content_type:
                    pdf_response = session.get(pdf_url, headers=headers, timeout=60, stream=True)
                    pdf_response.raise_for_status()
                    get_content_type = pdf_response.headers.get('Content-Type', '').lower()
                    if 'application/pdf' in get_content_type:
                        pdf_content = pdf_response.content
                        if len(pdf_content) < 1000:
                            continue
                        domain_match = re.search(r'https?://(?:www\.)?([a-zA-Z0-9.-]+)(?:/|$)', pdf_url)
                        source_domain = domain_match.group(1) if domain_match else "Unknown Domain"
                        return pdf_content, f"OBTENIDO (Google Scholar via {source_domain})"
                else:
                    pdf_response = session.get(pdf_url, headers=headers, timeout=60, stream=True)
                    pdf_response.raise_for_status()
                    get_content_type = pdf_response.headers.get('Content-Type', '').lower()
                    if 'application/pdf' in get_content_type:
                        pdf_content = pdf_response.content
                        if len(pdf_content) < 1000:
                            continue
                        domain_match = re.search(r'https?://(?:www\.)?([a-zA-Z0-9.-]+)(?:/|$)', pdf_url)
                        source_domain = domain_match.group(1) if domain_match else "Unknown Domain"
                        return pdf_content, f"OBTENIDO (Google Scholar via {source_domain} - GET Fallback)"
            except requests.exceptions.HTTPError:
                pass
            except requests.exceptions.Timeout:
                pass
            except requests.exceptions.RequestException:
                pass
            except Exception:
                pass
        return None, f"FALLO - No PDF en Google Scholar ({scholar_url})"
    except requests.exceptions.RequestException:
        return None, f"FALLO - Error búsqueda Google Scholar ({scholar_url})"
    except Exception:
        return None, f"FALLO - Error inesperado Google Scholar ({scholar_url})"

def download_with_selenium_google_scholar(driver, doi, title):
    # print(f"SELENIUM GS: Searching Google Scholar for DOI: {doi} (Title: {title if title else 'N/A'})") # Reduced noise
    scholar_url = f"https://scholar.google.com/scholar?hl=en&q={doi}"
    pdf_content = None
    status_message = f"FALLO - No PDF en Google Scholar (Selenium) ({scholar_url})"
    scholar_page_loaded = False
    initial_load_attempts = 0
    max_initial_load_attempts = 2
    while initial_load_attempts < max_initial_load_attempts and not scholar_page_loaded:
        try:
            driver.set_page_load_timeout(120)
            # print(f"SELENIUM GS: Attempt {initial_load_attempts + 1}/{max_initial_load_attempts} to load {scholar_url}") # Reduced noise
            driver.get(scholar_url)
            WebDriverWait(driver, 25).until(
                EC.presence_of_element_located((By.ID, "gs_res_ccl_mid"))
            )
            scholar_page_loaded = True
            # print(f"SELENIUM GS: Successfully loaded {scholar_url}") # Reduced noise
            break
        except TimeoutException as e_load:
            initial_load_attempts += 1
            # print(f"SELENIUM GS: Initial page load attempt {initial_load_attempts}/{max_initial_load_attempts} timed out for {scholar_url}.") # Reduced noise
            if initial_load_attempts < max_initial_load_attempts:
                time.sleep(5)
            else:
                # print(f"SELENIUM GS: All initial page load attempts timed out for {scholar_url}.") # Reduced noise
                return None, f"FALLO - Timeout persistente carga Google Scholar (Selenium) ({scholar_url})"
        except Exception as e_gen_load:
            initial_load_attempts += 1
            # print(f"SELENIUM GS: Unexpected error during initial page load attempt {initial_load_attempts}/{max_initial_load_attempts} for {scholar_url}: {e_gen_load}") # Reduced noise
            if initial_load_attempts < max_initial_load_attempts:
                time.sleep(5)
            else:
                return None, f"FALLO - Error inesperado carga Google Scholar (Selenium) ({scholar_url}, {str(e_gen_load)[:100]})"
    if not scholar_page_loaded:
        return None, f"FALLO - Timeout persistente carga Google Scholar (Selenium) ({scholar_url})"
    try:
        pdf_links_elements = []
        try:
            pdf_links_elements = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.PARTIAL_LINK_TEXT, "[PDF]"))
            )
        except TimeoutException:
            # print(f"SELENIUM GS: No direct '[PDF]' links found for {doi}. Trying other methods.") # Reduced noise
            pass
        if not pdf_links_elements:
            try:
                pdf_links_elements = WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf')]"))
                )
            except TimeoutException:
                # print(f"SELENIUM GS: No links with '.pdf' in href found for {doi}.") # Reduced noise
                pass
        # print(f"SELENIUM GS: Found {len(pdf_links_elements)} potential PDF link elements.") # Reduced noise
        pdf_urls_to_try = []
        for link_el in pdf_links_elements:
            href = link_el.get_attribute('href')
            if href and href not in pdf_urls_to_try:
                pdf_urls_to_try.append(href)
        # print(f"SELENIUM GS: Extracted {len(pdf_urls_to_try)} unique URLs to attempt.") # Reduced noise
        for pdf_url_attempt in pdf_urls_to_try:
            # print(f"SELENIUM GS: Processing link: {pdf_url_attempt}") # Reduced noise
            if pdf_url_attempt.lower().endswith('.pdf'):
                # print(f"SELENIUM GS: Attempting direct JS fetch for PDF-like URL: {pdf_url_attempt}") # Reduced noise
                pdf_content = get_pdf_content_via_js(driver, pdf_url_attempt)
                if pdf_content and len(pdf_content) > 1024:
                    # print(f"SELENIUM GS: Successfully fetched PDF via JS from direct URL: {pdf_url_attempt}") # Reduced noise
                    return pdf_content, f"OBTENIDO (Google Scholar Selenium JS Fetch - {pdf_url_attempt})"
                else:
                    # print(f"SELENIUM GS: JS fetch from {pdf_url_attempt} did not yield valid PDF content.") # Reduced noise
                    pdf_content = None
            if not pdf_content:
                # print(f"SELENIUM GS: Attempting navigation to: {pdf_url_attempt}") # Reduced noise
                try:
                    driver.get(pdf_url_attempt)
                    time.sleep(5)
                    current_url_after_nav = driver.current_url
                    # print(f"SELENIUM GS: Navigated. Current URL: {current_url_after_nav}") # Reduced noise
                    pdf_content = get_pdf_content_via_js(driver, current_url_after_nav)
                    if pdf_content and len(pdf_content) > 1024:
                        # print(f"SELENIUM GS: Successfully fetched PDF via JS after navigation from {pdf_url_attempt} to {current_url_after_nav}") # Reduced noise
                        return pdf_content, f"OBTENIDO (Google Scholar Selenium Nav & JS Fetch - {current_url_after_nav})"
                    else:
                        # if pdf_content is None: # Reduced noise
                            # print(f"SELENIUM GS: JS Fetch failed or returned non-PDF for {current_url_after_nav} (after nav from {pdf_url_attempt}). Will check for embeds.") # Reduced noise
                        # elif pdf_content: # Reduced noise
                            # print(f"SELENIUM GS: JS Fetched content from {current_url_after_nav} (after nav from {pdf_url_attempt}) was too small ({len(pdf_content)} bytes). Will check for embeds.") # Reduced noise
                        # else: # Reduced noise
                            # print(f"SELENIUM GS: JS fetch from {current_url_after_nav} (after nav from {pdf_url_attempt}) did not yield valid PDF. Will check for embeds.") # Reduced noise
                        pdf_content = None
                    if not pdf_content:
                        # print(f"SELENIUM GS: Checking for embedded PDF on {current_url_after_nav}") # Reduced noise
                        try:
                            embed_element = WebDriverWait(driver, 5).until(
                                EC.presence_of_element_located((By.XPATH, "//embed[@type='application/pdf']"))
                            )
                            if embed_element:
                                embed_src = embed_element.get_attribute('src')
                                if embed_src:
                                    embed_src_abs = urljoin(current_url_after_nav, embed_src)
                                    # print(f"SELENIUM GS: Found <embed> with src: {embed_src_abs}. Attempting JS fetch.") # Reduced noise
                                    pdf_content = get_pdf_content_via_js(driver, embed_src_abs)
                                    if pdf_content and len(pdf_content) > 1024:
                                        # print(f"SELENIUM GS: Successfully fetched PDF from <embed> src: {embed_src_abs}") # Reduced noise
                                        return pdf_content, f"OBTENIDO (Google Scholar Selenium Embed JS Fetch - {embed_src_abs})"
                                    else:
                                        pdf_content = None
                        except TimeoutException:
                            # print(f"SELENIUM GS: No <embed type='application/pdf'> found on {current_url_after_nav}.") # Reduced noise
                            pass
                        if not pdf_content:
                            try:
                                iframe_element = WebDriverWait(driver, 5).until(
                                    EC.presence_of_element_located((By.XPATH, "//iframe[contains(@src, '.pdf')]"))
                                )
                                if iframe_element:
                                    iframe_src = iframe_element.get_attribute('src')
                                    if iframe_src:
                                        iframe_src_abs = urljoin(current_url_after_nav, iframe_src)
                                        # print(f"SELENIUM GS: Found <iframe> with PDF-like src: {iframe_src_abs}. Attempting JS fetch.") # Reduced noise
                                        pdf_content = get_pdf_content_via_js(driver, iframe_src_abs)
                                        if pdf_content and len(pdf_content) > 1024:
                                            # print(f"SELENIUM GS: Successfully fetched PDF from <iframe> src: {iframe_src_abs}") # Reduced noise
                                            return pdf_content, f"OBTENIDO (Google Scholar Selenium Iframe JS Fetch - {iframe_src_abs})"
                                        else:
                                            pdf_content = None
                            except TimeoutException:
                                # print(f"SELENIUM GS: No <iframe> with PDF-like src found on {current_url_after_nav}.") # Reduced noise
                                pass
                except TimeoutException as e_nav_timeout:
                    # print(f"SELENIUM GS: Timeout during navigation or subsequent operations for {pdf_url_attempt}: {e_nav_timeout}") # Reduced noise
                    pass
                except Exception as e_nav:
                    # print(f"SELENIUM GS: Error during navigation or subsequent operations for {pdf_url_attempt}: {e_nav}") # Reduced noise
                    pass
            if pdf_content:
                return pdf_content, status_message
        status_message = f"FALLO - No PDF obtained after trying all potential links (Selenium GS)"
    except TimeoutException as e_element_find:
        # print(f"SELENIUM GS: Element TimeoutException after page load, while finding links for {doi} on {driver.current_url if driver else scholar_url}: {e_element_find}") # Reduced noise
        status_message = f"FALLO - Timeout localizando elementos post-carga en Google Scholar (Selenium) ({driver.current_url if driver else scholar_url})"
    except NoSuchElementException as e_no_such:
        # print(f"SELENIUM GS: NoSuchElementException after page load, while finding links for {doi} on {driver.current_url if driver else scholar_url}: {e_no_such}") # Reduced noise
        status_message = f"FALLO - Elemento no encontrado post-carga en Google Scholar (Selenium) ({driver.current_url if driver else scholar_url})"
    except Exception as e_general:
        # print(f"SELENIUM GS: An unexpected error occurred after page load for DOI {doi} at {driver.current_url if driver else scholar_url}: {e_general}") # Reduced noise
        status_message = f"FALLO - Error inesperado post-carga en Google Scholar (Selenium) ({driver.current_url if driver else scholar_url}, {str(e_general)[:100]})"
    return None, status_message

def download_with_selenium_pmc(driver, doi, title):
    # print(f"SELENIUM PMC: Searching PubMed Central for DOI: {doi} (Title: {title if title else 'N/A'})") # Reduced noise
    search_url = f"https://www.ncbi.nlm.nih.gov/pmc/?term={doi}"
    pdf_content = None
    status_message = f"FALLO - No PDF en PMC (Selenium) ({search_url})"
    article_url = None
    try:
        driver.set_page_load_timeout(120)
        driver.get(search_url)
        WebDriverWait(driver, 25).until(
            EC.presence_of_element_located((By.CLASS_NAME, "rprt"))
        )
        article_link_element = None
        try:
            possible_article_links = driver.find_elements(By.CSS_SELECTOR, "div.rprt .title a")
            if not possible_article_links:
                 possible_article_links = driver.find_elements(By.XPATH, "//div[contains(@class, 'rprt')]//a[contains(@href, 'articles/PMC')]")
            if possible_article_links:
                article_link_element = possible_article_links[0]
                article_url = article_link_element.get_attribute('href')
                # print(f"SELENIUM PMC: Found article link: {article_url}. Navigating...") # Reduced noise
                driver.set_page_load_timeout(120)
                driver.get(article_url)
                # print(f"SELENIUM PMC: Navigation to article page {article_url} presumably successful.") # Reduced noise
            else:
                # print(f"SELENIUM PMC: No clear article link found in search results for {doi}. Assuming current page ({driver.current_url}) might be the article page or search failed.") # Reduced noise
                article_url = driver.current_url
        except Exception as e_inner_nav:
            # print(f"SELENIUM PMC: Error during article link navigation for DOI {doi}: {e_inner_nav}. Proceeding with current page {driver.current_url}.") # Reduced noise
            article_url = driver.current_url
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(@href, '.pdf') or contains(translate(., 'PDF', 'pdf'), 'pdf')]"))
        )
        pdf_link_elements = []
        selectors = [
            (By.XPATH, "//a[contains(@class, 'format-pdf') and contains(@href, '.pdf')]"),
            (By.XPATH, "//a[contains(translate(., 'PDF', 'pdf'), 'pdf') and contains(@href, '.pdf')]"),
            (By.PARTIAL_LINK_TEXT, "Download PDF"),
            (By.CSS_SELECTOR, "a.pdf-button[href$='.pdf']"),
            (By.XPATH, "//a[contains(@href, '.pdf') and .//span[contains(translate(., 'PDF', 'pdf'), 'pdf')]]")
        ]
        for by, selector_val in selectors:
            try:
                elements = WebDriverWait(driver, 5).until(
                    EC.presence_of_all_elements_located((by, selector_val))
                )
                if elements:
                    pdf_link_elements.extend(elements)
                    # print(f"SELENIUM PMC: Found elements with selector {by} {selector_val}") # Reduced noise
            except TimeoutException:
                # print(f"SELENIUM PMC: Timeout for selector {by} {selector_val}") # Reduced noise
                pass
        # print(f"SELENIUM PMC: Found {len(pdf_link_elements)} potential PDF link elements on article page {article_url if article_url else driver.current_url}.") # Reduced noise
        pdf_urls_to_try = []
        for link_el in pdf_link_elements:
            href = link_el.get_attribute('href')
            if href and href not in pdf_urls_to_try:
                if not href.startswith('http'):
                    base_for_relative = driver.current_url
                    if "ncbi.nlm.nih.gov" in base_for_relative:
                        base_for_relative = "https://www.ncbi.nlm.nih.gov"
                    href = urljoin(base_for_relative, href)
                pdf_urls_to_try.append(href)
        # print(f"SELENIUM PMC: Extracted {len(pdf_urls_to_try)} unique absolute URLs to attempt.") # Reduced noise
        for pdf_url_attempt in pdf_urls_to_try:
            # print(f"SELENIUM PMC: Processing link: {pdf_url_attempt}") # Reduced noise
            if pdf_url_attempt.lower().endswith('.pdf'):
                # print(f"SELENIUM PMC: Attempting direct JS fetch for PDF-like URL: {pdf_url_attempt}") # Reduced noise
                pdf_content = get_pdf_content_via_js(driver, pdf_url_attempt)
                if pdf_content and len(pdf_content) > 1024:
                    # print(f"SELENIUM PMC: Successfully fetched PDF via JS from direct URL: {pdf_url_attempt}") # Reduced noise
                    return pdf_content, f"OBTENIDO (PMC Selenium JS Fetch Direct - {pdf_url_attempt})"
                else:
                    # print(f"SELENIUM PMC: JS fetch from {pdf_url_attempt} (direct) did not yield valid PDF.") # Reduced noise
                    pdf_content = None
            if not pdf_content:
                # print(f"SELENIUM PMC: Attempting navigation to: {pdf_url_attempt}") # Reduced noise
                try:
                    driver.get(pdf_url_attempt)
                    time.sleep(7)
                    current_url_after_nav = driver.current_url
                    # print(f"SELENIUM PMC: Navigated. Current URL is now: {current_url_after_nav}") # Reduced noise
                    # print(f"SELENIUM PMC: Attempting JS fetch on current URL post-navigation: {current_url_after_nav}") # Reduced noise
                    pdf_content = get_pdf_content_via_js(driver, current_url_after_nav)
                    if pdf_content and len(pdf_content) > 1024:
                        # print(f"SELENIUM PMC: Successfully fetched PDF via JS from {current_url_after_nav} (after nav from {pdf_url_attempt})") # Reduced noise
                        return pdf_content, f"OBTENIDO (PMC Selenium Nav & JS Fetch - {current_url_after_nav})"
                    else:
                        # print(f"SELENIUM PMC: JS fetch from {current_url_after_nav} (after nav) did not yield PDF.") # Reduced noise
                        pdf_content = None
                    if not pdf_content:
                        # print(f"SELENIUM PMC: Checking for embedded PDF on {current_url_after_nav}") # Reduced noise
                        try:
                            embed_element = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.XPATH, "//embed[@type='application/pdf']"))
                            )
                            if embed_element:
                                embed_src = embed_element.get_attribute('src')
                                if embed_src:
                                    embed_src_abs = urljoin(current_url_after_nav, embed_src)
                                    # print(f"SELENIUM PMC: Found <embed> with src: {embed_src_abs}. Attempting JS fetch.") # Reduced noise
                                    pdf_content = get_pdf_content_via_js(driver, embed_src_abs)
                                    if pdf_content and len(pdf_content) > 1024:
                                        # print(f"SELENIUM PMC: Successfully fetched PDF from <embed> src: {embed_src_abs}") # Reduced noise
                                        return pdf_content, f"OBTENIDO (PMC Selenium Embed JS Fetch - {embed_src_abs})"
                                    else:
                                        pdf_content = None
                        except TimeoutException:
                            # print(f"SELENIUM PMC: No <embed type='application/pdf'> found on {current_url_after_nav}.") # Reduced noise
                            pass
                        if not pdf_content:
                            try:
                                iframe_element = WebDriverWait(driver, 5).until(
                                    EC.presence_of_element_located((By.XPATH, "//iframe[contains(@src, '.pdf')] | //iframe[contains(@src, 'pdfviewer')]"))
                                )
                                if iframe_element:
                                    iframe_src = iframe_element.get_attribute('src')
                                    if iframe_src:
                                        iframe_src_abs = urljoin(current_url_after_nav, iframe_src)
                                        # print(f"SELENIUM PMC: Found <iframe> with PDF-like src: {iframe_src_abs}. Attempting JS fetch.") # Reduced noise
                                        pdf_content = get_pdf_content_via_js(driver, iframe_src_abs)
                                        if pdf_content and len(pdf_content) > 1024:
                                            # print(f"SELENIUM PMC: Successfully fetched PDF from <iframe> src: {iframe_src_abs}") # Reduced noise
                                            return pdf_content, f"OBTENIDO (PMC Selenium Iframe JS Fetch - {iframe_src_abs})"
                                        else:
                                            pdf_content = None
                            except TimeoutException:
                                # print(f"SELENIUM PMC: No <iframe> with PDF-like src found on {current_url_after_nav}.") # Reduced noise
                                pass
                except TimeoutException as e_nav_timeout:
                    # print(f"SELENIUM PMC: Timeout during navigation to or processing of {pdf_url_attempt}: {e_nav_timeout}") # Reduced noise
                    pass
                except Exception as e_nav:
                    # print(f"SELENIUM PMC: Error during navigation to or processing of {pdf_url_attempt}: {e_nav}") # Reduced noise
                    pass
            if pdf_content:
                 return pdf_content, status_message
        status_message = f"FALLO - No PDF obtained after trying all potential links (Selenium PMC)"
    except TimeoutException as e:
        current_url_for_log = driver.current_url if driver else search_url
        if current_url_for_log == search_url or (article_url and current_url_for_log == article_url and not pdf_content) or current_url_for_log == "about:blank":
            # print(f"SELENIUM PMC: Page load TimeoutException for {current_url_for_log} (DOI {doi}): {e}") # Reduced noise
            status_message = f"FALLO - Timeout carga página PMC (Selenium) ({current_url_for_log})"
        else:
            # print(f"SELENIUM PMC: Element TimeoutException en PMC (Selenium) for DOI {doi} at {current_url_for_log}: {e}") # Reduced noise
            status_message = f"FALLO - Timeout localizando elemento en PMC (Selenium) ({current_url_for_log})"
    except NoSuchElementException as e:
        current_url_for_log = driver.current_url if driver else search_url
        # print(f"SELENIUM PMC: NoSuchElementException en PMC (Selenium) for DOI {doi} at {current_url_for_log}: {e}") # Reduced noise
        status_message = f"FALLO - Elemento no encontrado en PMC (Selenium) ({current_url_for_log})"
    except Exception as e:
        current_url_for_log = driver.current_url if driver else search_url
        # print(f"SELENIUM PMC: An unexpected error occurred with Selenium for DOI {doi} at {current_url_for_log}: {e}") # Reduced noise
        status_message = f"FALLO - Error inesperado en PMC (Selenium) ({current_url_for_log}, {str(e)[:100]})"
    return pdf_content, status_message

import itertools

def download_from_pmc(doi, title, session):
    # print(f"FIXED PMC: Attempting PubMed Central download for DOI: {doi}") # Reduced noise
    try:
        session.headers.update({'User-Agent': STANDARD_USER_AGENT})
        id_conv_url = f"https://www.ncbi.nlm.nih.gov/pmc/utils/idconv/v1.0/?ids={doi}&format=json&tool=my_awesome_tool&email=myemail@example.com"
        try:
            response_id_conv = session.get(id_conv_url, timeout=20)
            response_id_conv.raise_for_status()
            data_id_conv = response_id_conv.json()
        except requests.exceptions.RequestException as e:
            return None, f"FALLO - Error API conversión PMCID para {doi} ({str(e)[:50]})"
        except json.JSONDecodeError:
            return None, f"FALLO - Error decodificando respuesta PMCID para {doi}"
        pmcid = None
        if data_id_conv.get("records") and len(data_id_conv["records"]) > 0:
            record = data_id_conv["records"][0]
            if record.get("pmcid"):
                 pmcid = record["pmcid"]
                 if record.get("status") == "error" and record.get("errmsg") == "invalid article id":
                      pass
            elif record.get("status") == "error":
                return None, f"FALLO - PMCID no encontrado, API devolvió '{record.get('errmsg', 'Unknown error')}' para DOI {doi}"
        if not pmcid:
            return None, f"FALLO - PMCID no encontrado para DOI {doi} (Respuesta: {data_id_conv})"
        efetch_url = f"https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?db=pmc&id={pmcid}&rettype=xml&tool=my_awesome_tool&email=myemail@example.com"
        try:
            response_efetch = session.get(efetch_url, timeout=45)
            response_efetch.raise_for_status()
            xml_content_bytes = response_efetch.content
            root = ET.fromstring(xml_content_bytes)
            pdf_filename_from_xml = None
            pdf_url_from_xml_constructed = None
            namespaces = {'xlink': 'http://www.w3.org/1999/xlink'}
            for tag_name_to_search in ["self-uri", "uri"]:
                combined_iterator = itertools.chain(
                    root.iterfind(f".//{tag_name_to_search}"),
                    root.iterfind(f".//{{{namespaces['xlink']}}}{tag_name_to_search}")
                )
                for element in combined_iterator:
                    content_type = element.get("content-type", "").lower()
                    href_xlink = element.get(f"{{{namespaces['xlink']}}}href")
                    href_plain = element.get("href")
                    current_href_value = None
                    if href_xlink:
                        current_href_value = href_xlink
                    elif href_plain:
                        current_href_value = href_plain
                    if "pdf" in content_type and current_href_value:
                        current_href_value = current_href_value.strip()
                        if current_href_value.lower().endswith('.pdf') or '.pdf?' in current_href_value.lower():
                            pdf_filename_from_xml = current_href_value
                            break
                if pdf_filename_from_xml:
                    break
            if not pdf_filename_from_xml:
                for element in root.iterfind(".//article-id[@pub-id-type='pmc-pdf']"):
                    if element.text:
                        href_value = element.text.strip()
                        if href_value.lower().endswith('.pdf') or '.pdf?' in href_value.lower():
                            pdf_filename_from_xml = href_value
                            break
            if pdf_filename_from_xml:
                if pdf_filename_from_xml.startswith('http://') or pdf_filename_from_xml.startswith('https://'):
                    pdf_url_from_xml_constructed = pdf_filename_from_xml
                elif not pdf_filename_from_xml.startswith('/'):
                    pdf_url_from_xml_constructed = f"https://www.ncbi.nlm.nih.gov/pmc/articles/{pmcid}/pdf/{pdf_filename_from_xml}"
                else:
                    pdf_url_from_xml_constructed = urljoin(f"https://www.ncbi.nlm.nih.gov/pmc/articles/{pmcid}/", pdf_filename_from_xml)
                try:
                    head_response = session.head(pdf_url_from_xml_constructed, timeout=20, allow_redirects=True)
                    head_response.raise_for_status()
                    content_type_head = head_response.headers.get('Content-Type', '').lower()
                    if 'application/pdf' in content_type_head:
                        pdf_response = session.get(pdf_url_from_xml_constructed, timeout=60, stream=True)
                        pdf_response.raise_for_status()
                        content_type_get = pdf_response.headers.get('Content-Type', '').lower()
                        if 'application/pdf' in content_type_get:
                            pdf_content_bytes = pdf_response.content
                            if len(pdf_content_bytes) > 1000:
                                return pdf_content_bytes, f"OBTENIDO (PMC Efetch XML {pmcid})"
                    else:
                        pdf_response = session.get(pdf_url_from_xml_constructed, timeout=60, stream=True)
                        pdf_response.raise_for_status()
                        content_type_get = pdf_response.headers.get('Content-Type', '').lower()
                        if 'application/pdf' in content_type_get:
                            pdf_content_bytes = pdf_response.content
                            if len(pdf_content_bytes) > 1000:
                                return pdf_content_bytes, f"OBTENIDO (PMC Efetch XML - GET Fallback {pmcid})"
                except requests.exceptions.RequestException as e_dl:
                    pass
        except requests.exceptions.RequestException:
            pass
        except ET.ParseError:
            pass
        article_url = f"https://www.ncbi.nlm.nih.gov/pmc/articles/{pmcid}/"
        try:
            response_html = session.get(article_url, timeout=30)
            response_html.raise_for_status()
        except requests.exceptions.RequestException as e:
            return None, f"FALLO - Error obteniendo página HTML PMC ({article_url}, {str(e)[:50]})"
        soup = BeautifulSoup(response_html.content, 'html.parser')
        potential_html_pdf_links = []
        selectors = [
            'div.format-menu a[href$=".pdf"]', 'ul.format-menu a[href$=".pdf"]',
            'div.full-text-links a[href$=".pdf"]', 'li.pdf-link a[href$=".pdf"]',
            'a.format-pdf[href$=".pdf"]', 'a.pdf-button[href$=".pdf"]', 'a.pdf-btn[href$=".pdf"]',
            'a[title*="PDF"][href$=".pdf"]', 'a[data-format="pdf"][href$=".pdf"]',
            'a[download$=".pdf"]',
            'div.buttons.article-actions a.pdf-download[href*="pdf"]'
        ]
        for selector in selectors:
            for link_tag in soup.select(selector):
                href = link_tag.get('href')
                if href: potential_html_pdf_links.append(urljoin(article_url, href.strip()))
        if not potential_html_pdf_links:
            for link_tag in soup.find_all('a', href=lambda h: h is not None and (h.lower().endswith('.pdf') or '.pdf?' in h.lower())):
                href = link_tag.get('href')
                if pmcid.lower() in href.lower() or "articles" in href.lower() or "ftrender" in href.lower():
                     potential_html_pdf_links.append(urljoin(article_url, href.strip()))
        unique_html_pdf_links = []
        for plink in potential_html_pdf_links:
            if plink not in unique_html_pdf_links: unique_html_pdf_links.append(plink)
        if not unique_html_pdf_links:
            return None, f"FALLO - PDF no encontrado en HTML página PMC ({article_url})"
        for pdf_url_html in unique_html_pdf_links:
            try:
                pdf_response = session.get(pdf_url_html, timeout=60, stream=True)
                pdf_response.raise_for_status()
                content_type = pdf_response.headers.get('Content-Type', '').lower()
                if 'application/pdf' in content_type:
                    pdf_data = pdf_response.content
                    if len(pdf_data) > 1000:
                        return pdf_data, f"OBTENIDO (PMC HTML {pmcid})"
            except requests.exceptions.RequestException:
                continue
        return None, f"FALLO - No se pudo descargar PDF desde enlaces HTML PMC ({article_url})"
    except Exception as e:
        return None, f"FALLO - Error inesperado en PubMed Central para {doi} ({str(e)[:50]})"

def print_to_console(message, orig_stdout):
    print(message, file=orig_stdout)

def get_general_source_name(source_string):
    source_lower = source_string.lower()
    if "sci-hub" in source_lower:
        return "Sci-Hub"
    if "google scholar" in source_lower:
        return "Google Scholar"
    if "pmc" in source_lower or "pubmed" in source_lower:
        return "PubMed Central"

    # Fallback to extract domain
    try:
        domain_match = re.search(r'https?://(?:www\.)?([a-zA-Z0-9.-]+)(?:/|$)', source_string)
        if domain_match:
            return domain_match.group(1)
    except Exception:
        pass

    return "Otro"

def write_excel_report(excel_path, successful_data, failed_data, all_logs, original_columns, base_scihub_url, queue):
    if not excel_path:
        return

    try:
        ob_cols = ['DOI', 'Title', 'Successful_Mirror'] + [c for c in original_columns if c not in ['DOI', 'Title', 'Successful_Mirror']] + ['SciHub_Link']
        fa_cols = ['DOI', 'Title'] + [c for c in original_columns if c not in ['DOI', 'Title', 'Failure_Reason', 'Detailed_Status']] + ['Failure_Reason', 'Detailed_Status', 'SciHub_Link']
        ti_cols = ['DOI', 'Title'] + [c for c in original_columns if c not in ['DOI', 'Title', 'Successful_Mirror', 'Start_Time', 'End_Time', 'Duration_Seconds', 'Detailed_Status', 'Failure_Reason']] + ['Successful_Mirror', 'Start_Time', 'End_Time', 'Duration_Seconds', 'Detailed_Status', 'Failure_Reason', 'SciHub_Link']

        def create_ordered_df(data, cols):
            df_ = pd.DataFrame(data)
            df_['SciHub_Link'] = df_.apply(lambda r: f"{base_scihub_url}{r.get('DOI', r.get('doi', ''))}" if pd.notna(r.get('DOI', r.get('doi', ''))) else '', axis=1)
            for c in cols:
                if c not in df_.columns:
                    df_[c] = pd.NA
            all_data_keys = set()
            if data:
                all_data_keys.update(k for item in data for k in item.keys())
            final_cols_ordered = [c for c in cols if c in all_data_keys or c == 'SciHub_Link']
            final_cols_ordered.extend([k for k in all_data_keys if k not in final_cols_ordered and k != 'SciHub_Link'])
            return df_.reindex(columns=final_cols_ordered)

        df_obtenidos = create_ordered_df(successful_data, ob_cols)
        df_fallidos = create_ordered_df(failed_data, fa_cols)
        df_tiempos = create_ordered_df(all_logs, ti_cols)

        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            df_obtenidos.to_excel(writer, sheet_name='Obtenidos', index=False)
            df_fallidos.to_excel(writer, sheet_name='Fallidos', index=False)
            df_tiempos.to_excel(writer, sheet_name='Tiempos', index=False)
    except Exception as e:
        # Using print here because this function might be called where the queue is not available
        # or in a context where a GUI update is not the main point.
        print(f"Error guardando reporte Excel: {e}")
        if queue:
            queue.put({'type': 'error', 'message': f"Error guardando Excel: {e}"})


def download_pdfs_from_file(config, queue):
    driver = None
    original_stdout = sys.stdout
    temp_pdf_paths = []
    try:
        # --- WebDriver Initialization ---
        try:
            os.environ['WDM_LOG_LEVEL'] = '0'
            options = webdriver.ChromeOptions()
            options.add_argument('--headless')
            options.add_argument('--disable-gpu')
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--log-level=3')
            options.add_experimental_option('excludeSwitches', ['enable-logging'])
            # print("Inicializando WebDriver de Selenium en modo headless...") # Reduced noise
            driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
            # print("WebDriver de Selenium inicializado correctamente.") # Reduced noise
        except Exception as e:
            # print(f"Error al inicializar WebDriver de Selenium: {e}") # Reduced noise
            # print("Las descargas basadas en Selenium se omitirán.") # Reduced noise
            driver = None

        # --- Configuration Unpacking ---
        df = None
        input_file_path = config["input_file_path"]
        zip_path = config["zip_path"]
        excel_report_path_config = config["excel_report_path"]
        user_inter_doi_delay = config["user_inter_doi_delay"]
        user_mirror_switch_delay = config["user_mirror_switch_delay"]
        raw_mirrors_from_gui = config["user_defined_mirrors"]

        user_defined_mirrors = []
        for mirror_url in raw_mirrors_from_gui:
            if not mirror_url.startswith(("http://", "https://")):
                mirror_url = "https://" + mirror_url
            if not mirror_url.endswith('/'):
                mirror_url += '/'
            user_defined_mirrors.append(mirror_url)
        sci_hub_base_url_for_report = user_defined_mirrors[0] if user_defined_mirrors else "N/A"

        print("\n--- Configuración Aplicada ---")
        print(f"Archivo de entrada: {input_file_path}")
        print(f"Archivo ZIP de salida: {zip_path}")
        if excel_report_path_config: print(f"Archivo de reporte Excel: {excel_report_path_config}")
        else: print("Reporte Excel: No se generará (ruta no especificada).")
        print(f"Retraso Inter-DOI: {user_inter_doi_delay}s")
        print(f"Retraso Cambio de Mirror: {user_mirror_switch_delay}s")
        print(f"Mirrors Sci-Hub a utilizar: {', '.join(user_defined_mirrors)}")
        print("-----------------------------------------------------\n")

        # --- File Reading ---
        session = requests.Session()
        session.headers.update({'User-Agent': STANDARD_USER_AGENT})
        all_articles_log = []
        successful_articles_data = []
        failed_articles_data = []
        original_input_columns = []

        try:
            file_extension = os.path.splitext(input_file_path)[1].lower()
            if file_extension in ['.xlsx', '.xls']:
                df = pd.read_excel(input_file_path)
            elif file_extension == '.csv':
                df = pd.read_csv(input_file_path)
            else:
                raise ValueError(f"Tipo de archivo no soportado: {file_extension}")

            if df is not None:
                original_input_columns = [col for col in df.columns if col not in ['DOI', 'Title']]
            else:
                raise ValueError("DataFrame no fue cargado correctamente.")
        except Exception as e:
            print(f"Error fatal al leer archivo de entrada: {e}")
            queue.put({'type': 'error', 'message': f"Error al leer archivo: {e}"})
            return

        if df is None:
            print("Error Crítico: DataFrame (df) no fue inicializado.")
            queue.put({'type': 'error', 'message': "DataFrame no pudo ser cargado."})
            return

        # Initial Excel file creation
        if excel_report_path_config:
            write_excel_report(excel_report_path_config, [], [], [], original_input_columns, sci_hub_base_url_for_report, queue)

        # --- Main Processing Loop ---
        successful_downloads = 0
        total_downloaded_size_bytes = 0
        total_articles = len(df)
        source_stats = {}
        process_start_time = time.time()
        queue.put({'type': 'total', 'value': total_articles})

        try:
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for index, row in df.iterrows():
                    original_row_data = row.to_dict()
                    start_time = datetime.now()
                    doi = str(original_row_data.get('DOI', original_row_data.get('doi', ''))).strip()
                    title = str(original_row_data.get('Title', original_row_data.get('title', ''))).strip()

                    article_display_text = f"Título: {title}\nDOI: {doi}"
                    queue.put({'type': 'current_article', 'value': article_display_text})
                    effective_title = title if title else doi

                    current_article_num_for_log = index + 1
                    mirror_attempts_details_for_doi = []
                    overall_doi_status = "FALTANTE"

                    if not doi:
                        failure_reason_for_report = "DOI vacío"
                        detailed_status_for_log = "Skipped_DOI_Missing"
                        overall_doi_status = "FALTANTE (DOI Vacío)"
                        format_and_log_article_status(original_row_data, doi, effective_title, current_article_num_for_log, total_articles, successful_downloads, mirror_attempts_details_for_doi, overall_doi_status, user_inter_doi_delay, failed_articles_data_len=len(failed_articles_data))
                        end_time = datetime.now()
                        log_entry = {**original_row_data, 'Start_Time': start_time.strftime("%Y-%m-%d %H:%M:%S"), 'End_Time': end_time.strftime("%Y-%m-%d %H:%M:%S"), 'Duration_Seconds': (end_time - start_time).total_seconds(), 'Detailed_Status': detailed_status_for_log, 'Failure_Reason': failure_reason_for_report, 'Successful_Mirror': ""}
                        all_articles_log.append(log_entry)
                        failed_articles_data.append({**original_row_data, 'Failure_Reason': failure_reason_for_report, 'Detailed_Status': detailed_status_for_log, 'original_index': index})
                        print_to_console("===============================================================================================", original_stdout)
                        time.sleep(user_inter_doi_delay)
                        continue

                    pdf_filename_in_zip = clean_filename(effective_title)[:150] + ".pdf"
                    mirrors_to_try_for_this_doi = list(user_defined_mirrors)
                    pdf_content = None
                    download_successful_this_doi = False
                    successful_mirror_for_this_doi = ""
                    temp_detailed_status_for_log = ""
                    temp_failure_reason_for_log = ""

                    for mirror_idx, current_mirror_base_url in enumerate(mirrors_to_try_for_this_doi):
                        full_sci_hub_url_for_html_page = f"{current_mirror_base_url}{doi}"
                        mirror_status_str = "FALLO"
                        mirror_reason_str = ""

                        actual_pdf_download_url = extract_pdf_link_from_html(full_sci_hub_url_for_html_page, session)
                        if actual_pdf_download_url:
                            try:
                                response = session.get(actual_pdf_download_url, timeout=60)
                                response.raise_for_status()
                                content_type = response.headers.get('Content-Type', '').lower()
                                if 'application/pdf' in content_type:
                                    pdf_content = response.content
                                    mirror_status_str = "OBTENIDO (Extracción Iframe/Embed)"
                                    temp_detailed_status_for_log = f"Success_iframe_or_embed_extraction_from_{current_mirror_base_url}"
                                else:
                                    mirror_reason_str = f"Content-Type no PDF ({content_type})"
                                    temp_detailed_status_for_log = f"Failure_iframe_or_embed_extraction_not_pdf_from_{current_mirror_base_url}"
                            except requests.exceptions.HTTPError as e:
                                mirror_reason_str = f"HTTPError {e.response.status_code}"
                                temp_detailed_status_for_log = f"Failure_iframe_or_embed_extraction_HTTP{e.response.status_code}_from_{current_mirror_base_url}"
                            except requests.exceptions.RequestException as e:
                                mirror_reason_str = "Error de conexión/RequestException en extracción"
                                temp_detailed_status_for_log = f"Failure_iframe_or_embed_extraction_RequestException_from_{current_mirror_base_url}"
                            except Exception as e:
                                mirror_reason_str = "Error inesperado en extracción"
                                temp_detailed_status_for_log = f"Failure_iframe_or_embed_extraction_Unexpected_from_{current_mirror_base_url}"
                        else:
                            mirror_reason_str = "No se encontró enlace PDF en HTML"
                            temp_detailed_status_for_log = f"Failure_No_PDF_Link_Found_In_HTML_from_{current_mirror_base_url}"

                        if not pdf_content:
                            temp_failure_reason_for_log = mirror_reason_str
                            try:
                                response = session.get(full_sci_hub_url_for_html_page, timeout=30)
                                response.raise_for_status()
                                content_type = response.headers.get('Content-Type', '').lower()
                                if 'application/pdf' in content_type:
                                    pdf_content = response.content
                                    mirror_status_str = "OBTENIDO (Acceso Directo Fallback)"
                                    mirror_reason_str = ""
                                    temp_detailed_status_for_log = f"Success_direct_DOI_access_fallback_from_{current_mirror_base_url}"
                                else:
                                    mirror_reason_str = f"Content-Type no PDF ({content_type}) en acceso directo"
                                    temp_detailed_status_for_log = f"Failure_direct_DOI_access_not_pdf_from_{current_mirror_base_url}"
                            except requests.exceptions.HTTPError as e:
                                mirror_reason_str = f"HTTPError {e.response.status_code} en acceso directo"
                                temp_detailed_status_for_log = f"Failure_direct_DOI_access_HTTP{e.response.status_code}_from_{current_mirror_base_url}"
                            except requests.exceptions.RequestException as e:
                                mirror_reason_str = "Error de conexión/RequestException en fallback"
                                temp_detailed_status_for_log = f"Failure_direct_DOI_access_RequestException_from_{current_mirror_base_url}"
                            except Exception as e:
                                mirror_reason_str = "Error inesperado en fallback"
                                temp_detailed_status_for_log = f"Failure_direct_DOI_access_Unexpected_from_{current_mirror_base_url}"

                        specific_reason_for_temp_log = mirror_reason_str

                        log_display_reason_for_sci_hub_attempt = mirror_reason_str
                        if mirror_status_str == "FALLO":
                            log_display_reason_for_sci_hub_attempt = full_sci_hub_url_for_html_page
                            temp_failure_reason_for_log = specific_reason_for_temp_log
                        else:
                            temp_failure_reason_for_log = ""

                        mirror_attempts_details_for_doi.append((current_mirror_base_url, mirror_status_str, log_display_reason_for_sci_hub_attempt))

                        if pdf_content:
                            download_successful_this_doi = True
                            successful_mirror_for_this_doi = current_mirror_base_url
                            if mirror_status_str.startswith("OBTENIDO"):
                                overall_doi_status = mirror_status_str
                            else:
                                overall_doi_status = "OBTENIDO"
                            temp_failure_reason_for_log = ""
                            break
                        else:
                            if mirror_idx < len(mirrors_to_try_for_this_doi) - 1:
                                time.sleep(user_mirror_switch_delay)

                    if not pdf_content and driver:
                        # print(f"INFO: DOI {doi} - Attempting Google Scholar (Selenium method)...") # Reduced noise
                        gs_selenium_pdf_content, gs_selenium_status_msg = download_with_selenium_google_scholar(driver, doi, effective_title)
                        if gs_selenium_pdf_content:
                            pdf_content = gs_selenium_pdf_content
                            download_successful_this_doi = True
                            successful_mirror_for_this_doi = "Google Scholar (Selenium)"
                            overall_doi_status = gs_selenium_status_msg
                            mirror_attempts_details_for_doi.append(("Google Scholar (Selenium)", "OBTENIDO", gs_selenium_status_msg))
                            temp_detailed_status_for_log = f"Success_GoogleScholar_Selenium_{gs_selenium_status_msg}"
                            temp_failure_reason_for_log = ""
                        else:
                            mirror_attempts_details_for_doi.append(("Google Scholar (Selenium)", "FALLO", gs_selenium_status_msg))
                            temp_failure_reason_for_log = gs_selenium_status_msg
                            temp_detailed_status_for_log = f"Failure_GoogleScholar_Selenium_{gs_selenium_status_msg}"

                    if not pdf_content and driver:
                        # print(f"INFO: DOI {doi} - Attempting PubMed Central (Selenium method)...") # Reduced noise
                        pmc_selenium_pdf_content, pmc_selenium_status_msg = download_with_selenium_pmc(driver, doi, effective_title)
                        if pmc_selenium_pdf_content:
                            pdf_content = pmc_selenium_pdf_content
                            download_successful_this_doi = True
                            successful_mirror_for_this_doi = pmc_selenium_status_msg
                            overall_doi_status = pmc_selenium_status_msg
                            mirror_attempts_details_for_doi.append(("PubMed Central (Selenium)", "OBTENIDO", pmc_selenium_status_msg))
                            temp_detailed_status_for_log = f"Success_PubMedCentral_Selenium_{pmc_selenium_status_msg}"
                            temp_failure_reason_for_log = ""
                        else:
                            mirror_attempts_details_for_doi.append(("PubMed Central (Selenium)", "FALLO", pmc_selenium_status_msg))
                            temp_failure_reason_for_log = pmc_selenium_status_msg
                            temp_detailed_status_for_log = f"Failure_PubMedCentral_Selenium_{pmc_selenium_status_msg}"

                    end_time = datetime.now()
                    if download_successful_this_doi and pdf_content:
                        successful_downloads += 1

                        source_name = get_general_source_name(successful_mirror_for_this_doi)
                        source_stats[source_name] = source_stats.get(source_name, 0) + 1
                        queue.put({'type': 'source_stat', 'stats': source_stats})

                        if not overall_doi_status.startswith("OBTENIDO"):
                            overall_doi_status = "OBTENIDO"
                        data_for_successful_sheet = original_row_data.copy()
                        data_for_successful_sheet['Successful_Mirror'] = successful_mirror_for_this_doi
                        successful_articles_data.append(data_for_successful_sheet)
                        temp_dir = "temp_scihub_pdfs"
                        if not os.path.exists(temp_dir):
                            os.makedirs(temp_dir)
                        temp_pdf_path = os.path.join(temp_dir, f"temp_{os.getpid()}_{index}_{pdf_filename_in_zip}")
                        with open(temp_pdf_path, 'wb') as f:
                            f.write(pdf_content)
                        try:
                            total_downloaded_size_bytes += os.path.getsize(temp_pdf_path)
                        except OSError as e:
                            print(f"Advertencia: tamaño temp {temp_pdf_path}: {e}")
                        zf.write(temp_pdf_path, arcname=pdf_filename_in_zip)
                        temp_pdf_paths.append(temp_pdf_path)
                    else:
                        overall_doi_status = "FALTANTE"
                        failed_articles_data.append({**original_row_data, 'Failure_Reason': temp_failure_reason_for_log, 'Detailed_Status': temp_detailed_status_for_log, 'original_index': index})

                    format_and_log_article_status(original_row_data, doi, effective_title, current_article_num_for_log, total_articles, successful_downloads, mirror_attempts_details_for_doi, overall_doi_status, user_inter_doi_delay, failed_articles_data_len=len(failed_articles_data))

                    log_entry_failure_reason = temp_failure_reason_for_log if not download_successful_this_doi else ""
                    log_entry_detailed_status = temp_detailed_status_for_log if not download_successful_this_doi else f"Success_{successful_mirror_for_this_doi}"
                    all_articles_log.append({**original_row_data, 'Start_Time': start_time.strftime("%Y-%m-%d %H:%M:%S"), 'End_Time': end_time.strftime("%Y-%m-%d %H:%M:%S"), 'Duration_Seconds': (end_time - start_time).total_seconds(), 'Detailed_Status': log_entry_detailed_status, 'Failure_Reason': log_entry_failure_reason, 'Successful_Mirror': successful_mirror_for_this_doi})

                    queue.put({'type': 'progress', 'searched': index + 1, 'found': successful_downloads})

                    # --- ETA Calculation ---
                    articles_processed = index + 1
                    elapsed_time = time.time() - process_start_time
                    if articles_processed > 0:
                        avg_time_per_article = elapsed_time / articles_processed
                        articles_remaining = total_articles - articles_processed
                        estimated_remaining_time = articles_remaining * avg_time_per_article

                        if estimated_remaining_time > 0:
                            mins, secs = divmod(estimated_remaining_time, 60)
                            hours, mins = divmod(mins, 60)
                            eta_string = ""
                            if hours > 0:
                                eta_string += f"{int(hours)}h "
                            if mins > 0 or hours > 0:
                                eta_string += f"{int(mins)}m "
                            eta_string += f"{int(secs)}s"
                        else:
                            eta_string = "Completado"

                        queue.put({'type': 'time_update', 'value': eta_string})

                    print_to_console("===============================================================================================", original_stdout)

                    # Dynamic Excel Report Writing
                    write_excel_report(excel_report_path_config, successful_articles_data, failed_articles_data, all_articles_log, original_input_columns, sci_hub_base_url_for_report, queue)

                    if current_article_num_for_log < total_articles:
                        time.sleep(user_inter_doi_delay)
                    if download_successful_this_doi and "Google Scholar" in successful_mirror_for_this_doi:
                        log_entry_detailed_status = f"Success_{successful_mirror_for_this_doi}"

        except FileNotFoundError as e:
            print(f"Error crítico: No se pudo crear ZIP (Directorio no encontrado): {zip_path}")
            queue.put({'type': 'error', 'message': f"No se pudo crear ZIP: {e}"})
            return
        except Exception as e:
            print(f"Error crítico: Excepción en ZIP o descargas: {e}.")
            queue.put({'type': 'error', 'message': f"Error inesperado: {e}"})
            return

        # --- Retry Phase ---
        if failed_articles_data:
            print("\n--- Iniciando Fase de Reintento para Artículos Fallidos ---")
        articles_successfully_retried_ids = []
        temp_failed_articles_data_for_iteration = list(failed_articles_data)
        mirrors_for_retry = list(user_defined_mirrors)
        for retry_idx, failed_article_entry in enumerate(temp_failed_articles_data_for_iteration):
            original_index_for_retry = failed_article_entry.get('original_index', -1)
            current_article_num_for_log_retry = original_index_for_retry + 1 if original_index_for_retry != -1 else retry_idx + 1
            doi_to_retry = str(failed_article_entry.get('DOI', failed_article_entry.get('doi', ''))).strip()
            effective_title_for_retry = str(failed_article_entry.get('Title', failed_article_entry.get('title', doi_to_retry))).strip() or doi_to_retry
            pdf_filename_in_zip_retry = clean_filename(effective_title_for_retry)[:150] + ".pdf"
            mirror_attempts_details_for_retry = []
            overall_retry_status = "FALTANTE"
            pdf_content_retry = None; retry_successful_this_doi = False; successful_mirror_for_retry = ""
            temp_detailed_status_for_retry_log = ""; temp_failure_reason_for_retry_log = ""
            retry_start_time_actual_attempt = datetime.now()
            for mirror_idx_retry, current_mirror_base_url_retry in enumerate(mirrors_for_retry):
                full_sci_hub_url_for_html_page_retry = f"{current_mirror_base_url_retry}{doi_to_retry}"
                mirror_status_str_retry = "FALLO"; mirror_reason_str_retry = ""
                actual_pdf_download_url_retry = extract_pdf_link_from_html(full_sci_hub_url_for_html_page_retry, session)
                if actual_pdf_download_url_retry:
                    try:
                        response = session.get(actual_pdf_download_url_retry, timeout=60); response.raise_for_status()
                        content_type = response.headers.get('Content-Type', '').lower()
                        if 'application/pdf' in content_type: pdf_content_retry = response.content; mirror_status_str_retry = "OBTENIDO (REINTENTO Extracción)"; temp_detailed_status_for_retry_log = f"Success_RETRY_iframe_embed_from_{current_mirror_base_url_retry}"
                        else: mirror_reason_str_retry = f"RETRY: Content-Type not PDF ({content_type})"; temp_detailed_status_for_retry_log = f"Failure_RETRY_iframe_embed_not_pdf_from_{current_mirror_base_url_retry}"
                    except requests.exceptions.HTTPError as e: mirror_reason_str_retry = f"RETRY: HTTPError {e.response.status_code}"; temp_detailed_status_for_retry_log = f"Failure_RETRY_iframe_embed_HTTPError_from_{current_mirror_base_url_retry}"
                    except requests.exceptions.RequestException as e: mirror_reason_str_retry = "RETRY: Error de conexión/RequestException en extracción"; temp_detailed_status_for_retry_log = f"Failure_RETRY_iframe_embed_RequestException_from_{current_mirror_base_url_retry}"
                    except Exception as e: mirror_reason_str_retry = "RETRY: Error inesperado en extracción"; temp_detailed_status_for_retry_log = f"Failure_RETRY_iframe_embed_Unexpected_from_{current_mirror_base_url_retry}"
                else: mirror_reason_str_retry = "RETRY: No se encontró enlace PDF en HTML"; temp_detailed_status_for_retry_log = f"Failure_RETRY_No_PDF_Link_Found_In_HTML_from_{current_mirror_base_url_retry}"
                if not pdf_content_retry:
                    temp_failure_reason_for_retry_log = mirror_reason_str_retry
                    try:
                        response = session.get(full_sci_hub_url_for_html_page_retry, timeout=30); response.raise_for_status()
                        content_type = response.headers.get('Content-Type', '').lower()
                        if 'application/pdf' in content_type:
                            pdf_content_retry = response.content; mirror_status_str_retry = "OBTENIDO (REINTENTO Fallback Directo)"; mirror_reason_str_retry = ""; temp_detailed_status_for_retry_log = f"Success_RETRY_direct_DOI_from_{current_mirror_base_url_retry}"
                        else:
                            mirror_reason_str_retry = f"RETRY: Content-Type not PDF ({content_type}) en fallback"; temp_detailed_status_for_retry_log = f"Failure_RETRY_direct_DOI_not_pdf_from_{current_mirror_base_url_retry}"
                    except requests.exceptions.HTTPError as e: mirror_reason_str_retry = f"RETRY: HTTPError {e.response.status_code} en fallback"; temp_detailed_status_for_retry_log = f"Failure_RETRY_direct_DOI_HTTPError_from_{current_mirror_base_url_retry}"
                    except requests.exceptions.RequestException as e: mirror_reason_str_retry = "RETRY: Error de conexión/RequestException en fallback"; temp_detailed_status_for_retry_log = f"Failure_RETRY_direct_DOI_RequestException_from_{current_mirror_base_url_retry}"
                    except Exception as e: mirror_reason_str_retry = "RETRY: Error inesperado en fallback"; temp_detailed_status_for_retry_log = f"Failure_RETRY_direct_DOI_Unexpected_from_{current_mirror_base_url_retry}"
                specific_reason_for_temp_retry_log = mirror_reason_str_retry
                log_display_reason_for_sci_hub_retry_attempt = mirror_reason_str_retry
                if mirror_status_str_retry == "FALLO":
                    log_display_reason_for_sci_hub_retry_attempt = full_sci_hub_url_for_html_page_retry
                    temp_failure_reason_for_retry_log = specific_reason_for_temp_retry_log
                else:
                    temp_failure_reason_for_retry_log = ""
                mirror_attempts_details_for_retry.append((current_mirror_base_url_retry, mirror_status_str_retry, log_display_reason_for_sci_hub_retry_attempt))
                if pdf_content_retry:
                    retry_successful_this_doi = True
                    successful_mirror_for_retry = current_mirror_base_url_retry
                    if mirror_status_str_retry.startswith("OBTENIDO"): overall_retry_status = mirror_status_str_retry
                    else: overall_retry_status = "OBTENIDO"
                    temp_failure_reason_for_retry_log = ""
                    break
                else:
                    if mirror_idx_retry < len(mirrors_for_retry) - 1:
                        time.sleep(user_mirror_switch_delay)
            if not pdf_content_retry and driver:
                # print(f"INFO: DOI {doi_to_retry} [RETRY] - Attempting Google Scholar (Selenium method)...") # Reduced noise
                gs_selenium_pdf_content_retry, gs_selenium_status_msg_retry = download_with_selenium_google_scholar(driver, doi_to_retry, effective_title_for_retry)
                if gs_selenium_pdf_content_retry:
                    pdf_content_retry = gs_selenium_pdf_content_retry
                    retry_successful_this_doi = True
                    successful_mirror_for_retry = "Google Scholar (Selenium)"
                    overall_retry_status = gs_selenium_status_msg_retry
                    mirror_attempts_details_for_retry.append(("Google Scholar (Selenium Retry)", "OBTENIDO", gs_selenium_status_msg_retry))
                    temp_detailed_status_for_retry_log = f"Success_RETRY_GoogleScholar_Selenium_{gs_selenium_status_msg_retry}"
                    temp_failure_reason_for_retry_log = ""
                else:
                    mirror_attempts_details_for_retry.append(("Google Scholar (Selenium Retry)", "FALLO", gs_selenium_status_msg_retry))
                    temp_failure_reason_for_retry_log = gs_selenium_status_msg_retry
                    temp_detailed_status_for_retry_log = f"Failure_RETRY_GoogleScholar_Selenium_{gs_selenium_status_msg_retry}"
            if not pdf_content_retry and driver:
                # print(f"INFO: DOI {doi_to_retry} [RETRY] - Attempting PubMed Central (Selenium method)...") # Reduced noise
                pmc_selenium_pdf_content_retry, pmc_selenium_status_msg_retry = download_with_selenium_pmc(driver, doi_to_retry, effective_title_for_retry)
                if pmc_selenium_pdf_content_retry:
                    pdf_content_retry = pmc_selenium_pdf_content_retry
                    retry_successful_this_doi = True
                    successful_mirror_for_retry = pmc_selenium_status_msg_retry
                    overall_retry_status = pmc_selenium_status_msg_retry
                    mirror_attempts_details_for_retry.append(("PubMed Central (Selenium Retry)", "OBTENIDO", pmc_selenium_status_msg_retry))
                    temp_detailed_status_for_retry_log = f"Success_RETRY_PMC_Selenium_{pmc_selenium_status_msg_retry}"
                    temp_failure_reason_for_retry_log = ""
                else:
                    mirror_attempts_details_for_retry.append(("PubMed Central (Selenium Retry)", "FALLO", pmc_selenium_status_msg_retry))
                    temp_failure_reason_for_retry_log = pmc_selenium_status_msg_retry
                    temp_detailed_status_for_retry_log = f"Failure_RETRY_PubMedCentral_Selenium_{pmc_selenium_status_msg_retry}"
            retry_end_time_actual_attempt = datetime.now()
            original_article_log_entry = next((log for log in all_articles_log if str(log.get('DOI', log.get('doi', ''))).strip() == doi_to_retry), None)
            if retry_successful_this_doi and pdf_content_retry:
                successful_downloads += 1

                source_name_retry = get_general_source_name(successful_mirror_for_retry)
                source_stats[source_name_retry] = source_stats.get(source_name_retry, 0) + 1
                queue.put({'type': 'source_stat', 'stats': source_stats})

                queue.put({'type': 'progress', 'searched': total_articles, 'found': successful_downloads})
                if not overall_retry_status.startswith("OBTENIDO"):
                     overall_retry_status = "OBTENIDO"
                articles_successfully_retried_ids.append(doi_to_retry)
                original_data_for_success = {k: v for k, v in failed_article_entry.items() if k not in ['Failure_Reason', 'Detailed_Status', 'original_index']}
                original_data_for_success['Successful_Mirror'] = successful_mirror_for_retry
                successful_articles_data.append(original_data_for_success)
                temp_dir_retry = "temp_scihub_pdfs"
                if not os.path.exists(temp_dir_retry): os.makedirs(temp_dir_retry)
                temp_pdf_path_retry = os.path.join(temp_dir_retry, f"temp_RETRY_{os.getpid()}_{retry_idx}_{pdf_filename_in_zip_retry}")
                with open(temp_pdf_path_retry, 'wb') as f:
                    f.write(pdf_content_retry)
                try:
                    total_downloaded_size_bytes += os.path.getsize(temp_pdf_path_retry)
                except OSError as e:
                    print(f"Advertencia: tamaño temp (reintento) {temp_pdf_path_retry}: {e}")
                try:
                    with zipfile.ZipFile(zip_path, 'a', zipfile.ZIP_DEFLATED) as zf_append:
                        zf_append.write(temp_pdf_path_retry, arcname=pdf_filename_in_zip_retry)
                except Exception as e:
                    print(f"Error CRÍTICO al agregar PDF (reintento) '{pdf_filename_in_zip_retry}' al ZIP: {e}")
                temp_pdf_paths.append(temp_pdf_path_retry)
                if original_article_log_entry:
                    original_article_log_entry['Detailed_Status'] = temp_detailed_status_for_retry_log if not ("Google Scholar" in successful_mirror_for_retry) else f"Success_RETRY_GoogleScholar_{successful_mirror_for_retry}"
                    original_article_log_entry['Failure_Reason'] = "" if retry_successful_this_doi else temp_failure_reason_for_retry_log
                    original_article_log_entry['Successful_Mirror'] = successful_mirror_for_retry
                    original_article_log_entry['End_Time'] = retry_end_time_actual_attempt.strftime("%Y-%m-%d %H:%M:%S")
                    original_start_dt = datetime.strptime(original_article_log_entry['Start_Time'], "%Y-%m-%d %H:%M:%S")
                    original_article_log_entry['Duration_Seconds'] = (retry_end_time_actual_attempt - original_start_dt).total_seconds()
            else:
                if not overall_retry_status.startswith("FALLO"):
                    overall_retry_status = "FALTANTE"
                if original_article_log_entry:
                    original_article_log_entry['Detailed_Status'] = temp_detailed_status_for_retry_log
                    original_article_log_entry['Failure_Reason'] = temp_failure_reason_for_retry_log
                for item in failed_articles_data:
                    if str(item.get('DOI',item.get('doi',''))).strip() == doi_to_retry:
                        item['Failure_Reason'] = temp_failure_reason_for_retry_log
                        item['Detailed_Status'] = temp_detailed_status_for_retry_log
                        break
            format_and_log_article_status(failed_article_entry, doi_to_retry, effective_title_for_retry, current_article_num_for_log_retry, total_articles, successful_downloads, mirror_attempts_details_for_retry, overall_retry_status, user_inter_doi_delay, is_retry=True, failed_articles_data_len=len(temp_failed_articles_data_for_iteration) - retry_idx)

            # Update the Excel report after the retry attempt
            if articles_successfully_retried_ids:
                # This is slightly inefficient as it rebuilds the list every time, but it's the safest way to ensure correctness
                current_failed_articles_data = [item for item in temp_failed_articles_data_for_iteration if str(item.get('DOI', item.get('doi', ''))).strip() not in articles_successfully_retried_ids]
                write_excel_report(excel_report_path_config, successful_articles_data, current_failed_articles_data, all_articles_log, original_input_columns, sci_hub_base_url_for_report, queue)
            else:
                 write_excel_report(excel_report_path_config, successful_articles_data, failed_articles_data, all_articles_log, original_input_columns, sci_hub_base_url_for_report, queue)

            print_to_console("===============================================================================================", original_stdout)
            if retry_idx < len(temp_failed_articles_data_for_iteration) - 1:
                time.sleep(user_inter_doi_delay)

        if articles_successfully_retried_ids:
            failed_articles_data = [item for item in failed_articles_data if str(item.get('DOI', item.get('doi', ''))).strip() not in articles_successfully_retried_ids]

        failed_downloads_summary_list = [{'title': str(item.get('Title','N/A')).strip(), 'doi': str(item.get('DOI','N/A')).strip(), 'reason': str(item.get('Failure_Reason','N/A')).strip()} for item in failed_articles_data]
        total_mb = total_downloaded_size_bytes / (1024 * 1024)

        summary_message = (f"Proceso completado.\n\nDescargas exitosas: {successful_downloads}\nDescargas fallidas: {len(failed_downloads_summary_list)}\n" f"Tamaño total PDFs: {total_mb:.2f} MB")
        if failed_downloads_summary_list:
            summary_message += "\n\nArtículos no descargados (post-reintentos):"
            for item in failed_downloads_summary_list:
                summary_message += f"\n- Título: {item['title']}, DOI: {item['doi']}, Razón: {item['reason']}"

        print("\n" + "="*50); print(summary_message); print("="*50);

        # The final report is now written after every step, so this final block is redundant.
        # However, we can keep a final print message.
        if excel_report_path_config:
            print(f"Reporte Excel final guardado en: {excel_report_path_config}")
        else:
            print("Generación de reporte Excel omitida (ruta no especificada).")

        queue.put({'type': 'done', 'message': f'Completado: {successful_downloads}/{total_articles} artículos descargados.'})

        print("\n" + "="*50)
        print("--- Archivos Generados ---")
        if zip_path:
            print(f"Archivo ZIP de PDFs guardado en:")
            print(f"file:///{os.path.abspath(zip_path)}")

        if excel_report_path_to_use:
            print(f"\nReporte Excel guardado en:")
            print(f"file:///{os.path.abspath(excel_report_path_to_use)}")
        print("="*50 + "\n")

    finally: 
        # WebDriver will be quit here
        if driver:
            print("Cerrando WebDriver de Selenium...")
            try:
                driver.quit()
                print("WebDriver de Selenium cerrado correctamente.")
            except Exception as e:
                print(f"Error al cerrar WebDriver de Selenium: {e}")

        if hasattr(original_stdout, 'write'):
             if sys.stdout != original_stdout:
                 sys.stdout = original_stdout
                 print("\n--- stdout restaurado a la consola original durante la limpieza final ---", file=original_stdout)

        print("\n--- Limpieza Final de Archivos Temporales ---", file=original_stdout if hasattr(original_stdout, 'write') else sys.__stdout__)

        for temp_path in temp_pdf_paths:
            try:
                os.remove(temp_path)
                print(f"Eliminado temp: {temp_path}", file=original_stdout if hasattr(original_stdout, 'write') else sys.__stdout__)
            except OSError as e:
                print(f"Error eliminando temp {temp_path}: {e}", file=original_stdout if hasattr(original_stdout, 'write') else sys.__stdout__)

        temp_dir_to_check = "temp_scihub_pdfs"
        if os.path.exists(temp_dir_to_check) and not os.listdir(temp_dir_to_check):
            try:
                os.rmdir(temp_dir_to_check)
                print(f"Eliminado dir temp: {temp_dir_to_check}", file=original_stdout if hasattr(original_stdout, 'write') else sys.__stdout__)
            except OSError as e:
                print(f"Error eliminando dir temp {temp_dir_to_check}: {e}", file=original_stdout if hasattr(original_stdout, 'write') else sys.__stdout__)

        print("--- Limpieza Finalizada ---", file=original_stdout if hasattr(original_stdout, 'write') else sys.__stdout__)

def print_to_console(message, orig_stdout):
    print(message, file=orig_stdout)

import itertools

if __name__ == "__main__":
    root = tk.Tk()
    app = ConfigurationGUI(root)
    root.mainloop()

    print("\nScript finalizado.")
