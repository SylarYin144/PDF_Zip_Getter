import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel, Frame, Label, Entry, Button, StringVar, IntVar, Checkbutton, scrolledtext
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
import threading

# --- Configuration Constants ---
DEFAULT_SCI_HUB_MIRRORS_EXAMPLE = ["https://sci-hub.se/", "https://sci-hub.st/", "https://sci-hub.box/", "https://sci-hub.ru/", "https://sci-hub.red/"]
INTER_DOI_DELAY_SECONDS = 5
MIRROR_SWITCH_DELAY_SECONDS = 3
STANDARD_USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.127 Safari/537.36'
# --- End Configuration Constants ---

class TextRedirector:
    def __init__(self, widget):
        self.widget = widget
        self.original_stdout = sys.stdout

    def write(self, str_):
        self.widget.configure(state='normal')
        self.widget.insert(tk.END, str_)
        self.widget.see(tk.END)
        self.widget.configure(state='disabled')

    def flush(self):
        self.widget.update_idletasks()

    def restore(self):
        sys.stdout = self.original_stdout

class Downloader:
    def __init__(self):
        self.pause_event = threading.Event()
        self.cancel_event = threading.Event()

    def get_pdf_content_via_js(self, driver, pdf_url):
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
                reader.readAsDataURL(blob);
            })
            .catch(error => callback({error: error.toString()}));
        """
        try:
            driver.set_script_timeout(90)
            result = driver.execute_async_script(script, pdf_url)
            return base64.b64decode(result) if not isinstance(result, dict) else None
        except Exception:
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
                return 'https:' + pdf_src if pdf_src.startswith('//') else urljoin(article_page_url, pdf_src)
            embed = soup.find('embed', attrs={'type': 'application/pdf'})
            if embed and embed.get('src'):
                return urljoin(article_page_url, embed['src'])
            return None
        except requests.exceptions.RequestException:
            return None

    def download_from_pmc(self, doi, session):
        print("-> Trying PMC (API)...")
        try:
            id_conv_url = f"https://www.ncbi.nlm.nih.gov/pmc/utils/idconv/v1.0/?ids={doi}&format=json"
            response_id_conv = session.get(id_conv_url, timeout=20)
            response_id_conv.raise_for_status()
            data_id_conv = response_id_conv.json()
            pmcid = data_id_conv.get("records", [{}])[0].get("pmcid")
            if not pmcid: return None, "FALLO - No PMCID found (API)"

            article_url = f"https://www.ncbi.nlm.nih.gov/pmc/articles/{pmcid}/"
            response_html = session.get(article_url, timeout=30)
            response_html.raise_for_status()
            soup = BeautifulSoup(response_html.content, 'html.parser')

            pdf_link_tag = soup.select_one("a.format-pdf, a.pdf-download, .format-menu a[href$='.pdf']")
            if pdf_link_tag and pdf_link_tag.get('href'):
                pdf_url = urljoin(article_url, pdf_link_tag['href'])
                pdf_response = session.get(pdf_url, timeout=60)
                pdf_response.raise_for_status()
                if 'application/pdf' in pdf_response.headers.get('Content-Type', ''):
                    return pdf_response.content, "OBTENIDO (PMC API)"
        except Exception as e:
            print(f"-> PMC (API) Error: {e}")
        return None, "FALLO - PMC (API)"

    def download_with_selenium_pmc(self, driver, doi, title):
        print("-> Trying PMC (Selenium)...")
        search_url = f"https://www.ncbi.nlm.nih.gov/pmc/?term={doi}"
        try:
            driver.get(search_url)
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.rprt .title a")))
            article_link = driver.find_element(By.CSS_SELECTOR, "div.rprt .title a").get_attribute('href')
            driver.get(article_link)
            pdf_link_element = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".format-menu a[href$='.pdf']")))
            pdf_url = pdf_link_element.get_attribute('href')
            pdf_content = self.get_pdf_content_via_js(driver, pdf_url)
            if pdf_content:
                return pdf_content, "OBTENIDO (PMC Selenium)"
        except Exception as e:
            print(f"-> PMC (Selenium) Error: {e}")
        return None, "FALLO - PMC (Selenium)"

    def download_with_selenium_google_scholar(self, driver, doi, title):
        print("-> Trying Google Scholar (Selenium)...")
        scholar_url = f"https://scholar.google.com/scholar?hl=en&q={doi}"
        try:
            driver.get(scholar_url)
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, "gs_res_ccl_mid")))

            pdf_links_elements = driver.find_elements(By.XPATH, "//a[contains(@href, '.pdf')]")
            if not pdf_links_elements:
                 pdf_links_elements = driver.find_elements(By.PARTIAL_LINK_TEXT, "[PDF]")

            for link_el in pdf_links_elements:
                pdf_url = link_el.get_attribute('href')
                if pdf_url:
                    pdf_content = self.get_pdf_content_via_js(driver, pdf_url)
                    if pdf_content:
                        return pdf_content, "OBTENIDO (GS Selenium)"
        except Exception as e:
            print(f"-> GS (Selenium) Error: {e}")
        return None, "FALLO - GS (Selenium)"

    def attempt_single_download(self, doi, title, mirrors, mirror_switch_delay, session, driver):
        # Attempt 1: Sci-Hub
        print("-> Trying Sci-Hub...")
        for mirror in mirrors:
            if self.cancel_event.is_set(): return None, "CANCELADO"
            try:
                pdf_url = self.extract_pdf_link_from_html(f"{mirror}{doi}", session)
                if pdf_url:
                    pdf_response = session.get(pdf_url, timeout=60)
                    if 'application/pdf' in pdf_response.headers.get('Content-Type', ''):
                        return pdf_response.content, f"OBTENIDO (Sci-Hub: {mirror.split('//')[1].split('/')[0]})"
            except requests.exceptions.RequestException:
                time.sleep(mirror_switch_delay)
                continue

        # Attempt 2: PMC API
        pdf_content, status = self.download_from_pmc(doi, session)
        if pdf_content:
            return pdf_content, status

        # Attempt 3: Google Scholar with Selenium
        if driver and not self.cancel_event.is_set():
            pdf_content, status = self.download_with_selenium_google_scholar(driver, doi, title)
            if pdf_content:
                return pdf_content, status

        # Attempt 4: PMC with Selenium
        if driver and not self.cancel_event.is_set():
            pdf_content, status = self.download_with_selenium_pmc(driver, doi, title)
            if pdf_content:
                return pdf_content, status

        return None, "FALLO"

    def run(self, input_file_path, zip_path, excel_report_path, inter_doi_delay, mirror_switch_delay, mirrors):
        driver = None
        try:
            options = webdriver.ChromeOptions()
            options.add_argument('--headless')
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--disable-gpu')
            options.add_argument('--disable-extensions')
            options.add_argument('--disable-popup-blocking')
            options.add_argument('--start-maximized')
            options.add_argument('--disable-infobars')
            options.add_argument('--ignore-certificate-errors')
            driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
            print("WebDriver de Selenium inicializado.")
        except Exception as e:
            print(f"Error al inicializar WebDriver: {e}. Se omitirán descargas con Selenium.")

        temp_pdf_paths = []
        successful_articles_data = []
        failed_articles_data = []

        try:
            session = requests.Session()
            session.headers.update({'User-Agent': STANDARD_USER_AGENT})

            try:
                df = pd.read_excel(input_file_path) if input_file_path.endswith(('.xlsx', '.xls')) else pd.read_csv(input_file_path)
            except Exception as e:
                print(f"Error al leer archivo de entrada: {e}")
                return

            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                total_articles = len(df)
                print(f"--- Proceso de Descarga Iniciado: {total_articles} artículos ---")

                # Main download loop
                for index, row in df.iterrows():
                    if self.cancel_event.is_set(): break
                    while self.pause_event.is_set(): time.sleep(1)

                    doi = str(row.get('DOI', '')).strip()
                    title = str(row.get('Title', '')).strip() or doi
                    if not doi: continue

                    print(f"\n({index+1}/{total_articles}) Procesando: {title}")
                    pdf_content, status = self.attempt_single_download(doi, title, mirrors, mirror_switch_delay, session, driver)

                    if pdf_content:
                        filename = self.clean_filename(title)[:150] + ".pdf"
                        temp_path = os.path.join("temp_pdfs", filename)
                        os.makedirs("temp_pdfs", exist_ok=True)
                        with open(temp_path, 'wb') as f: f.write(pdf_content)
                        zf.write(temp_path, arcname=filename)
                        temp_pdf_paths.append(temp_path)
                        print(f"-> ÉXITO: {status}")
                        successful_articles_data.append(row.to_dict())
                    else:
                        print(f"-> FALLO")
                        failed_articles_data.append(row.to_dict())

                    if index < total_articles - 1 and not self.cancel_event.is_set():
                        time.sleep(inter_doi_delay)

                # Retry loop
                if failed_articles_data and not self.cancel_event.is_set():
                    print(f"\n--- Reintentando {len(failed_articles_data)} artículos fallidos ---")
                    still_failed_data = []
                    for item_dict in failed_articles_data:
                        if self.cancel_event.is_set(): break
                        while self.pause_event.is_set(): time.sleep(1)

                        doi = str(item_dict.get('DOI', '')).strip()
                        title = str(item_dict.get('Title', '')).strip() or doi
                        print(f"\nReintentando: {title}")

                        pdf_content, status = self.attempt_single_download(doi, title, mirrors, mirror_switch_delay, session, driver)

                        if pdf_content:
                            filename = self.clean_filename(title)[:150] + ".pdf"
                            temp_path = os.path.join("temp_pdfs", filename)
                            with open(temp_path, 'wb') as f: f.write(pdf_content)
                            zf.write(temp_path, arcname=filename)
                            temp_pdf_paths.append(temp_path)
                            print(f"-> ÉXITO (en reintento): {status}")
                            successful_articles_data.append(item_dict)
                        else:
                            print(f"-> FALLO (en reintento)")
                            still_failed_data.append(item_dict)

                        if not self.cancel_event.is_set():
                            time.sleep(inter_doi_delay)

                    failed_articles_data = still_failed_data

            if excel_report_path:
                try:
                    with pd.ExcelWriter(excel_report_path, engine='openpyxl') as writer:
                        pd.DataFrame(successful_articles_data).to_excel(writer, sheet_name='Obtenidos', index=False)
                        pd.DataFrame(failed_articles_data).to_excel(writer, sheet_name='Fallidos', index=False)
                    print(f"\nReporte Excel guardado en: {excel_report_path}")
                except Exception as e:
                    print(f"\nError al guardar el reporte Excel: {e}")

            print("\n--- Proceso Finalizado ---")
        finally:
            if driver:
                driver.quit()
            for path in temp_pdf_paths:
                try: os.remove(path)
                except OSError: pass
            if os.path.exists("temp_pdfs") and not os.listdir("temp_pdfs"):
                try: os.rmdir("temp_pdfs")
                except OSError: pass

class SciHubDownloaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sci-Hub PDF Downloader")
        self.downloader = None
        self.thread = None
        self.redirector = None

        config_frame = Frame(self.root, padx=10, pady=10)
        config_frame.pack(fill="x")

        self.input_path = StringVar()
        self.output_path = StringVar()
        self.report_path = StringVar()
        self.inter_doi_delay = IntVar(value=INTER_DOI_DELAY_SECONDS)
        self.mirror_switch_delay = IntVar(value=MIRROR_SWITCH_DELAY_SECONDS)
        self.mirrors_str = StringVar(value=", ".join(DEFAULT_SCI_HUB_MIRRORS_EXAMPLE))

        Label(config_frame, text="Input File:").grid(row=0, column=0, sticky="w")
        Entry(config_frame, textvariable=self.input_path, width=50).grid(row=0, column=1)
        Button(config_frame, text="Browse...", command=lambda: self.browse_file(self.input_path, True)).grid(row=0, column=2)

        Label(config_frame, text="Output ZIP:").grid(row=1, column=0, sticky="w")
        Entry(config_frame, textvariable=self.output_path, width=50).grid(row=1, column=1)
        Button(config_frame, text="Browse...", command=lambda: self.browse_file(self.output_path, False)).grid(row=1, column=2)

        Label(config_frame, text="Excel Report (Optional):").grid(row=2, column=0, sticky="w")
        Entry(config_frame, textvariable=self.report_path, width=50).grid(row=2, column=1)
        Button(config_frame, text="Browse...", command=lambda: self.browse_file(self.report_path, False, is_report=True)).grid(row=2, column=2)

        Label(config_frame, text="Delay DOI (s):").grid(row=3, column=0, sticky="w")
        Entry(config_frame, textvariable=self.inter_doi_delay, width=10).grid(row=3, column=1, sticky="w")

        Label(config_frame, text="Delay Mirror (s):").grid(row=4, column=0, sticky="w")
        Entry(config_frame, textvariable=self.mirror_switch_delay, width=10).grid(row=4, column=1, sticky="w")

        Label(config_frame, text="Mirrors:").grid(row=5, column=0, sticky="w")
        Entry(config_frame, textvariable=self.mirrors_str, width=60).grid(row=5, column=1, columnspan=2, sticky="we")

        control_frame = Frame(self.root, padx=10, pady=10)
        control_frame.pack(fill="x")
        self.start_button = Button(control_frame, text="Start", command=self.start_download)
        self.start_button.pack(side="left")
        self.pause_button = Button(control_frame, text="Pause", command=self.pause_download, state="disabled")
        self.pause_button.pack(side="left")
        self.resume_button = Button(control_frame, text="Resume", command=self.resume_download, state="disabled")
        self.resume_button.pack(side="left")
        self.cancel_button = Button(control_frame, text="Cancel", command=self.cancel_download, state="disabled")
        self.cancel_button.pack(side="left")

        log_frame = Frame(self.root, padx=10, pady=10)
        log_frame.pack(fill="both", expand=True)
        self.log_widget = scrolledtext.ScrolledText(log_frame, state="disabled", wrap="word")
        self.log_widget.pack(fill="both", expand=True)

    def browse_file(self, path_var, is_open, is_report=False):
        if is_open:
            path = filedialog.askopenfilename(filetypes=(("Excel/CSV", "*.xlsx;*.xls;*.csv"), ("All files", "*.*")))
        else:
            if is_report:
                path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=(("Excel Report", "*.xlsx"),))
            else:
                path = filedialog.asksaveasfilename(defaultextension=".zip", filetypes=(("ZIP", "*.zip"),))
        if path:
            path_var.set(path)

    def start_download(self):
        if not self.input_path.get() or not self.output_path.get():
            messagebox.showerror("Error", "Input and Output paths are required.")
            return

        self.downloader = Downloader()
        mirrors = [m.strip() for m in self.mirrors_str.get().split(',') if m.strip()]

        self.thread = threading.Thread(target=self.downloader.run, args=(
            self.input_path.get(), self.output_path.get(), self.report_path.get(),
            self.inter_doi_delay.get(), self.mirror_switch_delay.get(), mirrors
        ), daemon=True)

        self.redirector = TextRedirector(self.log_widget)
        sys.stdout = self.redirector

        self.thread.start()
        self.update_ui_for_running(True)
        self.check_thread()

    def check_thread(self):
        if self.thread.is_alive():
            self.root.after(100, self.check_thread)
        else:
            self.update_ui_for_running(False)
            if self.redirector:
                self.redirector.restore()
            messagebox.showinfo("Complete", "Download process finished.")

    def update_ui_for_running(self, is_running):
        state = "disabled" if is_running else "normal"
        self.start_button.config(state=state)
        # Disable all config widgets when running
        for child in self.root.winfo_children():
            if isinstance(child, Frame):
                for widget in child.winfo_children():
                    if isinstance(widget, (Entry, Button)) and widget not in [self.pause_button, self.resume_button, self.cancel_button, self.start_button]:
                        widget.config(state=state)

        self.pause_button.config(state="normal" if is_running else "disabled")
        self.cancel_button.config(state="normal" if is_running else "disabled")
        self.resume_button.config(state="disabled")

    def pause_download(self):
        if self.downloader:
            self.downloader.pause_event.set()
            self.pause_button.config(state="disabled")
            self.resume_button.config(state="normal")

    def resume_download(self):
        if self.downloader:
            self.downloader.pause_event.clear()
            self.resume_button.config(state="disabled")
            self.pause_button.config(state="normal")

    def cancel_download(self):
        if self.downloader:
            self.downloader.cancel_event.set()

if __name__ == "__main__":
    root = tk.Tk()
    app = SciHubDownloaderApp(root)
    root.mainloop()
