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
                        # ... Sci-Hub logic ...
                        if pdf_content:
                            successful_source = "Sci-Hub"
                            break
                        if cancel_requested.is_set(): break

                if not pdf_content and use_scholar and driver:
                    # ... Google Scholar logic ...
                    if pdf_content: successful_source = "Google Scholar"

                if not pdf_content and use_pmc and driver:
                    # ... PMC logic ...
                    if pdf_content: successful_source = "PMC"

                if pdf_content:
                    successful_downloads += 1
                    # ... save file logic ...
                else:
                    failed_articles_data.append(original_row_data)

                queue.put({"type": "progress", "attempted": index + 1, "successful": successful_downloads, "total": total_articles})
                if index < total_articles - 1:
                    time.sleep(user_inter_doi_delay)

            # ... Retry logic ...

    except Exception as e:
        print(f"Error fatal en el hilo de trabajo: {e}")
    finally:
        if driver:
            driver.quit()
        # ... Temp file cleanup ...
        queue.put(None)

# --- GUI and File Reading ---
def read_dois_from_file(file_path):
    # ... (implementation is correct)
    pass
class TextRedirector:
    # ... (implementation is correct)
    pass
class ConfigFrame(tk.Frame):
    # ... (implementation is correct)
    pass
class MonitorFrame(tk.Frame):
    # ... (implementation is correct)
    pass
class SciHubDownloaderGUI:
    # ... (implementation is correct)
    pass
if __name__ == "__main__":
    root = tk.Tk()
    app = SciHubDownloaderGUI(root)
    root.mainloop()
