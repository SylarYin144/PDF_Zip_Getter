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
from collections import Counter
import base64
import itertools

DEFAULT_SCI_HUB_MIRRORS_EXAMPLE = ["https://sci-hub.se/", "https://sci-hub.st/", "https://sci-hub.box/", "https://sci-hub.ru/", "https://sci-hub.red/"]
STANDARD_USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.127 Safari/537.36'

def log_helper(queue, message):
    if queue:
        queue.put({'type': 'log', 'message': message})
    else:
        print(message)

def get_pdf_content_via_js(driver, pdf_url, log_queue):
    script = """
    const callback = arguments[arguments.length - 1];
    fetch(arguments[0])
        .then(response => { if (!response.ok) { throw new Error('Network response was not ok: ' + response.status + ' ' + response.statusText); } return response.blob(); })
        .then(blob => {
            if (blob.type !== 'application/pdf' && !arguments[0].toLowerCase().endsWith('.pdf') && blob.type !== 'application/octet-stream') {
                     throw new Error('Content-Type is not application/pdf or octet-stream: ' + blob.type);
            }
            const reader = new FileReader();
            reader.onloadend = () => {
                const base64Marker = ';base64,';
                const base64Data = reader.result.substring(reader.result.indexOf(base64Marker) + base64Marker.length);
                callback(base64Data);
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
            log_helper(log_queue, f"JS Fetch Helper: Error from JS: {result['error']}")
            return None
        return base64.b64decode(result) if result else None
    except Exception as e:
        log_helper(log_queue, f"JS Fetch Helper: Exception: {e}")
        return None

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
            return urljoin(article_page_url, embed['src'])
        return None
    except: return None

def download_with_selenium_google_scholar(driver, doi, title, log_queue, config):
    log_helper(log_queue, f"SELENIUM GS: Buscando DOI: {doi}")
    scholar_url = f"https://scholar.google.com/scholar?hl=en&q={doi}"
    try:
        driver.get(scholar_url)
        wait = WebDriverWait(driver, config.get('element_wait_timeout', 20))
        pdf_links = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//a[contains(translate(., 'PDF', 'pdf'), 'pdf') or contains(@href, '.pdf')]")))
        for link_element in pdf_links:
            pdf_url = link_element.get_attribute('href')
            if pdf_url and 'scholar.google.com' not in pdf_url:
                log_helper(log_queue, f"SELENIUM GS: Encontrado posible PDF en: {pdf_url}")
                pdf_content = get_pdf_content_via_js(driver, pdf_url, log_queue)
                if pdf_content: return pdf_content, f"Obtenido de Google Scholar"
        return None, "No se encontró un enlace PDF válido en Google Scholar."
    except TimeoutException: return None, "Timeout esperando enlaces PDF en Google Scholar."
    except Exception as e: return None, f"Error en Selenium (GS): {e}"

def download_with_selenium_pmc(driver, doi, title, log_queue, config):
    log_helper(log_queue, f"SELENIUM PMC: Buscando DOI: {doi}")
    search_url = f"https://www.ncbi.nlm.nih.gov/pmc/?term={doi}"
    try:
        driver.get(search_url)
        wait = WebDriverWait(driver, config.get('element_wait_timeout', 20))
        article_links = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.rprt .title a")))
        if not article_links: return None, "No se encontró el artículo en PMC."
        article_url = article_links[0].get_attribute('href')
        driver.get(article_url)
        pdf_links = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf') and contains(., 'PDF')]")))
        if not pdf_links: return None, "No se encontró enlace PDF en la página del artículo."
        pdf_url = pdf_links[0].get_attribute('href')
        log_helper(log_queue, f"SELENIUM PMC: Encontrado posible PDF en: {pdf_url}")
        pdf_content = get_pdf_content_via_js(driver, pdf_url, log_queue)
        if pdf_content: return pdf_content, f"Obtenido de PubMed Central"
        return None, "No se pudo descargar el PDF desde el enlace de PMC."
    except TimeoutException: return None, "Timeout esperando elementos en la página de PMC."
    except Exception as e: return None, f"Error en Selenium (PMC): {e}"

def download_pdfs_from_file(config, progress_queue=None, cancel_event=None, pause_event=None):
    def log(message):
        if progress_queue: progress_queue.put({'type': 'log', 'message': message})
        else: print(message)
    df = config.get('df'); zip_path = config.get('zip_path'); excel_report_path = config.get('excel_report_path')
    user_inter_doi_delay = config.get('inter_doi_delay', 5); user_mirror_switch_delay = config.get('mirror_switch_delay', 3)
    user_defined_mirrors = config.get('sci_hub_mirrors', []); use_selenium_gs = config.get('use_google_scholar', False)
    use_selenium_pmc = config.get('use_pmc', False); page_load_timeout = config.get('page_load_timeout', 60)
    element_wait_timeout = config.get('element_wait_timeout', 20)
    driver = None
    if use_selenium_gs or use_selenium_pmc:
        try:
            log("Inicializando WebDriver..."); options = webdriver.ChromeOptions(); options.add_argument('--headless'); options.add_argument('--disable-gpu'); options.add_argument('--no-sandbox'); options.add_argument('--disable-dev-shm-usage')
            driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options); driver.set_page_load_timeout(page_load_timeout); log("WebDriver inicializado.")
        except Exception as e: log(f"Error al inicializar WebDriver: {e}"); driver = None
    session = requests.Session(); session.headers.update({'User-Agent': STANDARD_USER_AGENT})
    successful_articles = []; failed_articles = []; source_counts = Counter()
    if excel_report_path:
        try:
            with pd.ExcelWriter(excel_report_path, engine='openpyxl') as writer:
                df_copy = df.copy(); df_copy['Failure_Reason'] = 'En cola'
                df_copy.to_excel(writer, sheet_name='Fallidos', index=False)
                pd.DataFrame(columns=list(df.columns) + ['Successful_Mirror']).to_excel(writer, sheet_name='Obtenidos', index=False)
        except Exception as e: log(f"No se pudo crear el reporte Excel inicial: {e}"); excel_report_path = None
    try:
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            total_articles = len(df)
            for index, row in df.iterrows():
                if cancel_event and cancel_event.is_set(): log("Proceso cancelado."); break
                while pause_event and pause_event.is_set(): time.sleep(1)
                original_row_data = row.to_dict(); doi = str(original_row_data.get('doi', '')).strip(); title = str(original_row_data.get('title', '')); effective_title = title if title else doi
                if progress_queue: progress_queue.put({'type': 'current_article', 'data': {**original_row_data, 'current': index + 1, 'total': total_articles}})
                if not doi:
                    log(f"DOI vacío en la fila {index+1}, saltando."); failed_articles.append({'data': original_row_data, 'reason': 'DOI vacío'})
                else:
                    pdf_content = None; download_source = None; failure_reason = "No encontrado en ninguna fuente"
                    if user_defined_mirrors:
                        for mirror in user_defined_mirrors:
                            if cancel_event and cancel_event.is_set(): break
                            log(f"Intentando con mirror de Sci-Hub: {mirror}")
                            pdf_link = extract_pdf_link_from_html(f"{mirror}{doi}", session)
                            if pdf_link:
                                try:
                                    pdf_response = session.get(pdf_link, timeout=60)
                                    if 'application/pdf' in pdf_response.headers.get('Content-Type',''): pdf_content = pdf_response.content; download_source = "Sci-Hub"; break
                                except Exception as e: log(f"Error descargando desde {pdf_link}: {e}")
                            time.sleep(user_mirror_switch_delay)
                    if not pdf_content and use_selenium_gs and driver and not (cancel_event and cancel_event.is_set()):
                        pdf_content, failure_reason = download_with_selenium_google_scholar(driver, doi, effective_title, progress_queue, config); download_source = "Google Scholar" if pdf_content else None
                    if not pdf_content and use_selenium_pmc and driver and not (cancel_event and cancel_event.is_set()):
                        pdf_content, failure_reason = download_with_selenium_pmc(driver, doi, effective_title, progress_queue, config); download_source = "PubMed Central" if pdf_content else None
                    if pdf_content:
                        successful_articles.append({'data': original_row_data, 'source': download_source}); source_counts[download_source] += 1
                        pdf_filename = clean_filename(effective_title)[:150] + ".pdf"; zf.writestr(pdf_filename, pdf_content)
                    else:
                        failed_articles.append({'data': original_row_data, 'reason': failure_reason})
                if progress_queue:
                    progress_queue.put({'type': 'kpi', 'obtained': len(successful_articles), 'failed': len(failed_articles), 'pending': total_articles - (index + 1), 'source_counts': source_counts})
                    progress_queue.put({'type': 'article_result', 'doi': doi, 'success': bool(pdf_content), 'source': download_source, 'reason': failure_reason})
                time.sleep(user_inter_doi_delay)
    except Exception as e: log(f"Error crítico: {e}")
    finally:
        if driver: driver.quit()
    summary = {
        'successful_count': len(successful_articles),
        'failed_count': len(failed_articles),
        'successful_articles': successful_articles,
        'failed_articles': failed_articles,
        'total_articles': len(df),
        'source_counts': source_counts,
        'was_cancelled': cancel_event.is_set() if cancel_event else False,
        'excel_report_path': excel_report_path,
        'zip_path': zip_path
    }
    if progress_queue: progress_queue.put({'type': 'finished', 'summary': summary})
    return summary

if __name__ == '__main__':
    root = tk.Tk(); root.withdraw()
    input_file = filedialog.askopenfilename(title="Seleccionar archivo con DOIs")
    if not input_file: sys.exit()
    df = pd.read_excel(input_file) if input_file.endswith('.xlsx') else pd.read_csv(input_file)
    zip_path = filedialog.asksaveasfilename(title="Guardar ZIP como", defaultextension=".zip")
    if not zip_path: sys.exit()
    excel_path = filedialog.asksaveasfilename(title="Guardar Reporte como", defaultextension=".xlsx")
    config = {'df': df, 'zip_path': zip_path, 'excel_report_path': excel_path, 'sci_hub_mirrors': DEFAULT_SCI_HUB_MIRRORS_EXAMPLE, 'use_google_scholar': True, 'use_pmc': True}
    def cli_callback(data): print(data)
    download_pdfs_from_file(config, progress_queue=None, cancel_event=None, pause_event=None)
