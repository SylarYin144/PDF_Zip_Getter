import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
# import tkinter.scrolledtext as st # GUI logging disabled
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

# --- Configuration Constants (Primarily for defaults now) ---
DEFAULT_SCI_HUB_MIRRORS_EXAMPLE = [ 
    "https://sci-hub.se/", "https://sci-hub.st/", "https://sci-hub.ru/",
    "https://sci-hub.red/", "https://NFTsci-hub.box/", "https://sci-hub.wf/", 
    "https://sci-hub.cat/"
]
INTER_DOI_DELAY_SECONDS = 5 
MIRROR_SWITCH_DELAY_SECONDS = 3
# --- End Configuration Constants ---

# class TextRedirector(object): # GUI logging disabled
#     def __init__(self, widget, original_stdout_ref, tag="stdout"):
#         self.widget = widget
#         self.original_stdout = original_stdout_ref
#         self.tag = tag
#     def write(self, str_):
#         self.widget.configure(state='normal')
#         self.widget.insert(tk.END, str_, (self.tag,))
#         self.widget.see(tk.END)
#         self.widget.configure(state='disabled')
#         self.original_stdout.write(str_)
#         self.original_stdout.flush()
#     def flush(self):
#         self.widget.update_idletasks()
#         self.original_stdout.flush()

# def on_log_window_close(log_window_ref, original_stdout_ref, root_tk_instance): # GUI logging disabled
#     if sys.stdout != original_stdout_ref:
#         print("Log window closed by user. Restoring original stdout.", file=original_stdout_ref)
#         sys.stdout = original_stdout_ref
#     if log_window_ref:
#         try: log_window_ref.destroy()
#         except tk.TclError: pass

def format_and_log_article_status(original_row_data, doi, title, current_article_num, total_articles, 
                                  successful_downloads_count_for_stats, # Count *after* this article's attempt
                                  mirror_attempts_details, 
                                  overall_doi_status, current_user_inter_doi_delay, is_retry=False):
    try:
        buscados_percentage = (current_article_num / total_articles) * 100 if total_articles > 0 else 0
        obtenidos_percentage = (successful_downloads_count_for_stats / total_articles) * 100 if total_articles > 0 else 0
        
        log_lines = []
        retry_prefix = "[REINTENTO] " if is_retry else ""
        prefix_applied_in_this_function = False

        # log_lines.append(f"{retry_prefix}Artículo: {current_article_num}/{total_articles} ({buscados_percentage:.2f}%)")
        # log_lines.append(f"Título: {title if title else 'N/A'}")

        # # Retrieve First Author information
        # author_val = original_row_data.get('First Author', 'N/A')
        # if author_val == 'N/A': # Fallback to "Autores" (case-insensitive)
        #     autores_keys = [k for k in original_row_data.keys() if str(k).lower() == 'autores']
        #     author_val = original_row_data.get(autores_keys[0], 'N/A') if autores_keys else 'N/A'
        # if author_val == 'N/A': # Fallback to "Authors" (case-insensitive)
        #     authors_keys_en = [k for k in original_row_data.keys() if str(k).lower() == 'authors']
        #     author_val = original_row_data.get(authors_keys_en[0], 'N/A') if authors_keys_en else 'N/A'
        # log_lines.append(f"First Author: {author_val}")

        # # Prioritize "Journal/Book" for journal title
        # journal_title_val = original_row_data.get('Journal/Book', 'N/A')

        # # If "Journal/Book" is not found (or has no value and resulted in 'N/A'),
        # # then try the old 'revista' logic as a fallback.
        # if journal_title_val == 'N/A':
        #     # Case-insensitive search for 'revista'
        #     revista_keys = [k for k in original_row_data.keys() if str(k).lower() == 'revista']
        #     journal_title_val = original_row_data.get(revista_keys[0], 'N/A') if revista_keys else 'N/A'
        # log_lines.append(f"Journal/Book: {journal_title_val}")

        # # Retrieve Publication Year information
        # pub_year_val = original_row_data.get('Publication Year', 'N/A')
        # if pub_year_val == 'N/A': # Fallback to "Fecha de publicación" (case-insensitive)
        #     fecha_pub_keys = [k for k in original_row_data.keys() if str(k).lower() == 'fecha de publicación']
        #     pub_year_val = original_row_data.get(fecha_pub_keys[0], 'N/A') if fecha_pub_keys else 'N/A'
        # if pub_year_val == 'N/A': # Fallback to "Year" (case-insensitive)
        #     year_keys = [k for k in original_row_data.keys() if str(k).lower() == 'year']
        #     pub_year_val = original_row_data.get(year_keys[0], 'N/A') if year_keys else 'N/A'
        # if pub_year_val == 'N/A': # Fallback to "Año" (case-insensitive)
        #     ano_keys = [k for k in original_row_data.keys() if str(k).lower() == 'año']
        #     pub_year_val = original_row_data.get(ano_keys[0], 'N/A') if ano_keys else 'N/A'
        # log_lines.append(f"Publication Year: {pub_year_val}")
        
        # log_lines.append(f"DOI: {doi}")

        for i, attempt in enumerate(mirror_attempts_details):
            line_prefix = ""
            if is_retry and not prefix_applied_in_this_function:
                line_prefix = retry_prefix
                prefix_applied_in_this_function = True

            mirror_url, status, reason = attempt
            try:
                domain_match = re.search(r'https?://(?:www\.)?([a-zA-Z0-9.-]+)(?:/|$)', mirror_url)
                mirror_short_name = domain_match.group(1) if domain_match else mirror_url[-20:]
            except Exception: mirror_short_name = mirror_url[-20:] 
            log_lines.append(f"{line_prefix}Intento Mirror {i+1} ({mirror_short_name}): {status}. {reason if reason else ''}".strip())

        final_doi_status_prefix = ""
        if is_retry and not prefix_applied_in_this_function:
            final_doi_status_prefix = retry_prefix
            prefix_applied_in_this_function = True
        log_lines.append(f"{final_doi_status_prefix}Artículo {overall_doi_status}")

        total_downloaded_prefix = ""
        if is_retry and not prefix_applied_in_this_function:
            total_downloaded_prefix = retry_prefix
            # prefix_applied_in_this_function = True # Not strictly needed to set here as it's the last use
        log_lines.append(f"{total_downloaded_prefix}Total Descargados (actualizado): {successful_downloads_count_for_stats}/{total_articles} ({obtenidos_percentage:.2f}%)")
        
        formatted_message = "\n".join(log_lines)
        print(f"\n{formatted_message}") 
        
        # Only print waiting message if it's not the very last article overall OR if it failed (implying more retries or end)
        # This avoids "Waiting..." after the final successful article of the main loop if no retries follow.
        is_last_article_overall = (current_article_num == total_articles and not is_retry and not failed_articles_data) # Heuristic
        
        if not (overall_doi_status == "OBTENIDO" and is_last_article_overall and not is_retry) : # Check if not last successful article
             # Check if it's not the last item in its current loop (main or retry)
            if not is_retry and current_article_num < total_articles : # Main loop, not last
                 print(f"Esperando {current_user_inter_doi_delay} segundos…\n")
            elif is_retry : # Always print for retries unless it's the very very last action. This is tricky.
                 # Let's say, if there are more retries OR if this isn't the absolute last processed item.
                 # This part of the logic for waiting message might need refinement based on exact desired flow.
                 # For now, print it if it's a retry and not the last one in the retry batch.
                 # The caller of format_and_log_article_status will handle the actual time.sleep()
                 print(f"Esperando {current_user_inter_doi_delay} segundos (post-reintento)...\n")


    except Exception as e:
        print(f"\nError al formatear log para DOI {doi}: {e}")
        print(f"Fallback Log: DOI: {doi}, Status: {overall_doi_status}\n")


def clean_filename(title):
    return re.sub(r'[\\/*?:"<>|]', '_', title)

def extract_pdf_link_from_html(article_page_url, session):
    # print(f"Extrayendo HTML de: {article_page_url}")
    try:
        response = session.get(article_page_url, timeout=30); response.raise_for_status()
        soup = BeautifulSoup(response.content, 'html.parser')
        iframe = soup.find('iframe', id='pdf')
        if iframe and iframe.get('src'):
            pdf_src = iframe['src'] #; print(f"Encontrado PDF en iframe: {pdf_src}")
            if pdf_src.startswith('//'): return 'https:' + pdf_src
            elif pdf_src.startswith('/'): return urljoin(article_page_url, pdf_src)
            return pdf_src
        embed = soup.find('embed', attrs={'type': 'application/pdf'})
        if embed and embed.get('src'):
            pdf_src = embed['src'] #; print(f"Encontrado PDF en embed: {pdf_src}")
            return urljoin(article_page_url, pdf_src)
        # print(f"No se encontró enlace PDF en iframe/embed para {article_page_url}")
        return None
    except requests.exceptions.RequestException as e: # print(f"Error al obtener HTML {article_page_url}: {e}");
        return None
    except Exception as e: # print(f"Error inesperado extrayendo PDF de {article_page_url}: {e}");
        return None

def download_from_google_scholar(doi, title, session):
    """
    Tries to download a PDF from Google Scholar using DOI.
    """
    # print(f"Searching Google Scholar for DOI: {doi} (Title: {title if title else 'N/A'})")
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
                    href = urljoin(scholar_url, href) # Ensure absolute URL

                # Basic check to avoid some common non-PDF page links that might contain "pdf"
                if 'pdf' in href.lower() and not any(x in href.lower() for x in ['view', 'download=false', 'scholar.google']):
                     potential_links.append(href)
                elif href.lower().endswith('.pdf'): # More direct .pdf links
                    potential_links.append(href)

        # De-duplicate while preserving order (important for prioritization if any)
        unique_potential_links = []
        for plink in potential_links:
            if plink not in unique_potential_links:
                unique_potential_links.append(plink)

        # print(f"Found {len(unique_potential_links)} potential PDF links on Google Scholar: {unique_potential_links}")

        for pdf_url in unique_potential_links:
            # print(f"Attempting to download PDF from: {pdf_url}")
            try:
                # Try HEAD request first (more efficient if server supports it well)
                head_response = session.head(pdf_url, headers=headers, timeout=20, allow_redirects=True)
                head_response.raise_for_status()
                content_type = head_response.headers.get('Content-Type', '').lower()

                if 'application/pdf' in content_type:
                    # print(f"HEAD request successful. Content-Type: {content_type}. Proceeding with GET.")
                    pdf_response = session.get(pdf_url, headers=headers, timeout=60, stream=True) # stream=True for large files
                    pdf_response.raise_for_status()

                    # Double check content type from GET response as well
                    get_content_type = pdf_response.headers.get('Content-Type', '').lower()
                    if 'application/pdf' in get_content_type:
                        pdf_content = pdf_response.content
                        # print(f"Successfully downloaded PDF from {pdf_url}")
                        domain_match = re.search(r'https?://(?:www\.)?([a-zA-Z0-9.-]+)(?:/|$)', pdf_url)
                        source_domain = domain_match.group(1) if domain_match else "Unknown Domain"
                        return pdf_content, f"OBTENIDO (Google Scholar via {source_domain})"
                    # else:
                        # print(f"GET request for {pdf_url} did not return PDF content-type, but: {get_content_type}")
                # else:
                    # print(f"HEAD request for {pdf_url} did not indicate PDF content-type: {content_type}")

            except requests.exceptions.HTTPError as e:
                # print(f"HTTP error when trying {pdf_url}: {e.response.status_code}")
                pass # Continue to next link
            except requests.exceptions.Timeout:
                # print(f"Timeout when trying {pdf_url}")
                pass # Continue to next link
            except requests.exceptions.RequestException as e:
                # print(f"Request error when trying {pdf_url}: {e}")
                pass # Continue to next link
            except Exception as e:
                # print(f"Unexpected error when trying {pdf_url}: {e}")
                pass # Continue to next link

        return None, "FALLO - No suitable PDF link from Google Scholar led to a successful download."

    except requests.exceptions.RequestException as e:
        # print(f"Error searching Google Scholar for DOI {doi}: {e}")
        return None, "FALLO - Google Scholar search error"
    except Exception as e:
        # print(f"Unexpected error during Google Scholar processing for DOI {doi}: {e}")
        return None, "FALLO - Unexpected error with Google Scholar"

def download_from_pmc(doi, title, session):
    # print(f"Attempting PubMed Central download for DOI: {doi}") # Keep internal prints commented for now
    try:
        # Step 1: Convert DOI to PMCID
        id_conv_url = f"https://www.ncbi.nlm.nih.gov/pmc/utils/idconv/v1.0/?ids={doi}&format=json&tool=my_tool&email=my_email@example.com"
        # print(f"PMC ID Converter URL: {id_conv_url}")
        try:
            response = session.get(id_conv_url, timeout=20)
            response.raise_for_status()
            data = response.json()
        except requests.exceptions.RequestException as e:
            # print(f"PMC ID Conv API request failed: {e}")
            return None, f"FALLO - Error API conversión PMCID ({str(e)[:50]})"
        except json.JSONDecodeError:
            # print(f"PMC ID Conv API JSON decode error. Response text: {response.text[:200]}")
            return None, "FALLO - Error decodificando respuesta PMCID"

        pmcid = None
        if data.get("records") and len(data["records"]) > 0:
            if "pmcid" in data["records"][0]:
                pmcid = data["records"][0]["pmcid"]
            # else:
                # print(f"PMCID not found in record: {data['records'][0]}")
        # else:
            # print(f"No records found in PMC ID Conv API response: {data}")

        if not pmcid:
            return None, f"FALLO - PMCID no encontrado para DOI {doi}"

        # print(f"Found PMCID: {pmcid}")

        # Step 2: Access PMC Article and Find PDF
        article_url = f"https://www.ncbi.nlm.nih.gov/pmc/articles/{pmcid}/"
        # print(f"PMC Article URL: {article_url}")
        try:
            response = session.get(article_url, timeout=30)
            response.raise_for_status()
        except requests.exceptions.RequestException as e:
            # print(f"PMC Article page request failed: {e}")
            return None, f"FALLO - Error obteniendo página PMC {pmcid} ({str(e)[:50]})"

        soup = BeautifulSoup(response.content, 'html.parser')

        pdf_links = []
        # Common pattern for PDF links on PMC
        for link_tag in soup.find_all('a', href=True):
            href = link_tag['href']
            # Check for links that explicitly point to a PDF rendition
            if '/pdf/' in href and href.lower().endswith('.pdf'):
                pdf_links.append(urljoin(article_url, href))
            # Check for links in the "Download" or "Formats" sections, often with specific attributes
            elif link_tag.get('title') == 'Download PDF' or 'pdf' in link_tag.get_text().lower():
                 if href.lower().endswith('.pdf') or '/pmc/articles/' in href.lower() and 'rendering=' in href.lower(): # Heuristic for render links
                    pdf_links.append(urljoin(article_url, href))


        # De-duplicate potential links
        unique_pdf_links = []
        for plink in pdf_links:
            if plink not in unique_pdf_links:
                unique_pdf_links.append(plink)

        # print(f"Found potential PMC PDF links: {unique_pdf_links}")

        if not unique_pdf_links:
            return None, f"FALLO - PDF no encontrado en la página PMC para {pmcid}"

        # Step 3: Download PDF
        for pdf_url_to_try in unique_pdf_links:
            # print(f"Attempting to download PDF from PMC link: {pdf_url_to_try}")
            try:
                head_response = session.head(pdf_url_to_try, timeout=20, allow_redirects=True)
                head_response.raise_for_status()
                content_type = head_response.headers.get('Content-Type', '').lower()

                if 'application/pdf' in content_type:
                    pdf_response = session.get(pdf_url_to_try, timeout=60, stream=True)
                    pdf_response.raise_for_status()
                    get_content_type = pdf_response.headers.get('Content-Type', '').lower() # Check again after GET
                    if 'application/pdf' in get_content_type:
                        # print(f"Successfully downloaded PDF from {pdf_url_to_try}")
                        return pdf_response.content, f"OBTENIDO (PubMed Central {pmcid})"
                    # else:
                        # print(f"PMC GET Content-Type not PDF: {get_content_type} from {pdf_url_to_try}")
                # else:
                    # print(f"PMC HEAD Content-Type not PDF: {content_type} from {pdf_url_to_try}")
            except requests.exceptions.RequestException as e:
                # print(f"Failed to download from PMC PDF link {pdf_url_to_try}: {e}")
                continue # Try next link

        return None, f"FALLO - No se pudo descargar PDF desde enlaces PMC para {pmcid}"

    except Exception as e:
        # print(f"Unexpected error in download_from_pmc for DOI {doi}: {e}")
        return None, f"FALLO - Error inesperado en PubMed Central ({str(e)[:50]})"

def print_to_console(message, orig_stdout):
    print(message, file=orig_stdout)

def download_pdfs_from_file():
    original_stdout = sys.stdout 
    # root = tk.Tk(); root.withdraw() # GUI elements removed
    # log_window = None; log_text_widget = None # GUI elements removed
    df = None # Initialize df to None
    temp_pdf_paths = [] # Initialize temp_pdf_paths as it's used in finally

    # GUI Log window and stdout redirection are disabled.
    # All print statements will go to the original stdout (console).

    try:
        # Since GUI dialogs are used, we still need a root Tk window for them,
        # but it can be managed more minimally if no log window is tied to it.
        # For now, assuming dialogs are essential and keeping root.
        root = tk.Tk()
        root.withdraw() # Keep it hidden as it's mainly for dialogs

        print("DIAG: --- Iniciando Configuración ---") # No longer using print_to_console
        input_file_path = filedialog.askopenfilename(title="Seleccionar archivo con DOIs (Excel o CSV)", filetypes=(("Archivos Excel", "*.xlsx *.xls"), ("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*")))
        print(f"DIAG: After input_file_path dialog, path: {input_file_path}")
        if not input_file_path: messagebox.showinfo("Información", "No se seleccionó ningún archivo de entrada. El programa terminará."); return
        print("DIAG: Before zip_path dialog")
        zip_path = filedialog.asksaveasfilename(title="Guardar archivo ZIP como...", defaultextension=".zip", filetypes=(("Archivos ZIP", "*.zip"), ("Todos los archivos", "*.*")))
        print(f"DIAG: After zip_path dialog, path: {zip_path}")
        if not zip_path: messagebox.showinfo("Información", "No se especificó la ubicación para guardar el ZIP. El programa terminará."); return
        print("DIAG: Before excel_report_path_config dialog")
        excel_report_path_config = filedialog.asksaveasfilename(title="Guardar Reporte Excel Opcional como...", defaultextension=".xlsx", initialfile="SciHub_Descarga_Reporte.xlsx", filetypes=(("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")))
        print(f"DIAG: After excel_report_path_config dialog, path: {excel_report_path_config}")
        if not excel_report_path_config: excel_report_path_config = ""; print("DIAG: Ruta para reporte Excel no especificada.")
        print("DIAG: Before user_inter_doi_delay dialog")
        user_inter_doi_delay = simpledialog.askinteger("Configurar Retraso Inter-DOI", "Ingrese el tiempo de espera (segundos) entre cada DOI:", initialvalue=INTER_DOI_DELAY_SECONDS, minvalue=0 )
        print(f"DIAG: After user_inter_doi_delay dialog, value: {user_inter_doi_delay}")
        if user_inter_doi_delay is None: user_inter_doi_delay = INTER_DOI_DELAY_SECONDS; messagebox.showinfo("Información", f"Retraso Inter-DOI no modificado, se usará el predeterminado: {user_inter_doi_delay}s")
        print("DIAG: Before user_mirror_switch_delay dialog")
        user_mirror_switch_delay = simpledialog.askinteger("Configurar Retraso Cambio de Mirror", "Ingrese el tiempo de espera (segundos) al cambiar de mirror:", initialvalue=MIRROR_SWITCH_DELAY_SECONDS, minvalue=0)
        print(f"DIAG: After user_mirror_switch_delay dialog, value: {user_mirror_switch_delay}")
        if user_mirror_switch_delay is None: user_mirror_switch_delay = MIRROR_SWITCH_DELAY_SECONDS; messagebox.showinfo("Información", f"Retraso por cambio de Mirror no modificado, se usará el predeterminado: {user_mirror_switch_delay}s")
        print("DIAG: Before user_mirror_list_str dialog")
        user_mirror_list_str = simpledialog.askstring("Configurar Mirrors de Sci-Hub (Obligatorio)", "Ingrese URLs de mirrors Sci-Hub, separadas por comas.\n(ej: https://sci-hub.se/,https://sci-hub.st/)\nESTOS SERÁN LOS ÚNICOS MIRRORS UTILIZADOS.\nDejar vacío para usar default (https://sci-hub.se/).")
        print(f"DIAG: After user_mirror_list_str dialog, value: {user_mirror_list_str}")
        user_defined_mirrors = []
        if user_mirror_list_str is None: messagebox.showerror("Configuración Requerida", "Configuración de mirrors cancelada. Terminando."); return
        if not user_mirror_list_str.strip(): messagebox.showinfo("Información de Mirrors", "No se ingresaron mirrors. Usando default: https://sci-hub.se/"); user_defined_mirrors = ["https://sci-hub.se/"]
        else:
            raw_mirrors = [mirror.strip() for mirror in user_mirror_list_str.split(',') if mirror.strip()]
            for mirror_url in raw_mirrors:
                if not mirror_url.startswith(("http://", "https://")): mirror_url = "https://" + mirror_url 
                if not mirror_url.endswith('/'): mirror_url += '/'
                user_defined_mirrors.append(mirror_url)
            if not user_defined_mirrors: messagebox.showerror("Error de Configuración", "Lista de mirrors vacía o inválida. Terminando."); return
        sci_hub_base_url_for_report = user_defined_mirrors[0]

        # print_to_console("DIAG: Before root.update() after all dialogs.", original_stdout); root.update(); print_to_console("DIAG: After root.update() after all dialogs.", original_stdout) # GUI logging disabled
        # print_to_console("DIAG: All dialogs complete. Before log_window creation.", original_stdout); log_window = tk.Toplevel(root); print_to_console("DIAG: After log_window = tk.Toplevel(root)", original_stdout) # GUI logging disabled
        # log_window.title("Log de Proceso de Descarga"); log_window.geometry("800x600") # GUI logging disabled
        # print_to_console("DIAG: Before log_text_widget creation", original_stdout); log_text_widget = st.ScrolledText(log_window, wrap=tk.WORD, state='disabled'); log_text_widget.pack(padx=10, pady=10, fill=tk.BOTH, expand=True); print_to_console("DIAG: After log_text_widget creation", original_stdout) # GUI logging disabled
        # log_window.protocol("WM_DELETE_WINDOW", lambda: on_log_window_close(log_window, original_stdout, root)) # GUI logging disabled
        # print_to_console("DIAG: Before sys.stdout redirection", original_stdout); sys.stdout = TextRedirector(log_text_widget, original_stdout) ; print_to_console("DIAG: After sys.stdout redirection. Subsequent app prints go to GUI and console.", original_stdout) # GUI logging disabled
        
        print("\n--- Configuración Aplicada ---")
        print(f"Archivo de entrada: {input_file_path}")
        print(f"Archivo ZIP de salida: {zip_path}")
        if excel_report_path_config: print(f"Archivo de reporte Excel: {excel_report_path_config}")
        else: print("Reporte Excel: No se generará (ruta no especificada).")
        print(f"Retraso Inter-DOI: {user_inter_doi_delay}s")
        print(f"Retraso Cambio de Mirror: {user_mirror_switch_delay}s")
        print(f"Mirrors Sci-Hub a utilizar: {', '.join(user_defined_mirrors)}")
        print("-----------------------------------------------------\n")
        # if log_window and log_window.winfo_exists(): log_window.update_idletasks() # GUI logging disabled

        session = requests.Session(); session.headers.update({'User-Agent': 'Mozilla/5.0...'})
        all_articles_log = []; successful_articles_data = []; failed_articles_data = []; original_input_columns = []
        
        try: # File reading try block
            file_extension = os.path.splitext(input_file_path)[1].lower()
            if file_extension in ['.xlsx', '.xls']:
                try:
                    df = pd.read_excel(input_file_path)
                except Exception as e:
                    messagebox.showerror("Error de Excel", f"Error al leer Excel: {e}")
                    raise # Re-raise to be caught by the outer exception handler for file reading
            elif file_extension == '.csv': # Explicitly handle CSV
                try:
                    df = pd.read_csv(input_file_path)
                except Exception as e:
                    messagebox.showerror("Error de CSV", f"Error al leer CSV: {e}")
                    raise # Re-raise
            else:
                messagebox.showerror("Error de Archivo", f"Tipo de archivo no soportado: {file_extension}\nPor favor, use Excel o CSV.")
                # Set df to None or raise an error to ensure it's handled before use
                # Raising an error is cleaner to be caught by the existing handler.
                raise ValueError(f"Tipo de archivo no soportado: {file_extension}")

            if df is not None: # Proceed only if df was loaded
                original_input_columns = [col for col in df.columns if col not in ['DOI', 'Title']]
            else: # Should not be reached if non-supported file types raise an error
                raise ValueError("DataFrame no fue cargado correctamente.")

        except Exception as e: # Catch any exception from file reading block
            print(f"Error fatal al leer archivo de entrada: {e}")
            # if log_window and log_window.winfo_exists(): log_window.destroy() # GUI logging disabled
            messagebox.showerror("Error Crítico de Archivo", f"No se pudo leer el archivo de entrada: {e}\nEl programa terminará.")
            return # Exit if file reading fails

        # Crucial check: If df is still None here, it means file reading failed in a way not caught above,
        # or a new path was introduced. This ensures we don't proceed.
        if df is None:
            print("Error Crítico: DataFrame (df) no fue inicializado. Terminando proceso.")
            # if log_window and log_window.winfo_exists(): log_window.destroy() # GUI logging disabled
            messagebox.showerror("Error Crítico", "DataFrame no pudo ser cargado. El programa terminará.")
            return

        successful_downloads = 0; failed_downloads_summary_list = []; total_downloaded_size_bytes = 0 # temp_pdf_paths already initialized
        zip_creation_or_main_loop_error = False 

        try: # Main processing try (zip creation and DOI loops)
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                total_articles = len(df) # df should be valid here
                for index, row in df.iterrows():
                    original_row_data = row.to_dict(); start_time = datetime.now()
                    doi = str(original_row_data.get('DOI', original_row_data.get('doi', ''))).strip()
                    title = str(original_row_data.get('Title', original_row_data.get('title', ''))).strip()
                    effective_title = title if title else doi
                    
                    current_article_num_for_log = index + 1
                    mirror_attempts_details_for_doi = []
                    overall_doi_status = "FALTANTE" # Default status

                    if not doi:
                        # ... (skip logic as before, but use local vars for log call)
                        # NOTE: The initial print block should be SKIPPED if DOI is empty.
                        # The existing 'continue' handles this.
                        failure_reason_for_report = "DOI vacío"; detailed_status_for_log = "Skipped_DOI_Missing"
                        overall_doi_status = "FALTANTE (DOI Vacío)"
                        # Call format_and_log_article_status here for skipped DOI
                        # successful_downloads count doesn't change for skipped
                        format_and_log_article_status(original_row_data, doi, effective_title, current_article_num_for_log, total_articles, successful_downloads, mirror_attempts_details_for_doi, overall_doi_status, user_inter_doi_delay)
                        end_time = datetime.now() # Ensure end_time for log
                        log_entry = {**original_row_data, 'Start_Time': start_time.strftime("%Y-%m-%d %H:%M:%S"), 'End_Time': end_time.strftime("%Y-%m-%d %H:%M:%S"), 'Duration_Seconds': (end_time - start_time).total_seconds(), 'Detailed_Status': detailed_status_for_log, 'Failure_Reason': failure_reason_for_report, 'Successful_Mirror': ""}
                        all_articles_log.append(log_entry); failed_articles_data.append({**original_row_data, 'Failure_Reason': failure_reason_for_report, 'Detailed_Status': detailed_status_for_log, 'original_index': index})
                        # if log_window and log_window.winfo_exists(): log_window.update_idletasks() # GUI logging disabled
                        # Separator after logging for a skipped DOI
                        print_to_console("===============================================================================================", original_stdout)
                        time.sleep(user_inter_doi_delay) # Still sleep for skipped
                        continue
                    
                    pdf_filename_in_zip = clean_filename(effective_title)[:150] + ".pdf"

                    # --- Initial Article Summary Print Block (Main Loop) ---
                    current_article_num = index + 1 # Already have current_article_num_for_log, can reuse
                    buscados_percentage = (current_article_num / total_articles) * 100 if total_articles > 0 else 0

                    author_val_initial = original_row_data.get('First Author', 'N/A')
                    if author_val_initial == 'N/A':
                        autores_keys_initial = [k for k in original_row_data.keys() if str(k).lower() == 'autores']
                        author_val_initial = original_row_data.get(autores_keys_initial[0], 'N/A') if autores_keys_initial else 'N/A'
                    if author_val_initial == 'N/A':
                        authors_keys_en_initial = [k for k in original_row_data.keys() if str(k).lower() == 'authors']
                        author_val_initial = original_row_data.get(authors_keys_en_initial[0], 'N/A') if authors_keys_en_initial else 'N/A'

                    journal_title_val_initial = original_row_data.get('Journal/Book', 'N/A')
                    if journal_title_val_initial == 'N/A':
                        revista_keys_initial = [k for k in original_row_data.keys() if str(k).lower() == 'revista']
                        journal_title_val_initial = original_row_data.get(revista_keys_initial[0], 'N/A') if revista_keys_initial else 'N/A'

                    pub_year_val_initial = original_row_data.get('Publication Year', 'N/A')
                    if pub_year_val_initial == 'N/A':
                        fecha_pub_keys_initial = [k for k in original_row_data.keys() if str(k).lower() == 'fecha de publicación']
                        pub_year_val_initial = original_row_data.get(fecha_pub_keys_initial[0], 'N/A') if fecha_pub_keys_initial else 'N/A'
                    if pub_year_val_initial == 'N/A':
                        year_keys_initial = [k for k in original_row_data.keys() if str(k).lower() == 'year']
                        pub_year_val_initial = original_row_data.get(year_keys_initial[0], 'N/A') if year_keys_initial else 'N/A'
                    if pub_year_val_initial == 'N/A':
                        ano_keys_initial = [k for k in original_row_data.keys() if str(k).lower() == 'año']
                        pub_year_val_initial = original_row_data.get(ano_keys_initial[0], 'N/A') if ano_keys_initial else 'N/A'

                    print(f"Artículo: {current_article_num}/{total_articles} ({buscados_percentage:.2f}%)")
                    print(f"Título: {effective_title if effective_title else 'N/A'}")
                    print(f"First Author: {author_val_initial}")
                    print(f"Journal/Book: {journal_title_val_initial}")
                    print(f"Publication Year: {pub_year_val_initial}")
                    print(f"DOI: {doi}")
                    print("-" * 30)
                    # --- End Initial Article Summary Print Block ---
                    
                    mirrors_to_try_for_this_doi = list(user_defined_mirrors)
                    pdf_content = None; download_successful_this_doi = False; successful_mirror_for_this_doi = ""
                    temp_detailed_status_for_log = ""; temp_failure_reason_for_log = ""

                    for mirror_idx, current_mirror_base_url in enumerate(mirrors_to_try_for_this_doi):
                        # Print for mirror attempt will be part of format_and_log_article_status details
                        full_sci_hub_url_for_html_page = f"{current_mirror_base_url}{doi}"
                        mirror_status_str = "FALLO"; mirror_reason_str = ""
                        
                        actual_pdf_download_url = extract_pdf_link_from_html(full_sci_hub_url_for_html_page, session)
                        if actual_pdf_download_url:
                            try:
                                response = session.get(actual_pdf_download_url, timeout=60); response.raise_for_status()
                                content_type = response.headers.get('Content-Type', '').lower()
                                if 'application/pdf' in content_type: 
                                    pdf_content = response.content; mirror_status_str = "OBTENIDO (Extracción Iframe/Embed)"; temp_detailed_status_for_log = f"Success_iframe_or_embed_extraction_from_{current_mirror_base_url}"
                                else: 
                                    mirror_reason_str = f"Content-Type no PDF ({content_type})"; temp_detailed_status_for_log = f"Failure_iframe_or_embed_extraction_not_pdf_from_{current_mirror_base_url}"
                            except requests.exceptions.HTTPError as e: mirror_reason_str = f"HTTPError {e.response.status_code}"; temp_detailed_status_for_log = f"Failure_iframe_or_embed_extraction_HTTP{e.response.status_code}_from_{current_mirror_base_url}"
                            except requests.exceptions.RequestException as e: mirror_reason_str = "Error de conexión/RequestException en extracción"; temp_detailed_status_for_log = f"Failure_iframe_or_embed_extraction_RequestException_from_{current_mirror_base_url}"
                            except Exception as e: mirror_reason_str = "Error inesperado en extracción"; temp_detailed_status_for_log = f"Failure_iframe_or_embed_extraction_Unexpected_from_{current_mirror_base_url}"
                        else: mirror_reason_str = "No se encontró enlace PDF en HTML"; temp_detailed_status_for_log = f"Failure_No_PDF_Link_Found_In_HTML_from_{current_mirror_base_url}"
                        
                        if not pdf_content: # Fallback if extraction failed
                            temp_failure_reason_for_log = mirror_reason_str # Keep reason from extraction attempt
                            try:
                                response = session.get(full_sci_hub_url_for_html_page, timeout=30); response.raise_for_status()
                                content_type = response.headers.get('Content-Type', '').lower()
                                if 'application/pdf' in content_type: 
                                    pdf_content = response.content; mirror_status_str = "OBTENIDO (Acceso Directo Fallback)"; mirror_reason_str = ""; temp_detailed_status_for_log = f"Success_direct_DOI_access_fallback_from_{current_mirror_base_url}"
                                else: 
                                    mirror_reason_str = f"Content-Type no PDF ({content_type}) en acceso directo"; temp_detailed_status_for_log = f"Failure_direct_DOI_access_not_pdf_from_{current_mirror_base_url}"
                            except requests.exceptions.HTTPError as e: mirror_reason_str = f"HTTPError {e.response.status_code} en acceso directo"; temp_detailed_status_for_log = f"Failure_direct_DOI_access_HTTP{e.response.status_code}_from_{current_mirror_base_url}"
                            except requests.exceptions.RequestException as e: mirror_reason_str = "Error de conexión/RequestException en fallback"; temp_detailed_status_for_log = f"Failure_direct_DOI_access_RequestException_from_{current_mirror_base_url}"
                            except Exception as e: mirror_reason_str = "Error inesperado en fallback"; temp_detailed_status_for_log = f"Failure_direct_DOI_access_Unexpected_from_{current_mirror_base_url}"
                        
                        # Logic for appending to mirror_attempts_details_for_doi and setting temp_failure_reason_for_log
                        specific_reason_for_temp_log = mirror_reason_str # Capture specific reason from this mirror

                        log_display_reason_for_sci_hub_attempt = mirror_reason_str
                        if mirror_status_str == "FALLO":
                            log_display_reason_for_sci_hub_attempt = full_sci_hub_url_for_html_page
                            temp_failure_reason_for_log = specific_reason_for_temp_log # Update overall DOI failure with specific reason from this mirror
                        else: # OBTENIDO from this mirror
                            temp_failure_reason_for_log = "" # Clear overall DOI failure reason

                        mirror_attempts_details_for_doi.append((current_mirror_base_url, mirror_status_str, log_display_reason_for_sci_hub_attempt))

                        if pdf_content: # Successfully got PDF from current_mirror_base_url
                            download_successful_this_doi = True
                            successful_mirror_for_this_doi = current_mirror_base_url
                            # overall_doi_status will be set to "OBTENIDO" or the more specific success message from mirror_status_str
                            if mirror_status_str.startswith("OBTENIDO"): overall_doi_status = mirror_status_str
                            else: overall_doi_status = "OBTENIDO"
                            temp_failure_reason_for_log = "" # Clear overall failure reason for the DOI
                            # temp_detailed_status_for_log is already set by the successful extraction/fallback
                            break # Break from Sci-Hub mirror loop
                        else:
                            # PDF not found with this mirror, temp_failure_reason_for_log has the specific reason
                            if mirror_idx < len(mirrors_to_try_for_this_doi) - 1:
                                time.sleep(user_mirror_switch_delay)

                    # After Sci-Hub loop, if still no pdf_content, try Google Scholar
                    if not pdf_content:
                        # print(f"Sci-Hub attempts failed for DOI {doi}. Trying Google Scholar.") # Debug print commented out
                        gs_pdf_content, gs_status_msg = download_from_google_scholar(doi, effective_title, session)
                        if gs_pdf_content:
                            pdf_content = gs_pdf_content
                            download_successful_this_doi = True
                            successful_mirror_for_this_doi = "Google Scholar"
                            overall_doi_status = gs_status_msg
                            mirror_attempts_details_for_doi.append(("Google Scholar", "OBTENIDO", gs_status_msg))
                            temp_detailed_status_for_log = f"Success_GoogleScholar_{gs_status_msg}"
                            temp_failure_reason_for_log = "" # Clear overall DOI failure as GS succeeded
                        else:
                            mirror_attempts_details_for_doi.append(("Google Scholar", "FALLO", gs_status_msg))
                            # temp_failure_reason_for_log is already set if all Sci-Hub mirrors failed.
                            # If Sci-Hub mirrors didn't run (e.g., empty list), then gs_status_msg becomes the failure reason.
                            # If Sci-Hub mirrors ran and failed, gs_status_msg will overwrite the last Sci-Hub specific error.
                            temp_failure_reason_for_log = gs_status_msg
                            temp_detailed_status_for_log = f"Failure_GoogleScholar_{gs_status_msg}"

                    # After Google Scholar attempt, if still no pdf_content, try PubMed Central
                    if not pdf_content:
                        # Optional: print(f"Sci-Hub and Google Scholar failed for {doi}. Trying PubMed Central.")
                        pmc_pdf_content, pmc_status_msg = download_from_pmc(doi, effective_title, session)
                        if pmc_pdf_content:
                            pdf_content = pmc_pdf_content
                            download_successful_this_doi = True
                            successful_mirror_for_this_doi = pmc_status_msg # e.g., "OBTENIDO (PubMed Central PMCID_HERE)"
                            overall_doi_status = pmc_status_msg
                            mirror_attempts_details_for_doi.append(("PubMed Central", "OBTENIDO", pmc_status_msg))
                            temp_detailed_status_for_log = f"Success_PubMedCentral_{pmc_status_msg}"
                            temp_failure_reason_for_log = "" # Clear failure reason as PMC succeeded
                        else:
                            mirror_attempts_details_for_doi.append(("PubMed Central", "FALLO", pmc_status_msg))
                            temp_failure_reason_for_log = pmc_status_msg # Update with PMC failure reason
                            temp_detailed_status_for_log = f"Failure_PubMedCentral_{pmc_status_msg}"

                    end_time = datetime.now()
                    if download_successful_this_doi and pdf_content:
                        successful_downloads += 1 # Increment *before* calling log for current stats
                        # overall_doi_status is already set if successful (either by SciHub or GS)
                        if not overall_doi_status.startswith("OBTENIDO"): # Ensure it's marked OBTENIDO if somehow missed
                            overall_doi_status = "OBTENIDO"
                        # ... (save PDF to zip as before) ...
                        data_for_successful_sheet = original_row_data.copy(); data_for_successful_sheet['Successful_Mirror'] = successful_mirror_for_this_doi; successful_articles_data.append(data_for_successful_sheet)
                        temp_dir = "temp_scihub_pdfs"; 
                        if not os.path.exists(temp_dir): os.makedirs(temp_dir)
                        temp_pdf_path = os.path.join(temp_dir, f"temp_{os.getpid()}_{index}_{pdf_filename_in_zip}")
                        with open(temp_pdf_path, 'wb') as f: f.write(pdf_content)
                        try: total_downloaded_size_bytes += os.path.getsize(temp_pdf_path)
                        except OSError as e: print(f"Advertencia: tamaño temp {temp_pdf_path}: {e}") # This print will go to log
                        zf.write(temp_pdf_path, arcname=pdf_filename_in_zip); temp_pdf_paths.append(temp_pdf_path)
                    else:
                        overall_doi_status = "FALTANTE"
                        # temp_failure_reason_for_log and temp_detailed_status_for_log will have the last failure
                        failed_articles_data.append({**original_row_data, 'Failure_Reason': temp_failure_reason_for_log, 'Detailed_Status': temp_detailed_status_for_log, 'original_index': index}) # Store original index

                    format_and_log_article_status(original_row_data, doi, effective_title, current_article_num_for_log, total_articles, successful_downloads, mirror_attempts_details_for_doi, overall_doi_status, user_inter_doi_delay)
                    
                    log_entry_failure_reason = temp_failure_reason_for_log if not download_successful_this_doi else ""
                    log_entry_detailed_status = temp_detailed_status_for_log if not download_successful_this_doi else f"Success_{successful_mirror_for_this_doi}" # Simplified success status for log

                    all_articles_log.append({**original_row_data, 'Start_Time': start_time.strftime("%Y-%m-%d %H:%M:%S"), 'End_Time': end_time.strftime("%Y-%m-%d %H:%M:%S"), 'Duration_Seconds': (end_time - start_time).total_seconds(), 'Detailed_Status': log_entry_detailed_status, 'Failure_Reason': log_entry_failure_reason, 'Successful_Mirror': successful_mirror_for_this_doi })
                    
                    # if log_window and log_window.winfo_exists(): log_window.update_idletasks() # GUI logging disabled
                    # Separator after logging for the current article in the main loop
                    print_to_console("===============================================================================================", original_stdout)
                    if current_article_num_for_log < total_articles : time.sleep(user_inter_doi_delay) # Sleep if not the last article
                    # Ensure overall_doi_status is updated for the log entry if it was a GS success
                    if download_successful_this_doi and "Google Scholar" in successful_mirror_for_this_doi:
                         log_entry_detailed_status = f"Success_{successful_mirror_for_this_doi}" # Update for GS success

        # ... (except FileNotFoundError, Exception for zip as before) ...
        except FileNotFoundError: messagebox.showerror("Error", f"No se pudo crear ZIP (Directorio no encontrado): {zip_path}"); print("Error crítico: FileNotFoundError al crear ZIP."); zip_creation_or_main_loop_error = True
        except Exception as e: messagebox.showerror("Error", f"Error inesperado en ZIP o descargas: {e}"); print(f"Error crítico: Excepción en ZIP o descargas: {e}."); zip_creation_or_main_loop_error = True


        if not zip_creation_or_main_loop_error:
            if failed_articles_data: 
                print("\n--- Iniciando Fase de Reintento para Artículos Fallidos ---")
                # if log_window and log_window.winfo_exists(): log_window.update_idletasks() # GUI logging disabled
            
            articles_successfully_retried_ids = [] 
            temp_failed_articles_data_for_iteration = list(failed_articles_data) 
            mirrors_for_retry = list(user_defined_mirrors)

            for retry_idx, failed_article_entry in enumerate(temp_failed_articles_data_for_iteration):
                original_index_for_retry = failed_article_entry.get('original_index', -1) # Get original index
                current_article_num_for_log_retry = original_index_for_retry + 1 if original_index_for_retry != -1 else retry_idx + 1 # Use original index for overall progress

                doi_to_retry = str(failed_article_entry.get('DOI', failed_article_entry.get('doi', ''))).strip()
                effective_title_for_retry = str(failed_article_entry.get('Title', failed_article_entry.get('title', doi_to_retry))).strip() or doi_to_retry
                pdf_filename_in_zip_retry = clean_filename(effective_title_for_retry)[:150] + ".pdf"

                # --- Initial Article Summary Print Block (Retry Loop) ---
                # current_article_num_for_log_retry and total_articles are already available
                buscados_percentage_retry = (current_article_num_for_log_retry / total_articles) * 100 if total_articles > 0 else 0

                # Extract author, journal, year from failed_article_entry (which is original_row_data for this item)
                author_val_retry = failed_article_entry.get('First Author', 'N/A')
                if author_val_retry == 'N/A':
                    autores_keys_retry = [k for k in failed_article_entry.keys() if str(k).lower() == 'autores']
                    author_val_retry = failed_article_entry.get(autores_keys_retry[0], 'N/A') if autores_keys_retry else 'N/A'
                if author_val_retry == 'N/A':
                    authors_keys_en_retry = [k for k in failed_article_entry.keys() if str(k).lower() == 'authors']
                    author_val_retry = failed_article_entry.get(authors_keys_en_retry[0], 'N/A') if authors_keys_en_retry else 'N/A'

                journal_title_val_retry = failed_article_entry.get('Journal/Book', 'N/A')
                if journal_title_val_retry == 'N/A':
                    revista_keys_retry = [k for k in failed_article_entry.keys() if str(k).lower() == 'revista']
                    journal_title_val_retry = failed_article_entry.get(revista_keys_retry[0], 'N/A') if revista_keys_retry else 'N/A'

                pub_year_val_retry = failed_article_entry.get('Publication Year', 'N/A')
                if pub_year_val_retry == 'N/A':
                    fecha_pub_keys_retry = [k for k in failed_article_entry.keys() if str(k).lower() == 'fecha de publicación']
                    pub_year_val_retry = failed_article_entry.get(fecha_pub_keys_retry[0], 'N/A') if fecha_pub_keys_retry else 'N/A'
                if pub_year_val_retry == 'N/A':
                    year_keys_retry = [k for k in failed_article_entry.keys() if str(k).lower() == 'year']
                    pub_year_val_retry = failed_article_entry.get(year_keys_retry[0], 'N/A') if year_keys_retry else 'N/A'
                if pub_year_val_retry == 'N/A':
                    ano_keys_retry = [k for k in failed_article_entry.keys() if str(k).lower() == 'año']
                    pub_year_val_retry = failed_article_entry.get(ano_keys_retry[0], 'N/A') if ano_keys_retry else 'N/A'

                print(f"[REINTENTO] Artículo: {current_article_num_for_log_retry}/{total_articles} ({buscados_percentage_retry:.2f}%)")
                print(f"[REINTENTO] Título: {effective_title_for_retry if effective_title_for_retry else 'N/A'}")
                print(f"[REINTENTO] First Author: {author_val_retry}")
                print(f"[REINTENTO] Journal/Book: {journal_title_val_retry}")
                print(f"[REINTENTO] Publication Year: {pub_year_val_retry}")
                print(f"[REINTENTO] DOI: {doi_to_retry}")
                print("-" * 30)
                # --- End Initial Article Summary Print Block (Retry Loop) ---
                
                mirror_attempts_details_for_retry = []
                overall_retry_status = "FALTANTE"
                pdf_content_retry = None; retry_successful_this_doi = False; successful_mirror_for_retry = ""
                temp_detailed_status_for_retry_log = ""; temp_failure_reason_for_retry_log = ""
                
                retry_start_time_actual_attempt = datetime.now() # For duration calculation of this specific retry in log

                for mirror_idx_retry, current_mirror_base_url_retry in enumerate(mirrors_for_retry):
                    # ... (retry mirror attempt logic, populate mirror_attempts_details_for_retry) ...
                    # ... (similar to main loop's mirror logic, use _retry suffixed variables)
                    full_sci_hub_url_for_html_page_retry = f"{current_mirror_base_url_retry}{doi_to_retry}"
                    mirror_status_str_retry = "FALLO"; mirror_reason_str_retry = ""
                    actual_pdf_download_url_retry = extract_pdf_link_from_html(full_sci_hub_url_for_html_page_retry, session)
                    if actual_pdf_download_url_retry:
                        try: # ... (extraction)
                            response = session.get(actual_pdf_download_url_retry, timeout=60); response.raise_for_status()
                            content_type = response.headers.get('Content-Type', '').lower()
                            if 'application/pdf' in content_type: pdf_content_retry = response.content; mirror_status_str_retry = "OBTENIDO (REINTENTO Extracción)"; temp_detailed_status_for_retry_log = f"Success_RETRY_iframe_embed_from_{current_mirror_base_url_retry}"
                            else: mirror_reason_str_retry = f"RETRY: Content-Type not PDF ({content_type})"; temp_detailed_status_for_retry_log = f"Failure_RETRY_iframe_embed_not_pdf_from_{current_mirror_base_url_retry}"
                        except requests.exceptions.HTTPError as e: mirror_reason_str_retry = f"RETRY: HTTPError {e.response.status_code}"; temp_detailed_status_for_retry_log = f"Failure_RETRY_iframe_embed_HTTPError_from_{current_mirror_base_url_retry}"
                        except requests.exceptions.RequestException as e: mirror_reason_str_retry = "RETRY: Error de conexión/RequestException en extracción"; temp_detailed_status_for_retry_log = f"Failure_RETRY_iframe_embed_RequestException_from_{current_mirror_base_url_retry}"
                        except Exception as e: mirror_reason_str_retry = "RETRY: Error inesperado en extracción"; temp_detailed_status_for_retry_log = f"Failure_RETRY_iframe_embed_Unexpected_from_{current_mirror_base_url_retry}"
                    else: mirror_reason_str_retry = "RETRY: No se encontró enlace PDF en HTML"; temp_detailed_status_for_retry_log = f"Failure_RETRY_No_PDF_Link_Found_In_HTML_from_{current_mirror_base_url_retry}"

                    if not pdf_content_retry: # Fallback
                        temp_failure_reason_for_retry_log = mirror_reason_str_retry # Preserve reason from extraction if it failed there
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

                    # Logic for appending to mirror_attempts_details_for_retry and setting temp_failure_reason_for_retry_log
                    specific_reason_for_temp_retry_log = mirror_reason_str_retry # Capture specific reason

                    log_display_reason_for_sci_hub_retry_attempt = mirror_reason_str_retry
                    if mirror_status_str_retry == "FALLO":
                        log_display_reason_for_sci_hub_retry_attempt = full_sci_hub_url_for_html_page_retry
                        temp_failure_reason_for_retry_log = specific_reason_for_temp_retry_log
                    else: # OBTENIDO from this mirror
                        temp_failure_reason_for_retry_log = ""

                    mirror_attempts_details_for_retry.append((current_mirror_base_url_retry, mirror_status_str_retry, log_display_reason_for_sci_hub_retry_attempt))

                    if pdf_content_retry: # Successfully got PDF from current_mirror_base_url_retry
                        retry_successful_this_doi = True
                        successful_mirror_for_retry = current_mirror_base_url_retry
                        if mirror_status_str_retry.startswith("OBTENIDO"): overall_retry_status = mirror_status_str_retry
                        else: overall_retry_status = "OBTENIDO"
                        temp_failure_reason_for_retry_log = ""
                        # temp_detailed_status_for_retry_log is already set
                        break # Break from Sci-Hub retry mirror loop
                    else:
                        # PDF not found with this mirror, temp_failure_reason_for_retry_log has the specific reason
                        if mirror_idx_retry < len(mirrors_for_retry) - 1:
                            time.sleep(user_mirror_switch_delay)

                # After Sci-Hub retry loop, if still no pdf_content_retry, try Google Scholar
                if not pdf_content_retry:
                    # print(f"Sci-Hub retry attempts failed for DOI {doi_to_retry}. Trying Google Scholar.") # Debug print commented out
                    gs_pdf_content_retry, gs_status_msg_retry = download_from_google_scholar(doi_to_retry, effective_title_for_retry, session)
                    if gs_pdf_content_retry:
                        pdf_content_retry = gs_pdf_content_retry
                        retry_successful_this_doi = True
                        successful_mirror_for_retry = "Google Scholar"
                        overall_retry_status = gs_status_msg_retry
                        mirror_attempts_details_for_retry.append(("Google Scholar (Retry)", "OBTENIDO", gs_status_msg_retry))
                        temp_detailed_status_for_retry_log = f"Success_RETRY_GoogleScholar_{gs_status_msg_retry}"
                        temp_failure_reason_for_retry_log = ""
                    else:
                        mirror_attempts_details_for_retry.append(("Google Scholar (Retry)", "FALLO", gs_status_msg_retry))
                        temp_failure_reason_for_retry_log = gs_status_msg_retry
                        temp_detailed_status_for_retry_log = f"Failure_RETRY_GoogleScholar_{gs_status_msg_retry}"

                # If Sci-Hub retry and Google Scholar retry failed, try PubMed Central for retry
                if not pdf_content_retry:
                    # print(f"Sci-Hub & Google Scholar retry attempts failed for DOI {doi_to_retry}. Trying PubMed Central.") # Debug print
                    pmc_pdf_content_retry, pmc_status_msg_retry = download_from_pmc(doi_to_retry, effective_title_for_retry, session)
                    if pmc_pdf_content_retry:
                        pdf_content_retry = pmc_pdf_content_retry
                        retry_successful_this_doi = True
                        successful_mirror_for_retry = "PubMed Central"
                        overall_retry_status = pmc_status_msg_retry
                        mirror_attempts_details_for_retry.append(("PubMed Central (Retry)", "OBTENIDO", pmc_status_msg_retry))
                        temp_detailed_status_for_retry_log = f"Success_RETRY_PMC_{pmc_status_msg_retry}"
                        temp_failure_reason_for_retry_log = "" # Clear overall failure
                    else:
                        mirror_attempts_details_for_retry.append(("PubMed Central (Retry)", "FALLO", pmc_status_msg_retry))
                        temp_failure_reason_for_retry_log = pmc_status_msg_retry
                        temp_detailed_status_for_retry_log = f"Failure_RETRY_PubMedCentral_{pmc_status_msg_retry}"

                retry_end_time_actual_attempt = datetime.now()
                original_article_log_entry = next((log for log in all_articles_log if str(log.get('DOI', log.get('doi', ''))).strip() == doi_to_retry), None)

                if retry_successful_this_doi and pdf_content_retry:
                    successful_downloads += 1 # Increment before log
                    # overall_retry_status is already set if successful
                    if not overall_retry_status.startswith("OBTENIDO"):
                         overall_retry_status = "OBTENIDO"
                    # ... (update successful_articles_data, save PDF, add to ZIP as before) ...
                    articles_successfully_retried_ids.append(doi_to_retry)
                    original_data_for_success = {k: v for k, v in failed_article_entry.items() if k not in ['Failure_Reason', 'Detailed_Status', 'original_index']}; original_data_for_success['Successful_Mirror'] = successful_mirror_for_retry; successful_articles_data.append(original_data_for_success)
                    temp_dir_retry = "temp_scihub_pdfs"; 
                    if not os.path.exists(temp_dir_retry): os.makedirs(temp_dir_retry)
                    temp_pdf_path_retry = os.path.join(temp_dir_retry, f"temp_RETRY_{os.getpid()}_{retry_idx}_{pdf_filename_in_zip_retry}")
                    with open(temp_pdf_path_retry, 'wb') as f: f.write(pdf_content_retry)
                    try: total_downloaded_size_bytes += os.path.getsize(temp_pdf_path_retry)
                    except OSError as e: print(f"Advertencia: tamaño temp (reintento) {temp_pdf_path_retry}: {e}")
                    try:
                        with zipfile.ZipFile(zip_path, 'a', zipfile.ZIP_DEFLATED) as zf_append: zf_append.write(temp_pdf_path_retry, arcname=pdf_filename_in_zip_retry)
                    except Exception as e: print(f"Error CRÍTICO al agregar PDF (reintento) '{pdf_filename_in_zip_retry}' al ZIP: {e}")
                    temp_pdf_paths.append(temp_pdf_path_retry) 
                    if original_article_log_entry: # Update original log entry
                        original_article_log_entry['Detailed_Status'] = temp_detailed_status_for_retry_log if not ("Google Scholar" in successful_mirror_for_retry) else f"Success_RETRY_GoogleScholar_{successful_mirror_for_retry}"
                        original_article_log_entry['Failure_Reason'] = "" if retry_successful_this_doi else temp_failure_reason_for_retry_log
                        original_article_log_entry['Successful_Mirror'] = successful_mirror_for_retry
                        original_article_log_entry['End_Time'] = retry_end_time_actual_attempt.strftime("%Y-%m-%d %H:%M:%S")
                        original_start_dt = datetime.strptime(original_article_log_entry['Start_Time'], "%Y-%m-%d %H:%M:%S")
                        original_article_log_entry['Duration_Seconds'] = (retry_end_time_actual_attempt - original_start_dt).total_seconds()
                else:
                    # overall_retry_status is already FALTANTE by default or set by GS if it failed
                    if not overall_retry_status.startswith("FALLO"): # ensure it is if we are in this block.
                        overall_retry_status = "FALTANTE"
                    if original_article_log_entry: # Update original log entry with latest failure
                        original_article_log_entry['Detailed_Status'] = temp_detailed_status_for_retry_log
                        original_article_log_entry['Failure_Reason'] = temp_failure_reason_for_retry_log
                    for item in failed_articles_data: # Update reason in failed_articles_data for Excel
                        if str(item.get('DOI',item.get('doi',''))).strip() == doi_to_retry: item['Failure_Reason'] = temp_failure_reason_for_retry_log; item['Detailed_Status'] = temp_detailed_status_for_retry_log; break
                
                # Log for this retry attempt
                # Pass original_row_data from failed_article_entry for Revista, Fecha Pub, etc.
                format_and_log_article_status(failed_article_entry, doi_to_retry, effective_title_for_retry, current_article_num_for_log_retry, total_articles, successful_downloads, mirror_attempts_details_for_retry, overall_retry_status, user_inter_doi_delay, is_retry=True)

                # Separator after logging for the current article in the retry loop
                print_to_console("===============================================================================================", original_stdout)
                # if log_window and log_window.winfo_exists(): log_window.update_idletasks() # GUI logging disabled
                if retry_idx < len(temp_failed_articles_data_for_iteration) - 1: time.sleep(user_inter_doi_delay)
            
            if articles_successfully_retried_ids: failed_articles_data = [item for item in failed_articles_data if str(item.get('DOI', item.get('doi', ''))).strip() not in articles_successfully_retried_ids]

            # if log_window and log_window.winfo_exists(): log_window.update_idletasks() # GUI logging disabled
            failed_downloads_summary_list = [{'title': str(item.get('Title',item.get('title','N/A'))).strip(), 'doi': str(item.get('DOI',item.get('doi','N/A'))).strip(), 'reason': str(item.get('Failure_Reason','N/A')).strip()} for item in failed_articles_data]
            total_mb = total_downloaded_size_bytes / (1024 * 1024)
            summary_message = (f"Proceso completado.\n\nDescargas exitosas: {successful_downloads}\nDescargas fallidas: {len(failed_downloads_summary_list)}\n" f"Tamaño total PDFs: {total_mb:.2f} MB")
            if failed_downloads_summary_list:
                summary_message += "\n\nArtículos no descargados (post-reintentos):"
                for item in failed_downloads_summary_list:
                    summary_message += f"\n- Título: {item['title']}, DOI: {item['doi']}, Razón: {item['reason']}"
            print("\n" + "="*50); print(summary_message); print("="*50); 
            # if log_window and log_window.winfo_exists(): log_window.update_idletasks() # GUI logging disabled
            messagebox.showinfo("Resumen Descarga", summary_message) # Keep messagebox for summary
            generate_excel_report_prompt = messagebox.askyesno("Generar Reporte Excel", "¿Desea generar un reporte Excel detallado?")
            if generate_excel_report_prompt:
                excel_report_path_to_use = excel_report_path_config
                if not excel_report_path_to_use : 
                    excel_report_path_to_use = filedialog.asksaveasfilename(title="Guardar Reporte Excel como...", defaultextension=".xlsx", initialfile="SciHub_Descarga_Reporte.xlsx", filetypes=(("Archivos Excel", "*.xlsx"),("Todos los archivos", "*.*")))
                elif not messagebox.askyesno("Confirmar Ruta de Reporte", f"Se configuró guardar el reporte en:\n{excel_report_path_config}\n\n¿Usar esta ruta?"): 
                        excel_report_path_to_use = filedialog.asksaveasfilename(title="Guardar Reporte Excel en ruta alternativa...", defaultextension=".xlsx", initialfile="SciHub_Descarga_Reporte_Alternativo.xlsx", filetypes=(("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")))
                if excel_report_path_to_use:
                    print(f"Generando reporte Excel en: {excel_report_path_to_use}")
                    # if log_window and log_window.winfo_exists(): log_window.update_idletasks() # GUI logging disabled
                    try:
                        ob_cols = ['DOI','Title','Successful_Mirror'] + [c for c in original_input_columns if c not in ['DOI','Title','Successful_Mirror']] + ['SciHub_Link']
                        fa_cols = ['DOI','Title'] + [c for c in original_input_columns if c not in ['DOI','Title','Failure_Reason','Detailed_Status']] + ['Failure_Reason','Detailed_Status','SciHub_Link']
                        ti_cols = ['DOI','Title'] + [c for c in original_input_columns if c not in ['DOI','Title','Successful_Mirror','Start_Time','End_Time','Duration_Seconds','Detailed_Status','Failure_Reason']] + ['Successful_Mirror','Start_Time','End_Time','Duration_Seconds','Detailed_Status','Failure_Reason','SciHub_Link']
                        def create_ordered_df(data, cols):
                            df_ = pd.DataFrame(data); df_['SciHub_Link'] = df_.apply(lambda r: f"{sci_hub_base_url_for_report}{r.get('DOI',r.get('doi',''))}" if pd.notna(r.get('DOI',r.get('doi',''))) else '', axis=1)
                            for c in cols: 
                                if c not in df_.columns: df_[c] = pd.NA
                            # Ensure all original_row_data keys also get included if not in standard cols
                            all_data_keys = set()
                            if data: all_data_keys.update(k for item in data for k in item.keys())
                            final_cols_ordered = [c for c in cols if c in all_data_keys or c == 'SciHub_Link'] # Start with preferred that exist or are added
                            final_cols_ordered.extend([k for k in all_data_keys if k not in final_cols_ordered and k != 'SciHub_Link']) # Add remaining data keys
                            return df_.reindex(columns=final_cols_ordered)

                        df_obtenidos = create_ordered_df(successful_articles_data, ob_cols)
                        df_fallidos = create_ordered_df(failed_articles_data, fa_cols)
                        df_tiempos = create_ordered_df(all_articles_log, ti_cols)
                        with pd.ExcelWriter(excel_report_path_to_use, engine='openpyxl') as writer:
                            df_obtenidos.to_excel(writer, sheet_name='Obtenidos', index=False); df_fallidos.to_excel(writer, sheet_name='Fallidos', index=False); df_tiempos.to_excel(writer, sheet_name='Tiempos', index=False)
                        messagebox.showinfo("Reporte Excel", f"Reporte Excel guardado: {excel_report_path_to_use}"); print(f"Reporte Excel guardado: {excel_report_path_to_use}")
                    except Exception as e: messagebox.showerror("Error Guardando Excel", f"No se pudo guardar reporte Excel: {e}"); print(f"Error Excel: {e}")
                else: print("Generación de reporte Excel omitida (ruta no proporcionada).")
            else: print("Generación de reporte Excel omitida por usuario.")

        elif zip_creation_or_main_loop_error: print("Proceso interrumpido por error crítico inicial. No se generará resumen ni Excel.")
    
    finally: 
        # restored_to_original_console = False # Not needed, stdout not redirected
        # Ensure stdout is the original, in case it was somehow still a TextRedirector
        # (though it shouldn't be if TextRedirector class and its usage are commented out)
        if hasattr(original_stdout, 'write'): # Check if original_stdout was captured
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
            try: os.rmdir(temp_dir_to_check); print(f"Eliminado dir temp: {temp_dir_to_check}", file=original_stdout if hasattr(original_stdout, 'write') else sys.__stdout__)
            except OSError as e: print(f"Error eliminando dir temp {temp_dir_to_check}: {e}", file=original_stdout if hasattr(original_stdout, 'write') else sys.__stdout__)

        print("--- Limpieza Finalizada ---", file=original_stdout if hasattr(original_stdout, 'write') else sys.__stdout__)

        # if log_window: # GUI logging disabled
        #     try:
        #         if log_window.winfo_exists(): log_window.destroy()
        #     except tk.TclError: pass

        # Destroy the main hidden Tkinter window if it exists
        if 'root' in locals() and isinstance(root, tk.Tk):
            try:
                root.destroy()
            except tk.TclError:
                pass


if __name__ == "__main__":
    # Store the true original stdout at the very beginning if not already done for the main function
    # This is a fallback if download_pdfs_from_file itself had an issue before original_stdout was set.
    # However, original_stdout inside download_pdfs_from_file should be sys.__stdout__ if not redirected.
    # The check `isinstance(sys.stdout, TextRedirector)` is no longer relevant if TextRedirector is removed.

    # The critical part is ensuring sys.stdout is sys.__stdout__ if it got changed.
    # Since TextRedirector is commented out, sys.stdout should not be an instance of it.
    # The `finally` block in `download_pdfs_from_file` now tries to restore original_stdout.
    # A final check here could be made, but might be redundant if `download_pdfs_from_file` handles it.

    # Safest is to assume download_pdfs_from_file cleans up its own stdout if it changed it.
    # If TextRedirector and its instantiation are fully removed, sys.stdout should remain as console.
    download_pdfs_from_file()
    # No need to check for TextRedirector instance if it's removed.
    # if isinstance(sys.stdout, TextRedirector) and hasattr(sys.stdout, 'original_stdout'):
    #    sys.stdout = sys.stdout.original_stdout # Restore if it was a TextRedirector
    # elif isinstance(sys.stdout, TextRedirector): # Fallback if original_stdout attribute missing
    #    sys.stdout = sys.__stdout__

    print("\nScript finalizado.")
