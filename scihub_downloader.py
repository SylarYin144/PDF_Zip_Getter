import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import pandas as pd
import requests
import zipfile
import os
import re
import time
from datetime import datetime
from bs4 import BeautifulSoup
from urllib.parse import urljoin

# --- Configuration Constants ---
DEFAULT_SCI_HUB_MIRRORS = [
    "https://sci-hub.se/",
    "https://sci-hub.st/",
    "https://sci-hub.ru/",
    "https://sci-hub.red/",
    "https://NFTsci-hub.box/", # Assuming this is a valid Sci-Hub mirror URL ending with /
    "https://sci-hub.wf/", 
    "https://sci-hub.cat/"
]
INTER_DOI_DELAY_SECONDS = 5
MIRROR_SWITCH_DELAY_SECONDS = 3
# --- End Configuration Constants ---

def clean_filename(title):
    """
    Limpia el título para que sea un nombre de archivo válido.
    Reemplaza caracteres no permitidos con guiones bajos.
    """
    return re.sub(r'[\\/*?:"<>|]', '_', title)

def extract_pdf_link_from_html(article_page_url, session):
    """
    Extrae el enlace directo al PDF desde el HTML de la página del artículo en Sci-Hub.
    """
    print(f"Fetching HTML page: {article_page_url}")
    try:
        response = session.get(article_page_url, timeout=30)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, 'html.parser')

        # Buscar iframe con id="pdf"
        iframe = soup.find('iframe', id='pdf')
        if iframe and iframe.get('src'):
            pdf_src = iframe['src']
            print(f"Found PDF source in iframe: {pdf_src}")
            if pdf_src.startswith('//'):
                return 'https:' + pdf_src
            elif pdf_src.startswith('/'): # En caso de rutas relativas al servidor
                 return urljoin(article_page_url, pdf_src)
            return pdf_src

        # Fallback: Buscar embed tag
        embed = soup.find('embed', attrs={'type': 'application/pdf'})
        if embed and embed.get('src'):
            pdf_src = embed['src']
            print(f"Found PDF source in embed tag: {pdf_src}")
            # Los src de embed pueden ser relativos a la URL base de la página del artículo
            return urljoin(article_page_url, pdf_src)

        print(f"Could not find PDF source in iframe or embed tag for {article_page_url}")
        return None
    except requests.exceptions.RequestException as e:
        print(f"Error fetching HTML page {article_page_url}: {e}")
        return None
    except Exception as e:
        print(f"An unexpected error occurred while extracting PDF link from {article_page_url}: {e}")
        return None

def download_pdfs_from_file():
    """
    Función principal para descargar PDFs desde un archivo Excel o CSV de DOIs.
    """
    # Inicializar Tkinter (necesario para los diálogos)
    root = tk.Tk()
    root.withdraw()  # Ocultar la ventana principal de Tkinter

    # Initialize requests session
    session = requests.Session()
    session.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'})

    # Initialize data storage lists for Excel reporting
    all_articles_log = []
    successful_articles_data = []
    failed_articles_data = []
    
    original_input_columns = [] # To store original column names for Excel ordering

    # 1. Entrada de URL de Sci-Hub por el Usuario
    # The user_sci_hub_url will be the first one tried.
    # The DEFAULT_SCI_HUB_MIRRORS list will be used if the user-provided one fails or if user input is skipped.
    initial_mirror_to_try = DEFAULT_SCI_HUB_MIRRORS[0] if DEFAULT_SCI_HUB_MIRRORS else "https://sci-hub.se/"
    
    user_sci_hub_url_input = simpledialog.askstring(
        "URL de Sci-Hub Principal",
        "Ingresa la URL principal de Sci-Hub a intentar primero:",
        initialvalue=initial_mirror_to_try
    )

    primary_sci_hub_url_to_use = ""

    if user_sci_hub_url_input is None: # Usuario cerró el diálogo
        messagebox.showinfo("Información", "No se ingresó URL de Sci-Hub. El programa terminará.")
        return
    elif not user_sci_hub_url_input.strip(): # Usuario ingresó cadena vacía
        primary_sci_hub_url_to_use = initial_mirror_to_try
        messagebox.showinfo("Información", f"URL de Sci-Hub no especificada, se usará la predeterminada: {primary_sci_hub_url_to_use}")
    else:
        primary_sci_hub_url_to_use = user_sci_hub_url_input.strip()
        if not primary_sci_hub_url_to_use.endswith('/'):
            primary_sci_hub_url_to_use += '/'
    
    # This primary_sci_hub_url_to_use will be the one passed around and tried first.
    # The loop for mirrors will later incorporate this and then the DEFAULT_SCI_HUB_MIRRORS list.
    # The variable 'sci_hub_base_url_for_report' used for Excel reporting will remain primary_sci_hub_url_to_use.
    sci_hub_base_url_for_report = primary_sci_hub_url_to_use
    print(f"URL de Sci-Hub principal (para primer intento y reportes): {sci_hub_base_url_for_report}")

    # 2. Solicitar al usuario que seleccione el archivo Excel o CSV
    input_file_path = filedialog.askopenfilename(
        title="Seleccionar archivo con DOIs (Excel o CSV)",
        filetypes=(("Archivos Excel", "*.xlsx *.xls"), ("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*"))
    )
    if not input_file_path:
        messagebox.showinfo("Información", "No se seleccionó ningún archivo. El programa terminará.")
        return

    # Solicitar al usuario la ubicación y el nombre del archivo ZIP
    zip_path = filedialog.asksaveasfilename(
        title="Guardar archivo ZIP como...",
        defaultextension=".zip",
        filetypes=(("Archivos ZIP", "*.zip"), ("Todos los archivos", "*.*"))
    )
    if not zip_path:
        messagebox.showinfo("Información", "No se especificó la ubicación para guardar el ZIP. El programa terminará.")
        return

    # Leer el archivo de entrada
    try:
        file_extension = os.path.splitext(input_file_path)[1].lower()
        if file_extension in ['.xlsx', '.xls']:
            try:
                df = pd.read_excel(input_file_path)
                # Standardize column names if 'doi'/'titulo' are present and 'DOI'/'Title' are not
                if 'doi' in df.columns and 'DOI' not in df.columns:
                    df.rename(columns={'doi': 'DOI'}, inplace=True)
                if 'titulo' in df.columns and 'Title' not in df.columns:
                    df.rename(columns={'titulo': 'Title'}, inplace=True)
                if not ({'DOI', 'Title'} <= set(df.columns)):
                    messagebox.showerror("Error de Excel", "El archivo Excel debe contener columnas 'DOI' y 'Title'.")
                    return
            except Exception as e:
                 messagebox.showerror("Error de Excel", f"No se pudieron leer las columnas esperadas del archivo Excel: {e}")
                 return
        elif file_extension == '.csv':
            try:
                df = pd.read_csv(input_file_path)
                if 'doi' in df.columns and 'DOI' not in df.columns:
                    df.rename(columns={'doi': 'DOI'}, inplace=True)
                if 'titulo' in df.columns and 'Title' not in df.columns:
                    df.rename(columns={'titulo': 'Title'}, inplace=True)
                if not ({'DOI', 'Title'} <= set(df.columns)):
                    messagebox.showerror("Error de CSV", "El archivo CSV debe contener columnas 'DOI' y 'Title'.")
                    return
            except Exception as e:
                 messagebox.showerror("Error de CSV", f"No se pudieron leer las columnas esperadas del archivo CSV: {e}")
                 return
        else:
            messagebox.showerror("Error de Archivo", f"Formato de archivo no soportado: {file_extension}")
            return
        
        original_input_columns = [col for col in df.columns if col not in ['DOI', 'Title']]
            
    except FileNotFoundError:
        messagebox.showerror("Error", f"No se pudo encontrar el archivo: {input_file_path}")
        return
    except Exception as e: # Captura otros errores de pandas o generales al leer
        messagebox.showerror("Error de Lectura", f"Ocurrió un error al leer el archivo: {e}")
        return

    successful_downloads = 0
    failed_downloads = [] # For the summary message box
    temp_pdf_paths = [] # To store paths of all temporary PDFs for cleanup
    total_downloaded_size_bytes = 0

    # Outer try for main operations, ensuring cleanup happens in the finally block
    try:
        # Initial ZIP Writing and Main Download Loop
        try:
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                total_articles = len(df)
                for index, row in df.iterrows():
                    original_row_data = row.to_dict()
                start_time = datetime.now()
                
                doi = str(original_row_data.get('DOI', original_row_data.get('doi', ''))).strip()
                title = str(original_row_data.get('Title', original_row_data.get('title', ''))).strip()

                detailed_status = ""
                failure_reason_for_report = ""

                if not doi:
                    print(f"Advertencia: Fila {index+2} (archivo original) ignorada por DOI vacío.")
                    failure_reason_for_report = "DOI vacío en archivo de entrada"
                    detailed_status = "Skipped_DOI_Missing"
                    # Ensure title and doi for summary are somewhat sensible even if missing
                    title_for_summary = title or "Título Desconocido"
                    doi_for_summary = "DOI Desconocido"
                    failed_downloads.append({'title': title_for_summary, 'doi': doi_for_summary, 'reason': failure_reason_for_report})
                    
                    # Log for all_articles_log
                    end_time = datetime.now()
                    log_entry = {
                        **original_row_data,
                        'Start_Time': start_time.strftime("%Y-%m-%d %H:%M:%S"),
                        'End_Time': end_time.strftime("%Y-%m-%d %H:%M:%S"),
                        'Duration_Seconds': (end_time - start_time).total_seconds(),
                        'Detailed_Status': detailed_status,
                        'Failure_Reason': failure_reason_for_report
                    }
                    all_articles_log.append(log_entry)
                    
                    # Log for failed_articles_data
                    failed_entry_data = {
                        **original_row_data,
                        'Failure_Reason': failure_reason_for_report,
                        'Detailed_Status': detailed_status
                    }
                    failed_articles_data.append(failed_entry_data)
                    continue
                
                effective_title = title if title else doi
                clean_title_for_filename = clean_filename(effective_title)
                pdf_filename_in_zip = clean_title_for_filename[:150] + ".pdf"
                
                print(f"\nProcesando artículo ({index + 1}/{total_articles}): {effective_title} (DOI: {doi})")
                
                # --- Prepare list of mirrors for this DOI ---
                mirrors_to_try_for_this_doi = []
                if primary_sci_hub_url_to_use and primary_sci_hub_url_to_use not in mirrors_to_try_for_this_doi:
                    mirrors_to_try_for_this_doi.append(primary_sci_hub_url_to_use)
                for mirror in DEFAULT_SCI_HUB_MIRRORS:
                    if mirror not in mirrors_to_try_for_this_doi:
                        mirrors_to_try_for_this_doi.append(mirror)
                
                pdf_content = None
                download_successful_this_doi = False
                successful_mirror_for_this_doi = ""
                # detailed_status and failure_reason_for_report are initialized per DOI before this loop

                for mirror_idx, current_mirror_base_url in enumerate(mirrors_to_try_for_this_doi):
                    print(f"Intentando con mirror ({mirror_idx + 1}/{len(mirrors_to_try_for_this_doi)}): {current_mirror_base_url} para DOI: {doi}")
                    
                    full_sci_hub_url_for_html_page = f"{current_mirror_base_url}{doi}"
                    
                    # Reset reasons for each mirror attempt, but keep the last one if all fail
                    current_mirror_failure_reason = "" 
                    current_mirror_detailed_status = ""

                    # --- Primary Download Attempt using extracted link from current_mirror_base_url ---
                    print(f"Intentando extraer enlace PDF desde: {full_sci_hub_url_for_html_page}")
                    actual_pdf_download_url = extract_pdf_link_from_html(full_sci_hub_url_for_html_page, session)

                    if actual_pdf_download_url:
                        print(f"Intentando descargar PDF desde enlace extraído: {actual_pdf_download_url}")
                        try:
                            response = session.get(actual_pdf_download_url, timeout=60)
                            response.raise_for_status()
                            content_type = response.headers.get('Content-Type', '').lower()
                            if 'application/pdf' in content_type:
                                pdf_content = response.content
                                current_mirror_detailed_status = f"Success_iframe_or_embed_extraction_from_{current_mirror_base_url}"
                                print(f"ÉXITO (enlace extraído con {current_mirror_base_url}): PDF obtenido para '{effective_title}'.")
                            else:
                                current_mirror_failure_reason = f"Content-Type was not application/pdf ({content_type}) from extracted link via {current_mirror_base_url}"
                                current_mirror_detailed_status = f"Failure_iframe_or_embed_extraction_not_pdf_from_{current_mirror_base_url}"
                                print(f"FALLO (enlace extraído con {current_mirror_base_url}): Contenido no es PDF. Content-Type: {content_type}")
                        except requests.exceptions.HTTPError as e:
                            current_mirror_failure_reason = f"HTTPError ({e.response.status_code}) from extracted link {actual_pdf_download_url} (via {current_mirror_base_url}): {str(e)}"
                            current_mirror_detailed_status = f"Failure_iframe_or_embed_extraction_HTTP{e.response.status_code}_from_{current_mirror_base_url}"
                            print(f"FALLO (enlace extraído con {current_mirror_base_url}): HTTP Error {e.response.status_code}. Error: {e}")
                        except requests.exceptions.RequestException as e:
                            current_mirror_failure_reason = f"RequestException from extracted link {actual_pdf_download_url} (via {current_mirror_base_url}): {str(e)}"
                            current_mirror_detailed_status = f"Failure_iframe_or_embed_extraction_RequestException_from_{current_mirror_base_url}"
                            print(f"FALLO (enlace extraído con {current_mirror_base_url}): Error de red. Error: {e}")
                        except Exception as e:
                            current_mirror_failure_reason = f"Unexpected error from extracted link {actual_pdf_download_url} (via {current_mirror_base_url}): {str(e)}"
                            current_mirror_detailed_status = f"Failure_iframe_or_embed_extraction_Unexpected_from_{current_mirror_base_url}"
                            print(f"FALLO INESPERADO (enlace extraído con {current_mirror_base_url}): Error: {e}")
                    else:
                        current_mirror_failure_reason = f"No PDF link found in HTML from {full_sci_hub_url_for_html_page}"
                        current_mirror_detailed_status = f"Failure_No_PDF_Link_Found_In_HTML_from_{current_mirror_base_url}"
                        print(f"Fallo al extraer enlace: No se encontró iframe/embed con PDF en {full_sci_hub_url_for_html_page}")

                    # --- Fallback Download Attempt (direct access to current_mirror_base_url) ---
                    if not pdf_content:
                        # Log previous attempt's failure before trying direct, if any
                        failure_reason_for_report = current_mirror_failure_reason 
                        detailed_status = current_mirror_detailed_status

                        print(f"Fallback: Intentando descargar directamente desde {full_sci_hub_url_for_html_page}")
                        try:
                            response = session.get(full_sci_hub_url_for_html_page, timeout=30)
                            response.raise_for_status()
                            content_type = response.headers.get('Content-Type', '').lower()
                            if 'application/pdf' in content_type:
                                pdf_content = response.content
                                current_mirror_detailed_status = f"Success_direct_DOI_access_fallback_from_{current_mirror_base_url}"
                                current_mirror_failure_reason = "" # Clear failure reason
                                print(f"ÉXITO (directo-fallback con {current_mirror_base_url}): PDF obtenido para '{effective_title}'.")
                            else:
                                current_mirror_failure_reason = f"Content-Type was not application/pdf ({content_type}) from direct DOI access via {current_mirror_base_url}"
                                current_mirror_detailed_status = f"Failure_direct_DOI_access_not_pdf_from_{current_mirror_base_url}"
                                print(f"FALLO (directo-fallback con {current_mirror_base_url}): Contenido no es PDF. Content-Type: {content_type}.")
                        except requests.exceptions.HTTPError as e:
                            current_mirror_failure_reason = f"HTTPError ({e.response.status_code}) from direct DOI access {full_sci_hub_url_for_html_page}: {str(e)}"
                            current_mirror_detailed_status = f"Failure_direct_DOI_access_HTTP{e.response.status_code}_from_{current_mirror_base_url}"
                            print(f"FALLO (directo-fallback con {current_mirror_base_url}): HTTP Error {e.response.status_code}. Error: {e}")
                        except requests.exceptions.RequestException as e:
                            current_mirror_failure_reason = f"RequestException from direct DOI access {full_sci_hub_url_for_html_page}: {str(e)}"
                            current_mirror_detailed_status = f"Failure_direct_DOI_access_RequestException_from_{current_mirror_base_url}"
                            print(f"FALLO (directo-fallback con {current_mirror_base_url}): Error de red. Error: {e}")
                        except Exception as e:
                            current_mirror_failure_reason = f"Unexpected error from direct DOI access {full_sci_hub_url_for_html_page}: {str(e)}"
                            current_mirror_detailed_status = f"Failure_direct_DOI_access_Unexpected_from_{current_mirror_base_url}"
                            print(f"FALLO INESPERADO (directo-fallback con {current_mirror_base_url}): Error: {e}")
                    
                    # Update final status for the DOI based on this mirror's outcome
                    failure_reason_for_report = current_mirror_failure_reason
                    detailed_status = current_mirror_detailed_status

                    if pdf_content:
                        download_successful_this_doi = True
                        successful_mirror_for_this_doi = current_mirror_base_url
                        # detailed_status is already set by the successful path
                        failure_reason_for_report = "" # Clear if successful
                        print(f"DESCARGA EXITOSA para DOI {doi} usando mirror {current_mirror_base_url}")
                        break # Exit mirror loop for this DOI
                    else:
                        print(f"FALLO con el mirror {current_mirror_base_url} para DOI {doi}. Razón: {failure_reason_for_report}")
                        if mirror_idx < len(mirrors_to_try_for_this_doi) - 1:
                            print(f"Intentando siguiente mirror en {MIRROR_SWITCH_DELAY_SECONDS} segundos...")
                            time.sleep(MIRROR_SWITCH_DELAY_SECONDS)
                        else:
                            print(f"Todos los mirrors intentados para DOI {doi} sin éxito.")
                
                end_time = datetime.now()

                # Process pdf_content if obtained (after trying all mirrors)
                if download_successful_this_doi and pdf_content:
                    data_for_successful_sheet = original_row_data.copy()
                    data_for_successful_sheet['Successful_Mirror'] = successful_mirror_for_this_doi
                    successful_articles_data.append(data_for_successful_sheet)
                    # detailed_status should reflect the successful mirror and method
                    # failure_reason_for_report is empty for success

                    temp_dir = "temp_scihub_pdfs"
                    if not os.path.exists(temp_dir):
                        os.makedirs(temp_dir)
                    temp_pdf_path = os.path.join(temp_dir, f"temp_{os.getpid()}_{index}_{pdf_filename_in_zip}")
                    
                    with open(temp_pdf_path, 'wb') as f:
                        f.write(pdf_content)
                    
                    try:
                        pdf_size_bytes = os.path.getsize(temp_pdf_path)
                        total_downloaded_size_bytes += pdf_size_bytes
                    except OSError as e:
                        print(f"Advertencia: No se pudo obtener el tamaño del archivo temporal {temp_pdf_path}: {e}")

                    zf.write(temp_pdf_path, arcname=pdf_filename_in_zip)
                    temp_pdf_paths.append(temp_pdf_path)
                    
                    print(f"AGREGADO AL ZIP: '{effective_title}' como '{pdf_filename_in_zip}'.")
                    successful_downloads += 1
                else:
                    # Both attempts failed or no PDF content found
                    # detailed_status and failure_reason_for_report should be set from the last failure
                    print(f"FALLO FINAL: No se pudo descargar '{effective_title}' (DOI: {doi}). Razón para reporte: {failure_reason_for_report}, Estado Detallado: {detailed_status}")
                    
                    failed_entry_data = {
                        **original_row_data,
                        'Failure_Reason': failure_reason_for_report,
                        'Detailed_Status': detailed_status
                    }
                    failed_articles_data.append(failed_entry_data)

                    # Populate the old failed_downloads list for the existing summary messagebox
                    title_for_summary = original_row_data.get('Title', original_row_data.get('title', effective_title))
                    doi_for_summary = original_row_data.get('DOI', original_row_data.get('doi', doi))
                    failed_downloads.append({'title': title_for_summary, 'doi': doi_for_summary, 'reason': failure_reason_for_report})

                # Populate all_articles_log for every article
                log_entry = {
                    **original_row_data,
                    'Start_Time': start_time.strftime("%Y-%m-%d %H:%M:%S"),
                    'End_Time': end_time.strftime("%Y-%m-%d %H:%M:%S"),
                    'Duration_Seconds': (end_time - start_time).total_seconds(),
                    'Detailed_Status': detailed_status,
                    'Failure_Reason': failure_reason_for_report, # COMMA ADDED HERE
                    # SciHub_Link will be added later
                    # Add successful mirror to log
                    'Successful_Mirror': successful_mirror_for_this_doi 
                }
                all_articles_log.append(log_entry)
                
                # Introduce a delay before processing the next article (regardless of success/failure of this DOI)
                print(f"Esperando {INTER_DOI_DELAY_SECONDS} segundos antes del siguiente artículo...")
                time.sleep(INTER_DOI_DELAY_SECONDS)
    
    # --- End of Main DOI Loop ---

    # --- Retry Phase for Failed Downloads ---
    if failed_articles_data: # Only run if there are failures
        print(f"\n{'='*20} INICIANDO FASE DE REINTENTO {'='*20}")
        print(f"Se reintentarán {len(failed_articles_data)} artículos fallidos.")
        
        articles_successfully_retried_ids = [] # Store DOIs of successfully retried articles
        temp_failed_articles_data_for_iteration = list(failed_articles_data) # Iterate over a copy

        for retry_idx, failed_article_entry in enumerate(temp_failed_articles_data_for_iteration):
            doi_to_retry = str(failed_article_entry.get('DOI', failed_article_entry.get('doi', ''))).strip()
            original_title_for_retry = str(failed_article_entry.get('Title', failed_article_entry.get('title', doi_to_retry))).strip()
            effective_title_for_retry = original_title_for_retry if original_title_for_retry else doi_to_retry
            
            print(f"\nReintentando artículo ({retry_idx + 1}/{len(temp_failed_articles_data_for_iteration)}): {effective_title_for_retry} (DOI: {doi_to_retry})")

            # --- Mirror Cycling Logic (copied and adapted for retry) ---
            mirrors_to_try_for_this_doi_retry = []
            if primary_sci_hub_url_to_use and primary_sci_hub_url_to_use not in mirrors_to_try_for_this_doi_retry:
                mirrors_to_try_for_this_doi_retry.append(primary_sci_hub_url_to_use)
            for mirror in DEFAULT_SCI_HUB_MIRRORS:
                if mirror not in mirrors_to_try_for_this_doi_retry:
                    mirrors_to_try_for_this_doi_retry.append(mirror)

            pdf_content_retry = None
            retry_successful_this_doi = False
            successful_mirror_for_retry = ""
            retry_detailed_status = ""
            retry_failure_reason = ""
            retry_start_time = datetime.now() # For updating duration if successful

            for mirror_idx_retry, current_mirror_base_url_retry in enumerate(mirrors_to_try_for_this_doi_retry):
                print(f"Reintento con mirror ({mirror_idx_retry + 1}/{len(mirrors_to_try_for_this_doi_retry)}): {current_mirror_base_url_retry} para DOI: {doi_to_retry}")
                full_sci_hub_url_for_html_page_retry = f"{current_mirror_base_url_retry}{doi_to_retry}"
                
                current_mirror_failure_reason_retry = ""
                current_mirror_detailed_status_retry = ""

                actual_pdf_download_url_retry = extract_pdf_link_from_html(full_sci_hub_url_for_html_page_retry, session)
                if actual_pdf_download_url_retry:
                    try:
                        response = session.get(actual_pdf_download_url_retry, timeout=60)
                        response.raise_for_status()
                        content_type = response.headers.get('Content-Type', '').lower()
                        if 'application/pdf' in content_type:
                            pdf_content_retry = response.content
                            current_mirror_detailed_status_retry = f"Success_RETRY_iframe_embed_from_{current_mirror_base_url_retry}"
                        else:
                            current_mirror_failure_reason_retry = f"RETRY: Content-Type not PDF ({content_type}) from {current_mirror_base_url_retry}"
                            current_mirror_detailed_status_retry = f"Failure_RETRY_iframe_embed_not_pdf_from_{current_mirror_base_url_retry}"
                    except requests.exceptions.RequestException as e:
                        current_mirror_failure_reason_retry = f"RETRY: RequestException from {current_mirror_base_url_retry}: {str(e)}"
                        current_mirror_detailed_status_retry = f"Failure_RETRY_iframe_embed_RequestException_from_{current_mirror_base_url_retry}"
                    # Simplified error handling for brevity in retry, can be expanded
                
                if not pdf_content_retry: # Fallback if extraction failed or link didn't yield PDF
                    # Update status from extraction attempt before trying direct
                    retry_failure_reason = current_mirror_failure_reason_retry
                    retry_detailed_status = current_mirror_detailed_status_retry
                    try:
                        response = session.get(full_sci_hub_url_for_html_page_retry, timeout=30)
                        response.raise_for_status()
                        content_type = response.headers.get('Content-Type', '').lower()
                        if 'application/pdf' in content_type:
                            pdf_content_retry = response.content
                            current_mirror_detailed_status_retry = f"Success_RETRY_direct_DOI_from_{current_mirror_base_url_retry}"
                            current_mirror_failure_reason_retry = "" 
                        else:
                            current_mirror_failure_reason_retry = f"RETRY: Content-Type not PDF ({content_type}) on direct from {current_mirror_base_url_retry}"
                            current_mirror_detailed_status_retry = f"Failure_RETRY_direct_DOI_not_pdf_from_{current_mirror_base_url_retry}"
                    except requests.exceptions.RequestException as e:
                        current_mirror_failure_reason_retry = f"RETRY: RequestException direct DOI from {current_mirror_base_url_retry}: {str(e)}"
                        current_mirror_detailed_status_retry = f"Failure_RETRY_direct_DOI_RequestException_from_{current_mirror_base_url_retry}"

                retry_failure_reason = current_mirror_failure_reason_retry
                retry_detailed_status = current_mirror_detailed_status_retry

                if pdf_content_retry:
                    retry_successful_this_doi = True
                    successful_mirror_for_retry = current_mirror_base_url_retry
                    retry_failure_reason = "" # Clear failure reason
                    print(f"ÉXITO EN REINTENTO para DOI {doi_to_retry} usando mirror {successful_mirror_for_retry}")
                    break 
                else:
                    print(f"FALLO EN REINTENTO con mirror {current_mirror_base_url_retry} para DOI {doi_to_retry}. Razón: {retry_failure_reason}")
                    if mirror_idx_retry < len(mirrors_to_try_for_this_doi_retry) - 1:
                        print(f"Intentando siguiente mirror para reintento en {MIRROR_SWITCH_DELAY_SECONDS} segundos...")
                        time.sleep(MIRROR_SWITCH_DELAY_SECONDS)
            # --- End of Mirror Cycling for Retry ---

            retry_end_time = datetime.now()

            # Update data structures based on retry outcome
            if retry_successful_this_doi and pdf_content_retry:
                successful_downloads += 1
                articles_successfully_retried_ids.append(doi_to_retry)

                # Add to successful_articles_data (original data from failed_article_entry)
                original_data_for_success = {k: v for k, v in failed_article_entry.items() if k not in ['Failure_Reason', 'Detailed_Status']}
                original_data_for_success['Successful_Mirror'] = successful_mirror_for_retry
                successful_articles_data.append(original_data_for_success)
                
                # Save the PDF
                clean_title_for_filename_retry = clean_filename(effective_title_for_retry)
                pdf_filename_in_zip_retry = clean_title_for_filename_retry[:150] + ".pdf"
                temp_dir_retry = "temp_scihub_pdfs" # Assuming zf is still open or handle zipping differently for retries
                if not os.path.exists(temp_dir_retry): os.makedirs(temp_dir_retry)
                temp_pdf_path_retry = os.path.join(temp_dir_retry, f"temp_RETRY_{os.getpid()}_{retry_idx}_{pdf_filename_in_zip_retry}")
                with open(temp_pdf_path_retry, 'wb') as f: f.write(pdf_content_retry)
                try:
                    pdf_size_bytes_retry = os.path.getsize(temp_pdf_path_retry)
                    total_downloaded_size_bytes += pdf_size_bytes_retry
                except OSError as e: print(f"Advertencia: No se pudo obtener tamaño de archivo temporal (reintento) {temp_pdf_path_retry}: {e}")
                
                # This assumes zf is still open from the main loop. This might need adjustment if zf is closed.
                # For now, let's assume it's open or we collect paths and zip later.
                # If zf is closed, we'd need to reopen in append mode or handle zipping after retries.
                # Let's collect paths and add to zip after retry loop for simplicity if zf was closed.
                # However, the original structure has 'with zipfile.ZipFile(...) as zf:' enclosing the main loop.
                # This retry logic is outside that. So, zf is CLOSED here.
                # This part needs to be re-thought: PDFs from retries cannot be added to the same zip file easily.
                # Option 1: Create a separate ZIP for retried files. (No)
                # Option 2: Don't zip retried files for now, just save them and log. (No, we'll append)
                # Option 3: Re-open the original zip in append mode ( 'a' ) (Yes)
                
                # Add the successfully retried PDF to the main ZIP file
                # pdf_filename_in_zip_retry was already defined above where temp_pdf_path_retry is.
                # effective_title_for_retry is also available.
                # pdf_filename_in_zip_for_retry = clean_filename(effective_title_for_retry)[:150] + ".pdf"
                # This was already used for temp_pdf_path_retry's name construction.

                try:
                    with zipfile.ZipFile(zip_path, 'a', zipfile.ZIP_DEFLATED) as zf_append: # Open in append mode
                        zf_append.write(temp_pdf_path_retry, arcname=pdf_filename_in_zip_retry)
                    print(f"PDF '{pdf_filename_in_zip_retry}' (obtenido en reintento) agregado al ZIP: {zip_path}")
                except Exception as e:
                    print(f"Error CRÍTICO al agregar PDF de reintento '{pdf_filename_in_zip_retry}' al ZIP: {e}")
                    # This is a more critical error, as the PDF is downloaded but not archived as expected.
                    # Consider how to alert the user more strongly or log this problem.
                    # For now, a print message will suffice for the agent's task.
                
                temp_pdf_paths.append(temp_pdf_path_retry) # Still add to general list for cleanup

                # Update all_articles_log
                for log_idx, log_entry in enumerate(all_articles_log):
                    log_doi = str(log_entry.get('DOI', log_entry.get('doi', ''))).strip()
                    if log_doi == doi_to_retry:
                        all_articles_log[log_idx]['Detailed_Status'] = retry_detailed_status
                        all_articles_log[log_idx]['Failure_Reason'] = ""
                        all_articles_log[log_idx]['Successful_Mirror'] = successful_mirror_for_retry
                        all_articles_log[log_idx]['End_Time'] = retry_end_time.strftime("%Y-%m-%d %H:%M:%S")
                        # Recalculate duration based on original start time
                        original_start_time_str = all_articles_log[log_idx]['Start_Time']
                        original_start_time_dt = datetime.strptime(original_start_time_str, "%Y-%m-%d %H:%M:%S")
                        all_articles_log[log_idx]['Duration_Seconds'] = (retry_end_time - original_start_time_dt).total_seconds()
                        break 
            else: # Retry still failed
                print(f"REINTENTO FALLIDO para DOI {doi_to_retry} después de todos los mirrors.")
                # Update all_articles_log with the latest failure reason from retry
                for log_idx, log_entry in enumerate(all_articles_log):
                    log_doi = str(log_entry.get('DOI', log_entry.get('doi', ''))).strip()
                    if log_doi == doi_to_retry:
                        all_articles_log[log_idx]['Detailed_Status'] = retry_detailed_status
                        all_articles_log[log_idx]['Failure_Reason'] = retry_failure_reason
                        # Keep original End_Time and Duration unless we want to reflect retry attempt time
                        break
                # failed_article_entry remains in failed_articles_data if not removed, its 'Failure_Reason' needs update for Excel
                for item in failed_articles_data: # Update the original list
                    item_doi = str(item.get('DOI', item.get('doi', ''))).strip()
                    if item_doi == doi_to_retry:
                        item['Failure_Reason'] = retry_failure_reason
                        item['Detailed_Status'] = retry_detailed_status
                        break
            
            # Delay between retrying different DOIs
            if retry_idx < len(temp_failed_articles_data_for_iteration) - 1:
                 print(f"Esperando {INTER_DOI_DELAY_SECONDS} segundos antes del siguiente reintento de DOI...")
                 time.sleep(INTER_DOI_DELAY_SECONDS)

            # Remove successfully retried articles from the original failed_articles_data list
            if articles_successfully_retried_ids:
                failed_articles_data = [item for item in failed_articles_data if str(item.get('DOI', item.get('doi', ''))).strip() not in articles_successfully_retried_ids]
            
            print(f"\n{'='*20} FASE DE REINTENTO COMPLETADA {'='*20}")
            # Rebuild the simple 'failed_downloads' list for the summary message
            failed_downloads = [] 
            for failed_item in failed_articles_data:
                title_for_summary = str(failed_item.get('Title', failed_item.get('title', 'Unknown Title'))).strip()
                doi_for_summary = str(failed_item.get('DOI', failed_item.get('doi', 'Unknown DOI'))).strip()
                reason_for_summary = str(failed_item.get('Failure_Reason', 'Unknown reason from retry')).strip()
                failed_downloads.append({'title': title_for_summary, 'doi': doi_for_summary, 'reason': reason_for_summary})
        # --- End of Retry Phase ---

        # --- Summary Message Composition (after retries) ---
        total_mb = total_downloaded_size_bytes / (1024 * 1024)
        summary_message = (
            f"Proceso completado.\n\n"
            f"Descargas exitosas: {successful_downloads}\n"
            f"Descargas fallidas: {len(failed_downloads)}\n" # This uses the updated failed_downloads list
            f"Tamaño total de los PDFs descargados: {total_mb:.2f} MB"
        )
        if failed_downloads: # Use the potentially updated list
            summary_message += "\n\nArtículos que no se pudieron descargar (después de reintentos):"
            for item in failed_downloads:
                summary_message += f"\n- Título: {item['title']}, DOI: {item['doi']}, Razón: {item['reason']}"
        
        print("\n" + "="*50)
        print(summary_message)
        print("="*50)
        messagebox.showinfo("Resumen de Descarga", summary_message)

        # --- Excel Report Generation ---
        generate_excel_report = messagebox.askyesno(
            "Generar Reporte Excel",
            "¿Desea generar un reporte detallado en formato Excel con los resultados?"
        )
        if generate_excel_report:
            excel_report_path = filedialog.asksaveasfilename(
                title="Guardar Reporte Excel como...",
                defaultextension=".xlsx",
                initialfile="SciHub_Descarga_Reporte.xlsx",
                filetypes=(("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*"))
            )
            if not excel_report_path:
                messagebox.showinfo("Información", "Guardado de reporte Excel cancelado por el usuario.")
                print("Reporte Excel no generado (guardado cancelado).")
            else:
                print(f"Generando reporte Excel en: {excel_report_path}")
                try:
                    # Define preferred column orders
                    obtenidos_core_cols = ['DOI', 'Title']
                    obtenidos_status_cols = ['Successful_Mirror']
                    obtenidos_link_col = ['SciHub_Link']
                    remaining_obtenidos_cols = [col for col in original_input_columns if col not in obtenid_core_cols + obtenid_status_cols]
                    ordered_obtenidos_cols = obtenid_core_cols + obtenid_status_cols + remaining_obtenidos_cols + obtenid_link_col

                    fallidos_core_cols = ['DOI', 'Title']
                    fallidos_status_cols = ['Failure_Reason', 'Detailed_Status']
                    fallidos_link_col = ['SciHub_Link']
                    remaining_fallidos_cols = [col for col in original_input_columns if col not in fallidos_core_cols + fallidos_status_cols]
                    ordered_fallidos_cols = fallidos_core_cols + remaining_fallidos_cols + fallidos_status_cols + fallidos_link_col
                    
                    tiempos_core_cols = ['DOI', 'Title']
                    tiempos_status_cols = ['Successful_Mirror', 'Start_Time', 'End_Time', 'Duration_Seconds', 'Detailed_Status', 'Failure_Reason']
                    tiempos_link_col = ['SciHub_Link']
                    remaining_tiempos_cols = [col for col in original_input_columns if col not in tiempos_core_cols + tiempos_status_cols]
                    ordered_tiempos_cols = tiempos_core_cols + remaining_tiempos_cols + tiempos_status_cols + tiempos_link_col

                    def create_ordered_df(data_list, ordered_columns_prefs):
                        df_temp = pd.DataFrame(data_list)
                        if not df_temp.empty:
                            df_temp['SciHub_Link'] = df_temp.apply(
                                lambda row: f"{sci_hub_base_url_for_report}{row.get('DOI', row.get('doi', ''))}" if pd.notna(row.get('DOI', row.get('doi', ''))) else '', 
                                axis=1
                            )
                            for col in ordered_columns_prefs:
                                if col not in df_temp.columns:
                                    df_temp[col] = pd.NA
                            final_columns = ordered_columns_prefs + [col for col in df_temp.columns if col not in ordered_columns_prefs]
                            final_columns = [col for col in final_columns if col in df_temp.columns]
                            return df_temp.reindex(columns=final_columns)
                        else:
                            return pd.DataFrame(columns=ordered_columns_prefs)

                    df_obtenidos = create_ordered_df(successful_articles_data, ordered_obtenidos_cols)
                    df_fallidos = create_ordered_df(failed_articles_data, ordered_fallidos_cols)
                    df_tiempos = create_ordered_df(all_articles_log, ordered_tiempos_cols)

                    with pd.ExcelWriter(excel_report_path, engine='openpyxl') as writer:
                        df_obtenidos.to_excel(writer, sheet_name='Obtenidos', index=False)
                        df_fallidos.to_excel(writer, sheet_name='Fallidos', index=False)
                        df_tiempos.to_excel(writer, sheet_name='Tiempos', index=False)
                    
                    messagebox.showinfo("Reporte Excel", f"Reporte Excel guardado exitosamente en: {excel_report_path}")
                    print(f"Reporte Excel guardado en: {excel_report_path}")
                except Exception as e:
                    messagebox.showerror("Error al Guardar Excel", f"No se pudo guardar el reporte Excel: {e}")
                    print(f"Error al guardar Excel: {e}")
        else:
            print("Reporte Excel no generado (usuario eligió 'No').")

    finally: # This is the new outer finally for all cleanup
        print("\n--- Limpieza Final de Archivos Temporales ---")
        for temp_path in temp_pdf_paths:
            try:
                os.remove(temp_path)
                print(f"Archivo temporal eliminado: {temp_path}")
            except OSError as e:
                print(f"Error al eliminar archivo temporal {temp_path}: {e}")
        
        temp_dir_to_check = "temp_scihub_pdfs"
        if os.path.exists(temp_dir_to_check) and not os.listdir(temp_dir_to_check):
            try:
                os.rmdir(temp_dir_to_check)
                print(f"Directorio temporal eliminado: {temp_dir_to_check}")
            except OSError as e:
                print(f"Error al eliminar el directorio temporal {temp_dir_to_check}: {e}")
        print("--- Limpieza Finalizada ---")


if __name__ == "__main__":
    # Cambié el nombre de la función principal para reflejar que ahora maneja más que solo Excel
    download_pdfs_from_file()
    print("\nScript finalizado.")
