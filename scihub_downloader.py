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

# --- Configuration Constants (Primarily for defaults now) ---
DEFAULT_SCI_HUB_MIRRORS_EXAMPLE = [ # Renamed to avoid confusion, serves as example
    "https://sci-hub.se/",
    "https://sci-hub.st/",
    "https://sci-hub.ru/",
    "https://sci-hub.red/",
    "https://NFTsci-hub.box/",
    "https://sci-hub.wf/", 
    "https://sci-hub.cat/"
]
INTER_DOI_DELAY_SECONDS = 5 # Default value
MIRROR_SWITCH_DELAY_SECONDS = 3 # Default value
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

        iframe = soup.find('iframe', id='pdf')
        if iframe and iframe.get('src'):
            pdf_src = iframe['src']
            print(f"Found PDF source in iframe: {pdf_src}")
            if pdf_src.startswith('//'):
                return 'https:' + pdf_src
            elif pdf_src.startswith('/'): 
                 return urljoin(article_page_url, pdf_src)
            return pdf_src

        embed = soup.find('embed', attrs={'type': 'application/pdf'})
        if embed and embed.get('src'):
            pdf_src = embed['src']
            print(f"Found PDF source in embed tag: {pdf_src}")
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
    root = tk.Tk()
    root.withdraw() 

    session = requests.Session()
    session.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'})

    all_articles_log = []
    successful_articles_data = []
    failed_articles_data = []
    original_input_columns = [] 

    # --- Consolidated Configuration Input ---
    print("--- Iniciando Configuración ---")

    # 1. Input File Path (Mandatory)
    input_file_path = filedialog.askopenfilename(
        title="Seleccionar archivo con DOIs (Excel o CSV)",
        filetypes=(("Archivos Excel", "*.xlsx *.xls"), ("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*"))
    )
    if not input_file_path:
        messagebox.showinfo("Información", "No se seleccionó ningún archivo de entrada. El programa terminará.")
        return

    # 2. ZIP Output File Path (Mandatory)
    zip_path = filedialog.asksaveasfilename(
        title="Guardar archivo ZIP como...",
        defaultextension=".zip",
        filetypes=(("Archivos ZIP", "*.zip"), ("Todos los archivos", "*.*"))
    )
    if not zip_path:
        messagebox.showinfo("Información", "No se especificó la ubicación para guardar el ZIP. El programa terminará.")
        return

    # 3. Excel Report Output File Path (Optional)
    excel_report_path_config = filedialog.asksaveasfilename(
        title="Guardar Reporte Excel Opcional como...",
        defaultextension=".xlsx",
        initialfile="SciHub_Descarga_Reporte.xlsx",
        filetypes=(("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*"))
    )
    if not excel_report_path_config: # User cancelled
        excel_report_path_config = "" # Ensure it's an empty string for later checks
        print("Generación de reporte Excel omitida ya que no se especificó ruta.")
    
    # 4. Delay Configurations
    user_inter_doi_delay = simpledialog.askinteger(
        "Configurar Retraso Inter-DOI",
        "Ingrese el tiempo de espera (segundos) entre cada DOI:",
        initialvalue=INTER_DOI_DELAY_SECONDS, 
        minvalue=0
    )
    if user_inter_doi_delay is None: 
        user_inter_doi_delay = INTER_DOI_DELAY_SECONDS
        messagebox.showinfo("Información", f"Retraso Inter-DOI no modificado, se usará el predeterminado: {user_inter_doi_delay}s")
    
    user_mirror_switch_delay = simpledialog.askinteger(
        "Configurar Retraso Cambio de Mirror",
        "Ingrese el tiempo de espera (segundos) al cambiar de mirror para un DOI fallido:",
        initialvalue=MIRROR_SWITCH_DELAY_SECONDS, 
        minvalue=0
    )
    if user_mirror_switch_delay is None: 
        user_mirror_switch_delay = MIRROR_SWITCH_DELAY_SECONDS
        messagebox.showinfo("Información", f"Retraso por cambio de Mirror no modificado, se usará el predeterminado: {user_mirror_switch_delay}s")

    # 5. Sci-Hub Mirrors Configuration (Mandatory User Input)
    user_mirror_list_str = simpledialog.askstring(
        "Configurar Mirrors de Sci-Hub (Obligatorio)",
        "Ingrese las URLs de los mirrors de Sci-Hub, separadas por comas.\n"
        "(ej: https://sci-hub.se/,https://sci-hub.st/)\n"
        "ESTOS SERÁN LOS ÚNICOS MIRRORS UTILIZADOS.\n"
        "Dejar vacío para usar un único default interno (https://sci-hub.se/)."
    )

    user_defined_mirrors = []
    if user_mirror_list_str is None: # User cancelled dialog
        messagebox.showerror("Configuración Requerida", "La configuración de mirrors fue cancelada. El programa terminará.")
        return
    
    if not user_mirror_list_str.strip(): # User left it empty
        messagebox.showinfo("Información de Mirrors", "No se ingresaron mirrors. Se usará el mirror por defecto: https://sci-hub.se/")
        user_defined_mirrors = ["https://sci-hub.se/"]
    else:
        raw_mirrors = [mirror.strip() for mirror in user_mirror_list_str.split(',') if mirror.strip()]
        for mirror_url in raw_mirrors:
            if not mirror_url.startswith(("http://", "https://")):
                mirror_url = "https://" + mirror_url # Assume https if no scheme
            if not mirror_url.endswith('/'):
                mirror_url += '/'
            user_defined_mirrors.append(mirror_url)
        
        if not user_defined_mirrors:
            messagebox.showerror("Error de Configuración", "La lista de mirrors ingresada está vacía o es inválida. Se requiere al menos una URL de mirror. El programa terminará.")
            return

    sci_hub_base_url_for_report = user_defined_mirrors[0] # For Excel links, use the first user-provided mirror

    print("\n--- Configuración Aplicada ---")
    print(f"Archivo de entrada: {input_file_path}")
    print(f"Archivo ZIP de salida: {zip_path}")
    if excel_report_path_config:
        print(f"Archivo de reporte Excel: {excel_report_path_config}")
    else:
        print("Reporte Excel: No se generará (ruta no especificada).")
    print(f"Retraso Inter-DOI: {user_inter_doi_delay}s")
    print(f"Retraso Cambio de Mirror: {user_mirror_switch_delay}s")
    print(f"Mirrors Sci-Hub a utilizar: {', '.join(user_defined_mirrors)}")
    print("----------------------------\n")
    # --- End of Consolidated Configuration Input ---

    try:
        file_extension = os.path.splitext(input_file_path)[1].lower()
        if file_extension in ['.xlsx', '.xls']:
            try:
                df = pd.read_excel(input_file_path)
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
    except Exception as e: 
        messagebox.showerror("Error de Lectura", f"Ocurrió un error al leer el archivo: {e}")
        return

    successful_downloads = 0
    failed_downloads_summary_list = [] # Renamed for clarity, for the summary message box
    temp_pdf_paths = [] 
    total_downloaded_size_bytes = 0

    try:
        zip_creation_or_main_loop_error = False 
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
                        title_for_summary = title or "Título Desconocido"
                        doi_for_summary = "DOI Desconocido"
                        # failed_downloads_summary_list will be populated later from failed_articles_data
                        
                        end_time = datetime.now()
                        log_entry = {**original_row_data, 'Start_Time': start_time.strftime("%Y-%m-%d %H:%M:%S"), 'End_Time': end_time.strftime("%Y-%m-%d %H:%M:%S"), 'Duration_Seconds': (end_time - start_time).total_seconds(), 'Detailed_Status': detailed_status, 'Failure_Reason': failure_reason_for_report}
                        all_articles_log.append(log_entry)
                        failed_articles_data.append({**original_row_data, 'Failure_Reason': failure_reason_for_report, 'Detailed_Status': detailed_status})
                        continue
                    
                    effective_title = title if title else doi
                    clean_title_for_filename = clean_filename(effective_title)
                    pdf_filename_in_zip = clean_title_for_filename[:150] + ".pdf"
                    
                    print(f"\nProcesando artículo ({index + 1}/{total_articles}): {effective_title} (DOI: {doi})")
                    
                    # THIS IS WHERE user_defined_mirrors WILL BE USED IN THE NEXT STEP
                    mirrors_to_try_for_this_doi = list(user_defined_mirrors) # Use a copy
                                        
                    pdf_content = None
                    download_successful_this_doi = False
                    successful_mirror_for_this_doi = ""

                    for mirror_idx, current_mirror_base_url in enumerate(mirrors_to_try_for_this_doi):
                        print(f"Intentando con mirror ({mirror_idx + 1}/{len(mirrors_to_try_for_this_doi)}): {current_mirror_base_url} para DOI: {doi}")
                        full_sci_hub_url_for_html_page = f"{current_mirror_base_url}{doi}"
                        current_mirror_failure_reason = "" 
                        current_mirror_detailed_status = ""

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
                            except requests.exceptions.HTTPError as e:
                                current_mirror_failure_reason = f"HTTPError ({e.response.status_code}) from extracted link ... (via {current_mirror_base_url})"
                                current_mirror_detailed_status = f"Failure_iframe_or_embed_extraction_HTTP{e.response.status_code}_from_{current_mirror_base_url}"
                            except requests.exceptions.RequestException as e:
                                current_mirror_failure_reason = f"RequestException from extracted link ... (via {current_mirror_base_url})"
                                current_mirror_detailed_status = f"Failure_iframe_or_embed_extraction_RequestException_from_{current_mirror_base_url}"
                            except Exception as e:
                                current_mirror_failure_reason = f"Unexpected error from extracted link ... (via {current_mirror_base_url})"
                                current_mirror_detailed_status = f"Failure_iframe_or_embed_extraction_Unexpected_from_{current_mirror_base_url}"
                        else:
                            current_mirror_failure_reason = f"No PDF link found in HTML from {full_sci_hub_url_for_html_page}"
                            current_mirror_detailed_status = f"Failure_No_PDF_Link_Found_In_HTML_from_{current_mirror_base_url}"
                        
                        if not pdf_content:
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
                                    current_mirror_failure_reason = ""
                                else:
                                    current_mirror_failure_reason = f"Content-Type was not application/pdf ({content_type}) from direct DOI ... via {current_mirror_base_url}"
                                    current_mirror_detailed_status = f"Failure_direct_DOI_access_not_pdf_from_{current_mirror_base_url}"
                            except requests.exceptions.HTTPError as e:
                                current_mirror_failure_reason = f"HTTPError ({e.response.status_code}) from direct DOI ... (via {current_mirror_base_url})"
                                current_mirror_detailed_status = f"Failure_direct_DOI_access_HTTP{e.response.status_code}_from_{current_mirror_base_url}"
                            except requests.exceptions.RequestException as e:
                                current_mirror_failure_reason = f"RequestException from direct DOI ... (via {current_mirror_base_url})"
                                current_mirror_detailed_status = f"Failure_direct_DOI_access_RequestException_from_{current_mirror_base_url}"
                            except Exception as e:
                                current_mirror_failure_reason = f"Unexpected error from direct DOI ... (via {current_mirror_base_url})"
                                current_mirror_detailed_status = f"Failure_direct_DOI_access_Unexpected_from_{current_mirror_base_url}"
                        
                        failure_reason_for_report = current_mirror_failure_reason
                        detailed_status = current_mirror_detailed_status

                        if pdf_content:
                            download_successful_this_doi = True
                            successful_mirror_for_this_doi = current_mirror_base_url
                            failure_reason_for_report = "" 
                            print(f"DESCARGA EXITOSA para DOI {doi} usando mirror {current_mirror_base_url}")
                            break 
                        else:
                            print(f"FALLO con el mirror {current_mirror_base_url} para DOI {doi}. Razón: {failure_reason_for_report}")
                            if mirror_idx < len(mirrors_to_try_for_this_doi) - 1:
                                print(f"Intentando siguiente mirror en {user_mirror_switch_delay} segundos...")
                                time.sleep(user_mirror_switch_delay)
                            else:
                                print(f"Todos los mirrors intentados para DOI {doi} sin éxito.")
                    
                    end_time = datetime.now()

                    if download_successful_this_doi and pdf_content:
                        data_for_successful_sheet = original_row_data.copy()
                        data_for_successful_sheet['Successful_Mirror'] = successful_mirror_for_this_doi
                        successful_articles_data.append(data_for_successful_sheet)
                        temp_dir = "temp_scihub_pdfs"
                        if not os.path.exists(temp_dir): os.makedirs(temp_dir)
                        temp_pdf_path = os.path.join(temp_dir, f"temp_{os.getpid()}_{index}_{pdf_filename_in_zip}")
                        with open(temp_pdf_path, 'wb') as f: f.write(pdf_content)
                        try:
                            pdf_size_bytes = os.path.getsize(temp_pdf_path)
                            total_downloaded_size_bytes += pdf_size_bytes
                        except OSError as e: print(f"Advertencia: No se pudo obtener el tamaño del archivo temporal {temp_pdf_path}: {e}")
                        zf.write(temp_pdf_path, arcname=pdf_filename_in_zip)
                        temp_pdf_paths.append(temp_pdf_path)
                        print(f"AGREGADO AL ZIP: '{effective_title}' como '{pdf_filename_in_zip}'.")
                        successful_downloads += 1
                    else:
                        print(f"FALLO FINAL: No se pudo descargar '{effective_title}' (DOI: {doi}). Razón para reporte: {failure_reason_for_report}, Estado Detallado: {detailed_status}")
                        failed_entry_data = {**original_row_data, 'Failure_Reason': failure_reason_for_report, 'Detailed_Status': detailed_status}
                        failed_articles_data.append(failed_entry_data)
                        # failed_downloads_summary_list will be populated later

                    log_entry = {**original_row_data, 'Start_Time': start_time.strftime("%Y-%m-%d %H:%M:%S"), 'End_Time': end_time.strftime("%Y-%m-%d %H:%M:%S"), 'Duration_Seconds': (end_time - start_time).total_seconds(), 'Detailed_Status': detailed_status, 'Failure_Reason': failure_reason_for_report, 'Successful_Mirror': successful_mirror_for_this_doi }
                    all_articles_log.append(log_entry)
                    
                    print(f"Esperando {user_inter_doi_delay} segundos antes del siguiente artículo...")
                    time.sleep(user_inter_doi_delay)
        except FileNotFoundError: 
            messagebox.showerror("Error", f"No se pudo crear el archivo ZIP en la ruta especificada (Directorio no encontrado): {zip_path}")
            print("Error crítico: FileNotFoundError al crear ZIP inicial. Terminando.")
            zip_creation_or_main_loop_error = True 
        except Exception as e: 
            messagebox.showerror("Error", f"Ocurrió un error inesperado durante la creación del ZIP o el primer paso de descargas: {e}")
            print(f"Error crítico: Excepción durante ZIP inicial o descargas: {e}. Terminando.")
            zip_creation_or_main_loop_error = True

        if not zip_creation_or_main_loop_error:
            if failed_articles_data: 
                print(f"\n{'='*20} INICIANDO FASE DE REINTENTO {'='*20}")
                print(f"Se reintentarán {len(failed_articles_data)} artículos fallidos.")
            
            articles_successfully_retried_ids = [] 
            temp_failed_articles_data_for_iteration = list(failed_articles_data) 

            # THIS IS WHERE user_defined_mirrors WILL BE USED IN THE NEXT STEP for retry loop
            mirrors_for_retry = list(user_defined_mirrors) # Use a copy

            for retry_idx, failed_article_entry in enumerate(temp_failed_articles_data_for_iteration):
                doi_to_retry = str(failed_article_entry.get('DOI', failed_article_entry.get('doi', ''))).strip()
                original_title_for_retry = str(failed_article_entry.get('Title', failed_article_entry.get('title', doi_to_retry))).strip()
                effective_title_for_retry = original_title_for_retry if original_title_for_retry else doi_to_retry
                pdf_filename_in_zip_retry = clean_filename(effective_title_for_retry)[:150] + ".pdf" # Define here for zipping
                
                print(f"\nReintentando artículo ({retry_idx + 1}/{len(temp_failed_articles_data_for_iteration)}): {effective_title_for_retry} (DOI: {doi_to_retry})")
                
                pdf_content_retry = None
                retry_successful_this_doi = False
                successful_mirror_for_retry = ""
                retry_detailed_status = ""
                retry_failure_reason = ""
                retry_start_time = datetime.now() 

                for mirror_idx_retry, current_mirror_base_url_retry in enumerate(mirrors_for_retry): # Use parsed list
                    print(f"Reintento con mirror ({mirror_idx_retry + 1}/{len(mirrors_for_retry)}): {current_mirror_base_url_retry} para DOI: {doi_to_retry}")
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
                    
                    if not pdf_content_retry: 
                        retry_failure_reason = current_mirror_failure_reason_retry # Keep reason from extraction if it happened
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
                        retry_failure_reason = "" 
                        print(f"ÉXITO EN REINTENTO para DOI {doi_to_retry} usando mirror {successful_mirror_for_retry}")
                        break 
                    else:
                        print(f"FALLO EN REINTENTO con mirror {current_mirror_base_url_retry} para DOI {doi_to_retry}. Razón: {retry_failure_reason}")
                        if mirror_idx_retry < len(mirrors_for_retry) - 1: # Use current retry mirror list
                            print(f"Intentando siguiente mirror para reintento en {user_mirror_switch_delay} segundos...")
                            time.sleep(user_mirror_switch_delay)
                
                retry_end_time = datetime.now()

                if retry_successful_this_doi and pdf_content_retry:
                    successful_downloads += 1
                    articles_successfully_retried_ids.append(doi_to_retry)
                    original_data_for_success = {k: v for k, v in failed_article_entry.items() if k not in ['Failure_Reason', 'Detailed_Status']}
                    original_data_for_success['Successful_Mirror'] = successful_mirror_for_retry
                    successful_articles_data.append(original_data_for_success)
                    
                    temp_dir_retry = "temp_scihub_pdfs"
                    if not os.path.exists(temp_dir_retry): os.makedirs(temp_dir_retry)
                    temp_pdf_path_retry = os.path.join(temp_dir_retry, f"temp_RETRY_{os.getpid()}_{retry_idx}_{pdf_filename_in_zip_retry}")
                    with open(temp_pdf_path_retry, 'wb') as f: f.write(pdf_content_retry)
                    try:
                        pdf_size_bytes_retry = os.path.getsize(temp_pdf_path_retry)
                        total_downloaded_size_bytes += pdf_size_bytes_retry
                    except OSError as e: print(f"Advertencia: No se pudo obtener tamaño de archivo temporal (reintento) {temp_pdf_path_retry}: {e}")
                    
                    try:
                        with zipfile.ZipFile(zip_path, 'a', zipfile.ZIP_DEFLATED) as zf_append: 
                            zf_append.write(temp_pdf_path_retry, arcname=pdf_filename_in_zip_retry)
                        print(f"PDF '{pdf_filename_in_zip_retry}' (obtenido en reintento) agregado al ZIP: {zip_path}")
                    except Exception as e:
                        print(f"Error CRÍTICO al agregar PDF de reintento '{pdf_filename_in_zip_retry}' al ZIP: {e}")
                    temp_pdf_paths.append(temp_pdf_path_retry) 

                    for log_idx, log_entry in enumerate(all_articles_log):
                        log_doi = str(log_entry.get('DOI', log_entry.get('doi', ''))).strip()
                        if log_doi == doi_to_retry:
                            all_articles_log[log_idx]['Detailed_Status'] = retry_detailed_status
                            all_articles_log[log_idx]['Failure_Reason'] = ""
                            all_articles_log[log_idx]['Successful_Mirror'] = successful_mirror_for_retry
                            all_articles_log[log_idx]['End_Time'] = retry_end_time.strftime("%Y-%m-%d %H:%M:%S")
                            original_start_time_str = all_articles_log[log_idx]['Start_Time']
                            original_start_time_dt = datetime.strptime(original_start_time_str, "%Y-%m-%d %H:%M:%S")
                            all_articles_log[log_idx]['Duration_Seconds'] = (retry_end_time - original_start_time_dt).total_seconds()
                            break 
                else: 
                    print(f"REINTENTO FALLIDO para DOI {doi_to_retry} después de todos los mirrors.")
                    for log_idx, log_entry in enumerate(all_articles_log):
                        log_doi = str(log_entry.get('DOI', log_entry.get('doi', ''))).strip()
                        if log_doi == doi_to_retry:
                            all_articles_log[log_idx]['Detailed_Status'] = retry_detailed_status
                            all_articles_log[log_idx]['Failure_Reason'] = retry_failure_reason
                            break
                    for item in failed_articles_data: 
                        item_doi = str(item.get('DOI', item.get('doi', ''))).strip()
                        if item_doi == doi_to_retry:
                            item['Failure_Reason'] = retry_failure_reason
                            item['Detailed_Status'] = retry_detailed_status
                            break
                
                if retry_idx < len(temp_failed_articles_data_for_iteration) - 1:
                     print(f"Esperando {user_inter_doi_delay} segundos antes del siguiente reintento de DOI...")
                     time.sleep(user_inter_doi_delay)

            if articles_successfully_retried_ids:
                failed_articles_data = [item for item in failed_articles_data if str(item.get('DOI', item.get('doi', ''))).strip() not in articles_successfully_retried_ids]
            
            print(f"\n{'='*20} FASE DE REINTENTO COMPLETADA {'='*20}")
            
            failed_downloads_summary_list = [] 
            for failed_item in failed_articles_data:
                title_for_summary = str(failed_item.get('Title', failed_item.get('title', 'Unknown Title'))).strip()
                doi_for_summary = str(failed_item.get('DOI', failed_item.get('doi', 'Unknown DOI'))).strip()
                reason_for_summary = str(failed_item.get('Failure_Reason', 'Unknown reason from retry')).strip()
                failed_downloads_summary_list.append({'title': title_for_summary, 'doi': doi_for_summary, 'reason': reason_for_summary})

            total_mb = total_downloaded_size_bytes / (1024 * 1024)
            summary_message = (
                f"Proceso completado.\n\n"
                f"Descargas exitosas: {successful_downloads}\n"
                f"Descargas fallidas: {len(failed_downloads_summary_list)}\n" 
                f"Tamaño total de los PDFs descargados: {total_mb:.2f} MB"
            )
            if failed_downloads_summary_list: 
                summary_message += "\n\nArtículos que no se pudieron descargar (después de reintentos):"
                for item in failed_downloads_summary_list:
                    summary_message += f"\n- Título: {item['title']}, DOI: {item['doi']}, Razón: {item['reason']}"
            
            print("\n" + "="*50)
            print(summary_message)
            print("="*50)
            messagebox.showinfo("Resumen de Descarga", summary_message)

            # --- Excel Report Generation ---
            # Ask user if they want to generate a report at all
            generate_excel_report_prompt = messagebox.askyesno(
                "Generar Reporte Excel",
                "¿Desea generar un reporte detallado en formato Excel con los resultados?"
            )

            if generate_excel_report_prompt:
                excel_report_path_to_use = ""
                if excel_report_path_config: # Path was provided during initial config
                    # Confirm usage of pre-configured path
                    confirm_preconfigured_path = messagebox.askyesno(
                        "Confirmar Ruta de Reporte",
                        f"Se configuró guardar el reporte Excel en:\n{excel_report_path_config}\n\n¿Desea usar esta ruta?"
                    )
                    if confirm_preconfigured_path:
                        excel_report_path_to_use = excel_report_path_config
                    else: # User does not want to use pre-configured path, ask for a new one
                        excel_report_path_to_use = filedialog.asksaveasfilename(
                            title="Guardar Reporte Excel como...",
                            defaultextension=".xlsx",
                            initialfile="SciHub_Descarga_Reporte_Alternativo.xlsx",
                            filetypes=(("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*"))
                        )
                else: # No path was configured initially, ask for one now
                    excel_report_path_to_use = filedialog.asksaveasfilename(
                        title="Guardar Reporte Excel como...",
                        defaultextension=".xlsx",
                        initialfile="SciHub_Descarga_Reporte.xlsx",
                        filetypes=(("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*"))
                    )

                if excel_report_path_to_use: # If a path was determined (either pre-configured and confirmed, or newly selected)
                    print(f"Generando reporte Excel en: {excel_report_path_to_use}")
                    try:
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
                else: # User chose not to generate Excel here, even if path was set
                    print("Generación de reporte Excel omitida por el usuario en este punto.")
            elif not excel_report_path_config : # Path was not set during initial config
                 print("No se solicitó generar reporte Excel (ruta no especificada en la configuración inicial).")
        elif zip_creation_or_main_loop_error:
             print("Proceso interrumpido debido a un error crítico durante la fase inicial. No se generará resumen ni reporte Excel.")

    finally: 
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
    download_pdfs_from_file()
    print("\nScript finalizado.")
