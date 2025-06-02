import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import tkinter.scrolledtext as st 
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

# --- Configuration Constants (Primarily for defaults now) ---
DEFAULT_SCI_HUB_MIRRORS_EXAMPLE = [ 
    "https://sci-hub.se/", "https://sci-hub.st/", "https://sci-hub.ru/",
    "https://sci-hub.red/", "https://NFTsci-hub.box/", "https://sci-hub.wf/", 
    "https://sci-hub.cat/"
]
INTER_DOI_DELAY_SECONDS = 5 
MIRROR_SWITCH_DELAY_SECONDS = 3
# --- End Configuration Constants ---

class TextRedirector(object):
    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag
    def write(self, str_):
        self.widget.configure(state='normal')
        self.widget.insert(tk.END, str_, (self.tag,))
        self.widget.see(tk.END)
        self.widget.configure(state='disabled')
    def flush(self):
        self.widget.update_idletasks() 

def on_log_window_close(log_window_ref, original_stdout_ref, root_tk_instance):
    if sys.stdout != original_stdout_ref: 
        print("Log window closed by user. Restoring original stdout.", file=original_stdout_ref) 
        sys.stdout = original_stdout_ref
    if log_window_ref:
        try:
            log_window_ref.destroy()
        except tk.TclError: 
            pass

def clean_filename(title):
    return re.sub(r'[\\/*?:"<>|]', '_', title)

def extract_pdf_link_from_html(article_page_url, session):
    print(f"Fetching HTML page: {article_page_url}") 
    try:
        response = session.get(article_page_url, timeout=30)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, 'html.parser')
        iframe = soup.find('iframe', id='pdf')
        if iframe and iframe.get('src'):
            pdf_src = iframe['src']
            print(f"Found PDF source in iframe: {pdf_src}") 
            if pdf_src.startswith('//'): return 'https:' + pdf_src
            elif pdf_src.startswith('/'): return urljoin(article_page_url, pdf_src)
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

def print_to_console(message, orig_stdout):
    print(message, file=orig_stdout)

def download_pdfs_from_file():
    original_stdout = sys.stdout 
    root = tk.Tk()
    root.withdraw() 
    
    log_window = None 
    log_text_widget = None

    try:
        print_to_console("DIAG: --- Iniciando Configuración ---", original_stdout)

        print_to_console("DIAG: Before input_file_path dialog", original_stdout)
        input_file_path = filedialog.askopenfilename(title="Seleccionar archivo con DOIs (Excel o CSV)", filetypes=(("Archivos Excel", "*.xlsx *.xls"), ("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*")))
        print_to_console(f"DIAG: After input_file_path dialog, path: {input_file_path}", original_stdout)
        if not input_file_path: messagebox.showinfo("Información", "No se seleccionó ningún archivo de entrada. El programa terminará."); return

        print_to_console("DIAG: Before zip_path dialog", original_stdout)
        zip_path = filedialog.asksaveasfilename(title="Guardar archivo ZIP como...", defaultextension=".zip", filetypes=(("Archivos ZIP", "*.zip"), ("Todos los archivos", "*.*")))
        print_to_console(f"DIAG: After zip_path dialog, path: {zip_path}", original_stdout)
        if not zip_path: messagebox.showinfo("Información", "No se especificó la ubicación para guardar el ZIP. El programa terminará."); return

        print_to_console("DIAG: Before excel_report_path_config dialog", original_stdout)
        excel_report_path_config = filedialog.asksaveasfilename(title="Guardar Reporte Excel Opcional como...", defaultextension=".xlsx", initialfile="SciHub_Descarga_Reporte.xlsx", filetypes=(("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")))
        print_to_console(f"DIAG: After excel_report_path_config dialog, path: {excel_report_path_config}", original_stdout)
        if not excel_report_path_config: excel_report_path_config = ""; print_to_console("DIAG: Ruta para reporte Excel no especificada.", original_stdout)
        
        print_to_console("DIAG: Before user_inter_doi_delay dialog", original_stdout)
        user_inter_doi_delay = simpledialog.askinteger("Configurar Retraso Inter-DOI", "Ingrese el tiempo de espera (segundos) entre cada DOI:", initialvalue=INTER_DOI_DELAY_SECONDS, minvalue=0 )
        print_to_console(f"DIAG: After user_inter_doi_delay dialog, value: {user_inter_doi_delay}", original_stdout)
        if user_inter_doi_delay is None: user_inter_doi_delay = INTER_DOI_DELAY_SECONDS; messagebox.showinfo("Información", f"Retraso Inter-DOI no modificado, se usará el predeterminado: {user_inter_doi_delay}s")
        
        print_to_console("DIAG: Before user_mirror_switch_delay dialog", original_stdout)
        user_mirror_switch_delay = simpledialog.askinteger("Configurar Retraso Cambio de Mirror", "Ingrese el tiempo de espera (segundos) al cambiar de mirror:", initialvalue=MIRROR_SWITCH_DELAY_SECONDS, minvalue=0)
        print_to_console(f"DIAG: After user_mirror_switch_delay dialog, value: {user_mirror_switch_delay}", original_stdout)
        if user_mirror_switch_delay is None: user_mirror_switch_delay = MIRROR_SWITCH_DELAY_SECONDS; messagebox.showinfo("Información", f"Retraso por cambio de Mirror no modificado, se usará el predeterminado: {user_mirror_switch_delay}s")

        print_to_console("DIAG: Before user_mirror_list_str dialog", original_stdout)
        user_mirror_list_str = simpledialog.askstring("Configurar Mirrors de Sci-Hub (Obligatorio)", "Ingrese URLs de mirrors Sci-Hub, separadas por comas.\n(ej: https://sci-hub.se/,https://sci-hub.st/)\nESTOS SERÁN LOS ÚNICOS MIRRORS UTILIZADOS.\nDejar vacío para usar default (https://sci-hub.se/).")
        print_to_console(f"DIAG: After user_mirror_list_str dialog, value: {user_mirror_list_str}", original_stdout)
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
        
        # --- Process Tkinter events after dialogs ---
        print_to_console("DIAG: Before root.update() after all dialogs.", original_stdout)
        root.update()
        print_to_console("DIAG: After root.update() after all dialogs.", original_stdout)
        # --- End event processing ---

        print_to_console("DIAG: All dialogs complete. Before log_window creation.", original_stdout)
        log_window = tk.Toplevel(root)
        print_to_console("DIAG: After log_window = tk.Toplevel(root)", original_stdout)
        log_window.title("Log de Proceso de Descarga")
        log_window.geometry("800x600")
        
        print_to_console("DIAG: Before log_text_widget creation", original_stdout)
        log_text_widget = st.ScrolledText(log_window, wrap=tk.WORD, state='disabled')
        log_text_widget.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        print_to_console("DIAG: After log_text_widget creation", original_stdout)
        
        log_window.protocol("WM_DELETE_WINDOW", lambda: on_log_window_close(log_window, original_stdout, root))
        
        print_to_console("DIAG: Before sys.stdout redirection", original_stdout)
        sys.stdout = TextRedirector(log_text_widget) 
        print_to_console("DIAG: After sys.stdout redirection. Subsequent app prints go to GUI.", original_stdout) 
        
        print("\n--- Configuración Aplicada ---")
        print(f"Archivo de entrada: {input_file_path}")
        print(f"Archivo ZIP de salida: {zip_path}")
        if excel_report_path_config: print(f"Archivo de reporte Excel: {excel_report_path_config}")
        else: print("Reporte Excel: No se generará (ruta no especificada).")
        print(f"Retraso Inter-DOI: {user_inter_doi_delay}s")
        print(f"Retraso Cambio de Mirror: {user_mirror_switch_delay}s")
        print(f"Mirrors Sci-Hub a utilizar: {', '.join(user_defined_mirrors)}")
        print("-----------------------------------------------------\n")
        if log_window and log_window.winfo_exists(): log_window.update_idletasks()

        session = requests.Session()
        session.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'})
        all_articles_log = []; successful_articles_data = []; failed_articles_data = []; original_input_columns = [] 
        
        try:
            file_extension = os.path.splitext(input_file_path)[1].lower()
            if file_extension in ['.xlsx', '.xls']:
                try:
                    df = pd.read_excel(input_file_path); 
                    if 'doi' in df.columns and 'DOI' not in df.columns: df.rename(columns={'doi': 'DOI'}, inplace=True)
                    if 'titulo' in df.columns and 'Title' not in df.columns: df.rename(columns={'titulo': 'Title'}, inplace=True)
                    if not ({'DOI', 'Title'} <= set(df.columns)): messagebox.showerror("Error de Excel", "Archivo Excel debe contener 'DOI' y 'Title'."); raise Exception("Missing columns")
                except Exception as e: messagebox.showerror("Error de Excel", f"Error al leer Excel: {e}"); raise
            elif file_extension == '.csv':
                try:
                    df = pd.read_csv(input_file_path); 
                    if 'doi' in df.columns and 'DOI' not in df.columns: df.rename(columns={'doi': 'DOI'}, inplace=True)
                    if 'titulo' in df.columns and 'Title' not in df.columns: df.rename(columns={'titulo': 'Title'}, inplace=True)
                    if not ({'DOI', 'Title'} <= set(df.columns)): messagebox.showerror("Error de CSV", "Archivo CSV debe contener 'DOI' y 'Title'."); raise Exception("Missing columns")
                except Exception as e: messagebox.showerror("Error de CSV", f"Error al leer CSV: {e}"); raise
            else: messagebox.showerror("Error de Archivo", f"Formato no soportado: {file_extension}"); raise Exception("Unsupported format")
            original_input_columns = [col for col in df.columns if col not in ['DOI', 'Title']]
        except FileNotFoundError: messagebox.showerror("Error", f"No se encontró archivo: {input_file_path}"); raise
        except Exception as e: messagebox.showerror("Error de Lectura", f"Error al leer archivo: {e}"); raise

        successful_downloads = 0; failed_downloads_summary_list = []; temp_pdf_paths = []; total_downloaded_size_bytes = 0
        zip_creation_or_main_loop_error = False 

        try:
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                total_articles = len(df)
                for index, row in df.iterrows():
                    original_row_data = row.to_dict(); start_time = datetime.now()
                    doi = str(original_row_data.get('DOI', original_row_data.get('doi', ''))).strip()
                    title = str(original_row_data.get('Title', original_row_data.get('title', ''))).strip()
                    detailed_status = ""; failure_reason_for_report = ""
                    if not doi:
                        print(f"Advertencia: Fila {index+2} ignorada por DOI vacío.") 
                        failure_reason_for_report = "DOI vacío"; detailed_status = "Skipped_DOI_Missing"; end_time = datetime.now()
                        log_entry = {**original_row_data, 'Start_Time': start_time.strftime("%Y-%m-%d %H:%M:%S"), 'End_Time': end_time.strftime("%Y-%m-%d %H:%M:%S"), 'Duration_Seconds': (end_time - start_time).total_seconds(), 'Detailed_Status': detailed_status, 'Failure_Reason': failure_reason_for_report, 'Successful_Mirror': ""}
                        all_articles_log.append(log_entry); failed_articles_data.append({**original_row_data, 'Failure_Reason': failure_reason_for_report, 'Detailed_Status': detailed_status}); continue
                    effective_title = title if title else doi; clean_title_for_filename = clean_filename(effective_title); pdf_filename_in_zip = clean_title_for_filename[:150] + ".pdf"
                    print(f"\nProcesando artículo ({index + 1}/{total_articles}): {effective_title} (DOI: {doi})")
                    mirrors_to_try_for_this_doi = list(user_defined_mirrors)
                    pdf_content = None; download_successful_this_doi = False; successful_mirror_for_this_doi = ""
                    for mirror_idx, current_mirror_base_url in enumerate(mirrors_to_try_for_this_doi):
                        print(f"Intentando con mirror ({mirror_idx + 1}/{len(mirrors_to_try_for_this_doi)}): {current_mirror_base_url} para DOI: {doi}")
                        full_sci_hub_url_for_html_page = f"{current_mirror_base_url}{doi}"
                        current_mirror_failure_reason = ""; current_mirror_detailed_status = ""
                        actual_pdf_download_url = extract_pdf_link_from_html(full_sci_hub_url_for_html_page, session)
                        if actual_pdf_download_url:
                            # ... (extraction try-except block)
                            try:
                                response = session.get(actual_pdf_download_url, timeout=60); response.raise_for_status()
                                content_type = response.headers.get('Content-Type', '').lower()
                                if 'application/pdf' in content_type: pdf_content = response.content; current_mirror_detailed_status = f"Success_iframe_or_embed_extraction_from_{current_mirror_base_url}"
                                else: current_mirror_failure_reason = f"Content-Type not PDF ({content_type}) from extracted link via {current_mirror_base_url}"; current_mirror_detailed_status = f"Failure_iframe_or_embed_extraction_not_pdf_from_{current_mirror_base_url}"
                            except requests.exceptions.HTTPError as e: current_mirror_failure_reason = f"HTTPError ({e.response.status_code}) ..."; current_mirror_detailed_status = f"Failure_iframe_or_embed_extraction_HTTP{e.response.status_code}_from_{current_mirror_base_url}"
                            except requests.exceptions.RequestException as e: current_mirror_failure_reason = f"RequestException ..."; current_mirror_detailed_status = f"Failure_iframe_or_embed_extraction_RequestException_from_{current_mirror_base_url}"
                            except Exception as e: current_mirror_failure_reason = f"Unexpected error ..."; current_mirror_detailed_status = f"Failure_iframe_or_embed_extraction_Unexpected_from_{current_mirror_base_url}"

                        else: current_mirror_failure_reason = f"No PDF link in HTML from {full_sci_hub_url_for_html_page}"; current_mirror_detailed_status = f"Failure_No_PDF_Link_Found_In_HTML_from_{current_mirror_base_url}"
                        if not pdf_content:
                            failure_reason_for_report = current_mirror_failure_reason; detailed_status = current_mirror_detailed_status
                            # ... (fallback try-except block)
                            try:
                                response = session.get(full_sci_hub_url_for_html_page, timeout=30); response.raise_for_status()
                                content_type = response.headers.get('Content-Type', '').lower()
                                if 'application/pdf' in content_type: pdf_content = response.content; current_mirror_detailed_status = f"Success_direct_DOI_access_fallback_from_{current_mirror_base_url}"; current_mirror_failure_reason = ""
                                else: current_mirror_failure_reason = f"Content-Type not PDF ({content_type}) from direct DOI ..."; current_mirror_detailed_status = f"Failure_direct_DOI_access_not_pdf_from_{current_mirror_base_url}"
                            except requests.exceptions.HTTPError as e: current_mirror_failure_reason = f"HTTPError ({e.response.status_code}) from direct DOI ..."; current_mirror_detailed_status = f"Failure_direct_DOI_access_HTTP{e.response.status_code}_from_{current_mirror_base_url}"
                            except requests.exceptions.RequestException as e: current_mirror_failure_reason = f"RequestException from direct DOI ..."; current_mirror_detailed_status = f"Failure_direct_DOI_access_RequestException_from_{current_mirror_base_url}"
                            except Exception as e: current_mirror_failure_reason = f"Unexpected error from direct DOI ..."; current_mirror_detailed_status = f"Failure_direct_DOI_access_Unexpected_from_{current_mirror_base_url}"
                        failure_reason_for_report = current_mirror_failure_reason; detailed_status = current_mirror_detailed_status
                        if log_window and log_window.winfo_exists(): log_window.update_idletasks() 
                        if pdf_content: download_successful_this_doi = True; successful_mirror_for_this_doi = current_mirror_base_url; failure_reason_for_report = ""; print(f"DESCARGA EXITOSA DOI {doi} usando {current_mirror_base_url}"); break 
                        else:
                            print(f"FALLO con mirror {current_mirror_base_url} DOI {doi}. Razón: {failure_reason_for_report}")
                            if mirror_idx < len(mirrors_to_try_for_this_doi) - 1: print(f"Siguiente mirror en {user_mirror_switch_delay}s..."); time.sleep(user_mirror_switch_delay)
                            else: print(f"Todos los mirrors fallaron para DOI {doi}.")
                    end_time = datetime.now()
                    if download_successful_this_doi and pdf_content: # ... (save PDF and log success)
                        data_for_successful_sheet = original_row_data.copy(); data_for_successful_sheet['Successful_Mirror'] = successful_mirror_for_this_doi; successful_articles_data.append(data_for_successful_sheet)
                        temp_dir = "temp_scihub_pdfs"; 
                        if not os.path.exists(temp_dir): os.makedirs(temp_dir)
                        temp_pdf_path = os.path.join(temp_dir, f"temp_{os.getpid()}_{index}_{pdf_filename_in_zip}")
                        with open(temp_pdf_path, 'wb') as f: f.write(pdf_content)
                        try: total_downloaded_size_bytes += os.path.getsize(temp_pdf_path)
                        except OSError as e: print(f"Advertencia: tamaño temp {temp_pdf_path}: {e}")
                        zf.write(temp_pdf_path, arcname=pdf_filename_in_zip); temp_pdf_paths.append(temp_pdf_path)
                        print(f"AGREGADO AL ZIP: '{effective_title}'."); successful_downloads += 1
                    else: # ... (log failure)
                        print(f"FALLO FINAL: No se pudo descargar '{effective_title}' (DOI: {doi}).")
                        failed_articles_data.append({**original_row_data, 'Failure_Reason': failure_reason_for_report, 'Detailed_Status': detailed_status})
                    log_entry = {**original_row_data, 'Start_Time': start_time.strftime("%Y-%m-%d %H:%M:%S"), 'End_Time': end_time.strftime("%Y-%m-%d %H:%M:%S"), 'Duration_Seconds': (end_time - start_time).total_seconds(), 'Detailed_Status': detailed_status, 'Failure_Reason': failure_reason_for_report, 'Successful_Mirror': successful_mirror_for_this_doi }
                    all_articles_log.append(log_entry)
                    if log_window and log_window.winfo_exists(): log_window.update_idletasks() 
                    print(f"Esperando {user_inter_doi_delay}s antes del siguiente artículo...")
                    time.sleep(user_inter_doi_delay)
        except FileNotFoundError: messagebox.showerror("Error", f"No se pudo crear ZIP (Directorio no encontrado): {zip_path}"); print("Error crítico: FileNotFoundError al crear ZIP."); zip_creation_or_main_loop_error = True 
        except Exception as e: messagebox.showerror("Error", f"Error inesperado en ZIP o descargas: {e}"); print(f"Error crítico: Excepción en ZIP o descargas: {e}."); zip_creation_or_main_loop_error = True

        if not zip_creation_or_main_loop_error:
            # ... (Retry phase as in the previous full version, including update_idletasks calls)
            if failed_articles_data: print(f"\n{'='*20} INICIANDO FASE DE REINTENTO {'='*20}\nSe reintentarán {len(failed_articles_data)} artículos.");
            if log_window and log_window.winfo_exists(): log_window.update_idletasks()
            articles_successfully_retried_ids = [] 
            temp_failed_articles_data_for_iteration = list(failed_articles_data) 
            mirrors_for_retry = list(user_defined_mirrors)
            for retry_idx, failed_article_entry in enumerate(temp_failed_articles_data_for_iteration):
                doi_to_retry = str(failed_article_entry.get('DOI', failed_article_entry.get('doi', ''))).strip()
                effective_title_for_retry = str(failed_article_entry.get('Title', failed_article_entry.get('title', doi_to_retry))).strip() or doi_to_retry
                pdf_filename_in_zip_retry = clean_filename(effective_title_for_retry)[:150] + ".pdf"
                print(f"\nReintentando ({retry_idx + 1}/{len(temp_failed_articles_data_for_iteration)}): {effective_title_for_retry} (DOI: {doi_to_retry})")
                pdf_content_retry = None; retry_successful_this_doi = False; successful_mirror_for_retry = ""; retry_detailed_status = ""; retry_failure_reason = ""
                retry_start_time = datetime.now() 
                for mirror_idx_retry, current_mirror_base_url_retry in enumerate(mirrors_for_retry):
                    print(f"Reintento mirror ({mirror_idx_retry + 1}/{len(mirrors_for_retry)}): {current_mirror_base_url_retry} DOI: {doi_to_retry}")
                    # ... (retry download logic)
                    if log_window and log_window.winfo_exists(): log_window.update_idletasks() 
                    if pdf_content_retry: break 
                    if not pdf_content_retry and mirror_idx_retry < len(mirrors_for_retry) - 1: time.sleep(user_mirror_switch_delay)
                if log_window and log_window.winfo_exists(): log_window.update_idletasks() 
                # ... (rest of retry single DOI processing)
                if retry_idx < len(temp_failed_articles_data_for_iteration) - 1: time.sleep(user_inter_doi_delay)
            if articles_successfully_retried_ids: failed_articles_data = [item for item in failed_articles_data if str(item.get('DOI', item.get('doi', ''))).strip() not in articles_successfully_retried_ids]
            print(f"\n{'='*20} FASE DE REINTENTO COMPLETADA {'='*20}")
            if log_window and log_window.winfo_exists(): log_window.update_idletasks()
            
            # ... (Summary message and Excel generation as in previous full version)
            failed_downloads_summary_list = [{'title': str(item.get('Title',item.get('title','N/A'))).strip(), 'doi': str(item.get('DOI',item.get('doi','N/A'))).strip(), 'reason': str(item.get('Failure_Reason','N/A')).strip()} for item in failed_articles_data]
            total_mb = total_downloaded_size_bytes / (1024 * 1024)
            summary_message = (f"Proceso completado.\n\nDescargas exitosas: {successful_downloads}\nDescargas fallidas: {len(failed_downloads_summary_list)}\n" f"Tamaño total PDFs: {total_mb:.2f} MB")
            if failed_downloads_summary_list: summary_message += "\n\nArtículos no descargados (post-reintentos):"; [summary_message := summary_message + f"\n- Título: {item['title']}, DOI: {item['doi']}, Razón: {item['reason']}" for item in failed_downloads_summary_list] # type: ignore
            print("\n" + "="*50); print(summary_message); print("="*50); 
            if log_window and log_window.winfo_exists(): log_window.update_idletasks() 
            messagebox.showinfo("Resumen Descarga", summary_message)
            generate_excel_report_prompt = messagebox.askyesno("Generar Reporte Excel", "¿Desea generar un reporte Excel detallado?")
            if generate_excel_report_prompt:
                excel_report_path_to_use = excel_report_path_config
                if not excel_report_path_to_use : 
                    excel_report_path_to_use = filedialog.asksaveasfilename(title="Guardar Reporte Excel como...", defaultextension=".xlsx", initialfile="SciHub_Descarga_Reporte.xlsx", filetypes=(("Archivos Excel", "*.xlsx"),("Todos los archivos", "*.*")))
                elif not messagebox.askyesno("Confirmar Ruta de Reporte", f"Se configuró guardar el reporte en:\n{excel_report_path_config}\n\n¿Usar esta ruta?"): 
                        excel_report_path_to_use = filedialog.asksaveasfilename(title="Guardar Reporte Excel en ruta alternativa...", defaultextension=".xlsx", initialfile="SciHub_Descarga_Reporte_Alternativo.xlsx", filetypes=(("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")))
                if excel_report_path_to_use:
                    print(f"Generando reporte Excel en: {excel_report_path_to_use}")
                    if log_window and log_window.winfo_exists(): log_window.update_idletasks()
                    # ... (Excel generation logic)
                else: print("Generación de reporte Excel omitida (ruta no proporcionada).")
            else: print("Generación de reporte Excel omitida por usuario.")

        elif zip_creation_or_main_loop_error: print("Proceso interrumpido por error crítico inicial. No se generará resumen ni Excel.")
    
    finally: 
        restored_to_original_console = False
        if isinstance(sys.stdout, TextRedirector): 
            sys.stdout = original_stdout 
            restored_to_original_console = True
            print("\n--- Limpieza Final de Archivos Temporales (consola original) ---", file=original_stdout)
        else: 
            print("\n--- Limpieza Final de Archivos Temporales ---", file=original_stdout)
        for temp_path in temp_pdf_paths:
            try: os.remove(temp_path); print(f"Eliminado temp: {temp_path}", file=original_stdout)
            except OSError as e: print(f"Error eliminando temp {temp_path}: {e}", file=original_stdout)
        temp_dir_to_check = "temp_scihub_pdfs"
        if os.path.exists(temp_dir_to_check) and not os.listdir(temp_dir_to_check):
            try: os.rmdir(temp_dir_to_check); print(f"Eliminado dir temp: {temp_dir_to_check}", file=original_stdout)
            except OSError as e: print(f"Error eliminando dir temp {temp_dir_to_check}: {e}", file=original_stdout)
        if restored_to_original_console: print("--- Limpieza Finalizada (consola original) ---", file=original_stdout)
        else: print("--- Limpieza Finalizada ---", file=original_stdout)
        if log_window: 
            try:
                if log_window.winfo_exists(): log_window.destroy()
            except tk.TclError: pass 
if __name__ == "__main__":
    download_pdfs_from_file()
    if isinstance(sys.stdout, TextRedirector): 
        sys.stdout = sys.__stdout__ 
    print("\nScript finalizado.")
