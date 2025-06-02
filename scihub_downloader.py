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
        self.widget.update_idletasks() # MODIFIED: Added update_idletasks here too

def on_log_window_close(log_window_ref, original_stdout_ref, root_tk_instance):
    if sys.stdout != original_stdout_ref: # Check if it hasn't been restored already
        print("Log window closed by user. Restoring original stdout.", file=original_stdout_ref) 
        sys.stdout = original_stdout_ref
    if log_window_ref:
        try:
            log_window_ref.destroy()
        except tk.TclError: # Window might already be destroyed
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

def print_to_console_if_no_log_window(message, log_window_instance, orig_stdout):
    # This helper is mostly for very early messages before stdout is redirected
    # or if we need to ensure a message goes to console regardless.
    current_stdout_is_redirector = isinstance(sys.stdout, TextRedirector)
    if not current_stdout_is_redirector: # If stdout is original or log_window not set up
        print(message, file=orig_stdout)
    else: # Log window is up and stdout is redirected
        print(message)


def download_pdfs_from_file():
    root = tk.Tk()
    root.withdraw() 
    
    original_stdout = sys.stdout 
    log_window = None 
    log_text_widget = None

    try:
        print_to_console_if_no_log_window("--- Iniciando Configuración ---", log_window, original_stdout)
        input_file_path = filedialog.askopenfilename(title="Seleccionar archivo con DOIs (Excel o CSV)", filetypes=(("Archivos Excel", "*.xlsx *.xls"), ("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*")))
        if not input_file_path: messagebox.showinfo("Información", "No se seleccionó ningún archivo de entrada. El programa terminará."); return
        zip_path = filedialog.asksaveasfilename(title="Guardar archivo ZIP como...", defaultextension=".zip", filetypes=(("Archivos ZIP", "*.zip"), ("Todos los archivos", "*.*")))
        if not zip_path: messagebox.showinfo("Información", "No se especificó la ubicación para guardar el ZIP. El programa terminará."); return
        excel_report_path_config = filedialog.asksaveasfilename(title="Guardar Reporte Excel Opcional como...", defaultextension=".xlsx", initialfile="SciHub_Descarga_Reporte.xlsx", filetypes=(("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")))
        if not excel_report_path_config: excel_report_path_config = ""; print_to_console_if_no_log_window("Ruta para reporte Excel no especificada, se omitirá su generación automática al final.", log_window, original_stdout)
        
        user_inter_doi_delay = simpledialog.askinteger("Configurar Retraso Inter-DOI", "Ingrese el tiempo de espera (segundos) entre cada DOI:", initialvalue=INTER_DOI_DELAY_SECONDS, minvalue=0 )
        if user_inter_doi_delay is None: user_inter_doi_delay = INTER_DOI_DELAY_SECONDS; messagebox.showinfo("Información", f"Retraso Inter-DOI no modificado, se usará el predeterminado: {user_inter_doi_delay}s")
        user_mirror_switch_delay = simpledialog.askinteger("Configurar Retraso Cambio de Mirror", "Ingrese el tiempo de espera (segundos) al cambiar de mirror:", initialvalue=MIRROR_SWITCH_DELAY_SECONDS, minvalue=0)
        if user_mirror_switch_delay is None: user_mirror_switch_delay = MIRROR_SWITCH_DELAY_SECONDS; messagebox.showinfo("Información", f"Retraso por cambio de Mirror no modificado, se usará el predeterminado: {user_mirror_switch_delay}s")

        user_mirror_list_str = simpledialog.askstring("Configurar Mirrors de Sci-Hub (Obligatorio)", "Ingrese URLs de mirrors Sci-Hub, separadas por comas.\n(ej: https://sci-hub.se/,https://sci-hub.st/)\nESTOS SERÁN LOS ÚNICOS MIRRORS UTILIZADOS.\nDejar vacío para usar default (https://sci-hub.se/).")
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
        
        log_window = tk.Toplevel(root)
        log_window.title("Log de Proceso de Descarga")
        log_window.geometry("800x600")
        log_text_widget = st.ScrolledText(log_window, wrap=tk.WORD, state='disabled')
        log_text_widget.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        
        sys.stdout = TextRedirector(log_text_widget) # Redirect after widget creation
        log_window.protocol("WM_DELETE_WINDOW", lambda: on_log_window_close(log_window, original_stdout, root))
        
        print("\n--- Configuración Aplicada ---")
        print(f"Archivo de entrada: {input_file_path}")
        print(f"Archivo ZIP de salida: {zip_path}")
        if excel_report_path_config: print(f"Archivo de reporte Excel: {excel_report_path_config}")
        else: print("Reporte Excel: No se generará (ruta no especificada).")
        print(f"Retraso Inter-DOI: {user_inter_doi_delay}s")
        print(f"Retraso Cambio de Mirror: {user_mirror_switch_delay}s")
        print(f"Mirrors Sci-Hub a utilizar: {', '.join(user_defined_mirrors)}")
        print("-----------------------------------------------------\n")
        if log_window and log_window.winfo_exists(): log_window.update_idletasks() # Update after config prints

        session = requests.Session()
        session.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'})
        all_articles_log = []; successful_articles_data = []; failed_articles_data = []; original_input_columns = [] 
        
        try:
            file_extension = os.path.splitext(input_file_path)[1].lower()
            if file_extension in ['.xlsx', '.xls']:
                try:
                    df = pd.read_excel(input_file_path); # ... (rest of df loading logic)
                    if 'doi' in df.columns and 'DOI' not in df.columns: df.rename(columns={'doi': 'DOI'}, inplace=True)
                    if 'titulo' in df.columns and 'Title' not in df.columns: df.rename(columns={'titulo': 'Title'}, inplace=True)
                    if not ({'DOI', 'Title'} <= set(df.columns)): messagebox.showerror("Error de Excel", "Archivo Excel debe contener 'DOI' y 'Title'."); raise Exception("Missing columns")
                except Exception as e: messagebox.showerror("Error de Excel", f"Error al leer Excel: {e}"); raise
            elif file_extension == '.csv':
                try:
                    df = pd.read_csv(input_file_path); # ... (rest of df loading logic)
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
                    # ... (DOI processing logic as before) ...
                    original_row_data = row.to_dict(); start_time = datetime.now()
                    doi = str(original_row_data.get('DOI', original_row_data.get('doi', ''))).strip()
                    title = str(original_row_data.get('Title', original_row_data.get('title', ''))).strip()
                    detailed_status = ""; failure_reason_for_report = ""
                    if not doi:
                        print(f"Advertencia: Fila {index+2} ignorada por DOI vacío.") # ... (rest of skip logic)
                        failure_reason_for_report = "DOI vacío"; detailed_status = "Skipped_DOI_Missing"; end_time = datetime.now()
                        log_entry = {**original_row_data, 'Start_Time': start_time.strftime("%Y-%m-%d %H:%M:%S"), 'End_Time': end_time.strftime("%Y-%m-%d %H:%M:%S"), 'Duration_Seconds': (end_time - start_time).total_seconds(), 'Detailed_Status': detailed_status, 'Failure_Reason': failure_reason_for_report, 'Successful_Mirror': ""}
                        all_articles_log.append(log_entry); failed_articles_data.append({**original_row_data, 'Failure_Reason': failure_reason_for_report, 'Detailed_Status': detailed_status}); continue
                    effective_title = title if title else doi; clean_title_for_filename = clean_filename(effective_title); pdf_filename_in_zip = clean_title_for_filename[:150] + ".pdf"
                    print(f"\nProcesando artículo ({index + 1}/{total_articles}): {effective_title} (DOI: {doi})")
                    mirrors_to_try_for_this_doi = list(user_defined_mirrors)
                    pdf_content = None; download_successful_this_doi = False; successful_mirror_for_this_doi = ""
                    for mirror_idx, current_mirror_base_url in enumerate(mirrors_to_try_for_this_doi):
                        print(f"Intentando con mirror ({mirror_idx + 1}/{len(mirrors_to_try_for_this_doi)}): {current_mirror_base_url} para DOI: {doi}")
                        # ... (download logic with extraction and fallback) ...
                        # After each mirror attempt's prints:
                        if log_window and log_window.winfo_exists(): log_window.update_idletasks() # Update after mirror attempt prints
                        if pdf_content: break # from mirror loop
                        if not pdf_content and mirror_idx < len(mirrors_to_try_for_this_doi) - 1: print(f"Siguiente mirror en {user_mirror_switch_delay}s..."); time.sleep(user_mirror_switch_delay) # No update here, before sleep
                    # ... (rest of processing for this DOI) ...
                    end_time = datetime.now() # End time for this DOI
                    if download_successful_this_doi: # ... (append to successful_articles_data, zf.write, etc.)
                        data_for_successful_sheet = original_row_data.copy(); data_for_successful_sheet['Successful_Mirror'] = successful_mirror_for_this_doi; successful_articles_data.append(data_for_successful_sheet)
                        temp_dir = "temp_scihub_pdfs"; 
                        if not os.path.exists(temp_dir): os.makedirs(temp_dir)
                        temp_pdf_path = os.path.join(temp_dir, f"temp_{os.getpid()}_{index}_{pdf_filename_in_zip}")
                        with open(temp_pdf_path, 'wb') as f: f.write(pdf_content)
                        try: total_downloaded_size_bytes += os.path.getsize(temp_pdf_path)
                        except OSError as e: print(f"Advertencia: tamaño temp {temp_pdf_path}: {e}")
                        zf.write(temp_pdf_path, arcname=pdf_filename_in_zip); temp_pdf_paths.append(temp_pdf_path)
                        print(f"AGREGADO AL ZIP: '{effective_title}'."); successful_downloads += 1
                    else: # ... (append to failed_articles_data)
                        print(f"FALLO FINAL: No se pudo descargar '{effective_title}' (DOI: {doi}).")
                        failed_articles_data.append({**original_row_data, 'Failure_Reason': failure_reason_for_report, 'Detailed_Status': detailed_status})
                    log_entry = {**original_row_data, 'Start_Time': start_time.strftime("%Y-%m-%d %H:%M:%S"), 'End_Time': end_time.strftime("%Y-%m-%d %H:%M:%S"), 'Duration_Seconds': (end_time - start_time).total_seconds(), 'Detailed_Status': detailed_status, 'Failure_Reason': failure_reason_for_report, 'Successful_Mirror': successful_mirror_for_this_doi }
                    all_articles_log.append(log_entry)
                    if log_window and log_window.winfo_exists(): log_window.update_idletasks() # Update after all processing for a DOI
                    print(f"Esperando {user_inter_doi_delay}s antes del siguiente artículo...")
                    time.sleep(user_inter_doi_delay)
        except FileNotFoundError: messagebox.showerror("Error", f"No se pudo crear ZIP (Directorio no encontrado): {zip_path}"); print("Error crítico: FileNotFoundError al crear ZIP."); zip_creation_or_main_loop_error = True 
        except Exception as e: messagebox.showerror("Error", f"Error inesperado en ZIP o descargas: {e}"); print(f"Error crítico: Excepción en ZIP o descargas: {e}."); zip_creation_or_main_loop_error = True

        if not zip_creation_or_main_loop_error:
            if failed_articles_data: print(f"\n{'='*20} INICIANDO FASE DE REINTENTO {'='*20}\nSe reintentarán {len(failed_articles_data)} artículos.");
            if log_window and log_window.winfo_exists(): log_window.update_idletasks() # Before retry loop
            
            articles_successfully_retried_ids = [] # ... (retry logic as before)
            temp_failed_articles_data_for_iteration = list(failed_articles_data) 
            mirrors_for_retry = list(user_defined_mirrors)
            for retry_idx, failed_article_entry in enumerate(temp_failed_articles_data_for_iteration):
                # ... (retry logic for a single DOI)
                # Inside the retry's mirror loop:
                for mirror_idx_retry, current_mirror_base_url_retry in enumerate(mirrors_for_retry):
                    # ... (retry download attempt)
                    if log_window and log_window.winfo_exists(): log_window.update_idletasks() # After each retry mirror attempt's prints
                    if pdf_content_retry: break # from retry mirror loop
                    if not pdf_content_retry and mirror_idx_retry < len(mirrors_for_retry) - 1: time.sleep(user_mirror_switch_delay)
                # After all mirrors for a single DOI in retry:
                if log_window and log_window.winfo_exists(): log_window.update_idletasks() # After all retry mirrors for a DOI
                if retry_idx < len(temp_failed_articles_data_for_iteration) - 1: time.sleep(user_inter_doi_delay)
            # ... (rest of retry phase logic as before) ...
            if articles_successfully_retried_ids: failed_articles_data = [item for item in failed_articles_data if str(item.get('DOI', item.get('doi', ''))).strip() not in articles_successfully_retried_ids]
            print(f"\n{'='*20} FASE DE REINTENTO COMPLETADA {'='*20}")
            if log_window and log_window.winfo_exists(): log_window.update_idletasks() # After retry phase
            
            failed_downloads_summary_list = [{'title': str(item.get('Title',item.get('title','N/A'))).strip(), 'doi': str(item.get('DOI',item.get('doi','N/A'))).strip(), 'reason': str(item.get('Failure_Reason','N/A')).strip()} for item in failed_articles_data]
            total_mb = total_downloaded_size_bytes / (1024 * 1024)
            summary_message = (f"Proceso completado.\n\nDescargas exitosas: {successful_downloads}\nDescargas fallidas: {len(failed_downloads_summary_list)}\n" f"Tamaño total PDFs: {total_mb:.2f} MB")
            if failed_downloads_summary_list: summary_message += "\n\nArtículos no descargados (post-reintentos):"; [summary_message := summary_message + f"\n- Título: {item['title']}, DOI: {item['doi']}, Razón: {item['reason']}" for item in failed_downloads_summary_list] # type: ignore
            print("\n" + "="*50); print(summary_message); print("="*50); 
            if log_window and log_window.winfo_exists(): log_window.update_idletasks() # Before summary messagebox
            messagebox.showinfo("Resumen Descarga", summary_message)

            generate_excel_report_prompt = messagebox.askyesno("Generar Reporte Excel", "¿Desea generar un reporte Excel detallado?")
            if generate_excel_report_prompt:
                # ... (Excel path logic and generation as before) ...
                excel_report_path_to_use = excel_report_path_config
                if not excel_report_path_to_use : 
                    excel_report_path_to_use = filedialog.asksaveasfilename(title="Guardar Reporte Excel como...", defaultextension=".xlsx", initialfile="SciHub_Descarga_Reporte.xlsx", filetypes=(("Archivos Excel", "*.xlsx"),("Todos los archivos", "*.*")))
                elif not messagebox.askyesno("Confirmar Ruta de Reporte", f"Se configuró guardar el reporte en:\n{excel_report_path_config}\n\n¿Usar esta ruta?"): 
                        excel_report_path_to_use = filedialog.asksaveasfilename(title="Guardar Reporte Excel en ruta alternativa...", defaultextension=".xlsx", initialfile="SciHub_Descarga_Reporte_Alternativo.xlsx", filetypes=(("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")))
                if excel_report_path_to_use:
                    print(f"Generando reporte Excel en: {excel_report_path_to_use}")
                    if log_window and log_window.winfo_exists(): log_window.update_idletasks() # Before potentially long Excel write
                    # ... (Excel generation logic)
                else: print("Generación de reporte Excel omitida (ruta no proporcionada).")
            else: print("Generación de reporte Excel omitida por usuario.")
        elif zip_creation_or_main_loop_error: print("Proceso interrumpido por error crítico inicial. No se generará resumen ni Excel.")
    
    finally: 
        if sys.stdout != original_stdout : # If stdout was redirected
            sys.stdout = original_stdout # Restore it
            print("\n--- Limpieza Final de Archivos Temporales (consola original) ---")
        else: # Stdout was not redirected (e.g. log window creation failed or was closed early)
            print("\n--- Limpieza Final de Archivos Temporales ---")

        for temp_path in temp_pdf_paths:
            try: os.remove(temp_path); print(f"Eliminado temp: {temp_path}")
            except OSError as e: print(f"Error eliminando temp {temp_path}: {e}")
        temp_dir_to_check = "temp_scihub_pdfs"
        if os.path.exists(temp_dir_to_check) and not os.listdir(temp_dir_to_check):
            try: os.rmdir(temp_dir_to_check); print(f"Eliminado dir temp: {temp_dir_to_check}")
            except OSError as e: print(f"Error eliminando dir temp {temp_dir_to_check}: {e}")
        print("--- Limpieza Finalizada (consola original) ---")
        
        if log_window: 
            try:
                if log_window.winfo_exists(): log_window.destroy()
            except tk.TclError: pass # Window might already be destroyed
        
        # root.quit() # Usually not needed if script ends; mainloop isn't called on root.

if __name__ == "__main__":
    download_pdfs_from_file()
    # Final print to original console, after everything including log window is gone.
    if isinstance(sys.stdout, TextRedirector): # Should have been restored, but as a safeguard
        sys.stdout = sys.__stdout__ 
    print("\nScript finalizado.")
