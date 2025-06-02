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

    # 1. Entrada de URL de Sci-Hub por el Usuario
    default_sci_hub_url = "https://sci-hub.se/"
    user_sci_hub_url = simpledialog.askstring(
        "URL de Sci-Hub",
        "Ingresa la URL base de Sci-Hub:",
        initialvalue=default_sci_hub_url
    )

    if user_sci_hub_url is None: # Usuario cerró el diálogo
        messagebox.showinfo("Información", "No se ingresó URL de Sci-Hub. El programa terminará.")
        return
    elif not user_sci_hub_url.strip(): # Usuario ingresó cadena vacía
        sci_hub_base_url = default_sci_hub_url
        messagebox.showinfo("Información", f"URL de Sci-Hub no especificada, se usará la predeterminada: {default_sci_hub_url}")
    else:
        sci_hub_base_url = user_sci_hub_url.strip()
        if not sci_hub_base_url.endswith('/'):
            sci_hub_base_url += '/'
    
    print(f"Usando URL de Sci-Hub: {sci_hub_base_url}")

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
            # Intentar leer con nombres de columna comunes para Excel
            try:
                df = pd.read_excel(input_file_path, usecols=['DOI', 'Title'])
            except ValueError:
                try:
                    df = pd.read_excel(input_file_path, usecols=['doi', 'titulo'])
                    df.rename(columns={'doi': 'DOI', 'titulo': 'Title'}, inplace=True)
                except ValueError:
                    messagebox.showerror("Error de Excel", "No se pudieron encontrar las columnas 'DOI' y 'Title' (o 'doi' y 'titulo') en el archivo Excel.")
                    return
        elif file_extension == '.csv':
            # Intentar leer con nombres de columna comunes para CSV
            try:
                df = pd.read_csv(input_file_path, usecols=['DOI', 'Title'])
            except ValueError:
                try:
                    df = pd.read_csv(input_file_path, usecols=['doi', 'titulo'])
                    df.rename(columns={'doi': 'DOI', 'titulo': 'Title'}, inplace=True)
                except ValueError:
                    messagebox.showerror("Error de CSV", "No se pudieron encontrar las columnas 'DOI' y 'Title' (o 'doi' y 'titulo') en el archivo CSV.")
                    return
        else:
            messagebox.showerror("Error de Archivo", f"Formato de archivo no soportado: {file_extension}")
            return
            
    except FileNotFoundError:
        messagebox.showerror("Error", f"No se pudo encontrar el archivo: {input_file_path}")
        return
    except Exception as e: # Captura otros errores de pandas o generales al leer
        messagebox.showerror("Error de Lectura", f"Ocurrió un error al leer el archivo: {e}")
        return

    successful_downloads = 0
    failed_downloads = []
    temp_pdf_paths = [] # Para almacenar rutas de PDFs temporales
    total_downloaded_size_bytes = 0

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
                
                print(f"\nProcesando artículo ({index + 1}/{total_articles}): {effective_title}")
                
                full_sci_hub_url = f"{sci_hub_base_url}{doi}"
                pdf_content = None
                # download_attempt_reason = "" # Replaced by detailed_status and failure_reason_for_report

                # Primary Download Attempt using extracted link
                print(f"Intentando extraer enlace PDF desde: {full_sci_hub_url}")
                actual_pdf_download_url = extract_pdf_link_from_html(full_sci_hub_url, session)

                if actual_pdf_download_url:
                    print(f"Intentando descargar PDF desde enlace extraído: {actual_pdf_download_url}")
                    try:
                        response = session.get(actual_pdf_download_url, timeout=60)
                        response.raise_for_status()
                        content_type = response.headers.get('Content-Type', '').lower()
                        if 'application/pdf' in content_type:
                            pdf_content = response.content
                            detailed_status = "Success_iframe_or_embed_extraction"
                            print(f"ÉXITO (enlace extraído): PDF obtenido para '{effective_title}'.")
                        else:
                            failure_reason_for_report = f"Content-Type was not application/pdf, but {content_type}"
                            detailed_status = "Failure_iframe_or_embed_extraction_not_pdf"
                            print(f"FALLO (enlace extraído): Contenido no es PDF desde {actual_pdf_download_url}. Content-Type: {content_type}")
                    except requests.exceptions.HTTPError as e:
                        failure_reason_for_report = f"HTTPError: {str(e)} (Status: {e.response.status_code})"
                        detailed_status = f"Failure_iframe_or_embed_extraction_HTTP{e.response.status_code}"
                        print(f"FALLO (enlace extraído): HTTP Error {e.response.status_code} para {actual_pdf_download_url}. Error: {e}")
                    except requests.exceptions.RequestException as e:
                        failure_reason_for_report = str(e)
                        detailed_status = "Failure_iframe_or_embed_extraction_RequestException"
                        print(f"FALLO (enlace extraído): No se pudo descargar desde {actual_pdf_download_url}. Error: {e}")
                    except Exception as e:
                        failure_reason_for_report = f"Unexpected error: {str(e)}"
                        detailed_status = "Failure_iframe_or_embed_extraction_Unexpected"
                        print(f"FALLO INESPERADO (enlace extraído): Error descargando desde {actual_pdf_download_url}. Error: {e}")
                else:
                    failure_reason_for_report = "No PDF link found in HTML (iframe/embed)"
                    detailed_status = "Failure_No_PDF_Link_Found_In_HTML"
                    print(f"Fallo al extraer enlace: No se encontró iframe/embed con enlace PDF en {full_sci_hub_url}")


                # Fallback Download Attempt (direct access to Sci-Hub URL)
                if not pdf_content:
                    if actual_pdf_download_url: 
                        print(f"Fallback: Intentando descargar directamente desde {full_sci_hub_url} porque el enlace extraído ({actual_pdf_download_url}) falló.")
                    else: 
                        print(f"Fallback: No se pudo extraer enlace PDF. Intentando descargar directamente desde {full_sci_hub_url}")
                    
                    try:
                        response = session.get(full_sci_hub_url, timeout=30)
                        response.raise_for_status()
                        content_type = response.headers.get('Content-Type', '').lower()
                        if 'application/pdf' in content_type:
                            pdf_content = response.content
                            detailed_status = "Success_direct_DOI_access_fallback"
                            failure_reason_for_report = "" # Clear previous failure reason if fallback succeeded
                            print(f"ÉXITO (directo-fallback): PDF obtenido para '{effective_title}'.")
                        else:
                            # Keep the failure_reason_for_report from the extraction attempt if this also fails due to content type
                            if not detailed_status.startswith("Failure_No_PDF_Link_Found"): # if link was found but failed, this is a new reason
                                failure_reason_for_report = f"Content-Type was not application/pdf from direct DOI access, but {content_type}"
                            detailed_status = "Failure_direct_DOI_access_not_pdf"
                            print(f"FALLO (directo-fallback): Contenido no es PDF desde {full_sci_hub_url}. Content-Type: {content_type}.")
                    except requests.exceptions.HTTPError as e:
                        failure_reason_for_report = f"HTTPError direct DOI: {str(e)} (Status: {e.response.status_code})"
                        detailed_status = f"Failure_direct_DOI_access_HTTP{e.response.status_code}"
                        print(f"FALLO (directo-fallback): HTTP Error {e.response.status_code} para {full_sci_hub_url}. Error: {e}")
                    except requests.exceptions.RequestException as e:
                        failure_reason_for_report = f"RequestException direct DOI: {str(e)}"
                        detailed_status = "Failure_direct_DOI_access_RequestException"
                        print(f"FALLO (directo-fallback): No se pudo descargar desde {full_sci_hub_url}. Error: {e}")
                    except Exception as e:
                        failure_reason_for_report = f"Unexpected error direct DOI: {str(e)}"
                        detailed_status = "Failure_direct_DOI_access_Unexpected"
                        print(f"FALLO INESPERADO (directo-fallback): Error descargando desde {full_sci_hub_url}. Error: {e}")
                
                end_time = datetime.now()

                # Process pdf_content if obtained
                if pdf_content:
                    successful_articles_data.append(original_row_data)
                    # detailed_status is already set from successful download path
                    # failure_reason_for_report should be empty or reflect the successful method if needed, but generally empty for success

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
                    'Failure_Reason': failure_reason_for_report
                    # SciHub_Link will be added later
                }
                all_articles_log.append(log_entry)
                
                # Introduce a delay before processing the next article
                print(f"Esperando 5 segundos antes del siguiente artículo...")
                time.sleep(5)

    except FileNotFoundError:
        messagebox.showerror("Error", f"No se pudo crear el archivo ZIP en la ruta especificada (Directorio no encontrado): {zip_path}")
        return
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error inesperado durante la creación del ZIP: {e}")
        return
    finally:
        # Limpiar archivos PDF temporales y el directorio temporal si está vacío
        for temp_path in temp_pdf_paths:
            try:
                os.remove(temp_path)
                print(f"Archivo temporal eliminado: {temp_path}")
            except OSError as e:
                print(f"Error al eliminar archivo temporal {temp_path}: {e}")
        
        temp_dir_to_check = "temp_scihub_pdfs" # El mismo directorio usado antes
        if os.path.exists(temp_dir_to_check) and not os.listdir(temp_dir_to_check):
            try:
                os.rmdir(temp_dir_to_check)
                print(f"Directorio temporal eliminado: {temp_dir_to_check}")
            except OSError as e:
                print(f"Error al eliminar el directorio temporal {temp_dir_to_check}: {e}")


    # Informar al usuario
    total_mb = total_downloaded_size_bytes / (1024 * 1024)
    summary_message = (
        f"Proceso completado.\n\n"
        f"Descargas exitosas: {successful_downloads}\n"
        f"Descargas fallidas: {len(failed_downloads)}\n"
        f"Tamaño total de los PDFs descargados: {total_mb:.2f} MB" # 3. Reporte del tamaño
    )
    
    if failed_downloads:
        summary_message += "\n\nArtículos que no se pudieron descargar:"
        for item in failed_downloads:
            summary_message += f"\n- Título: {item['title']}, DOI: {item['doi']}, Razón: {item['reason']}"
    
    print("\n" + "="*50)
    print(summary_message)
    print("="*50)
    messagebox.showinfo("Resumen de Descarga", summary_message)

    # Prompt for Excel report generation
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
                # Prepare 'Obtenidos' sheet
                df_obtenidos = pd.DataFrame(successful_articles_data)
                if not df_obtenidos.empty:
                    df_obtenidos['SciHub_Link'] = df_obtenidos.apply(
                        lambda row: f"{sci_hub_base_url}{row.get('DOI', row.get('doi', ''))}" if pd.notna(row.get('DOI', row.get('doi', ''))) else '', 
                        axis=1
                    )
                else: # Ensure SciHub_Link column exists even if empty
                    # Need to know original columns if we want to preserve them all before adding SciHub_Link
                    # For simplicity, if completely empty, just define it. If it has columns, add to them.
                    if not 'SciHub_Link' in df_obtenidos.columns:
                         df_obtenidos['SciHub_Link'] = pd.Series(dtype='object')


                # Prepare 'Fallidos' sheet
                df_fallidos = pd.DataFrame(failed_articles_data)
                if not df_fallidos.empty:
                    df_fallidos['SciHub_Link'] = df_fallidos.apply(
                        lambda row: f"{sci_hub_base_url}{row.get('DOI', row.get('doi', ''))}" if pd.notna(row.get('DOI', row.get('doi', ''))) else '', 
                        axis=1
                    )
                else:
                    if not 'SciHub_Link' in df_fallidos.columns:
                         df_fallidos['SciHub_Link'] = pd.Series(dtype='object')
                    # It should already have Failure_Reason and Detailed_Status from data collection step
                
                # Prepare 'Tiempos' sheet
                df_tiempos = pd.DataFrame(all_articles_log)
                if not df_tiempos.empty:
                    df_tiempos['SciHub_Link'] = df_tiempos.apply(
                        lambda row: f"{sci_hub_base_url}{row.get('DOI', row.get('doi', ''))}" if pd.notna(row.get('DOI', row.get('doi', ''))) else '', 
                        axis=1
                    )
                else:
                     if not 'SciHub_Link' in df_tiempos.columns:
                         df_tiempos['SciHub_Link'] = pd.Series(dtype='object')
                # It should have all original columns + timing + status columns

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

if __name__ == "__main__":
    # Cambié el nombre de la función principal para reflejar que ahora maneja más que solo Excel
    download_pdfs_from_file()
    print("\nScript finalizado.")
