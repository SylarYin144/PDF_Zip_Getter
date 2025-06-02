import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import pandas as pd
import requests
import zipfile
import os
import re
import time
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
                doi = str(row.get('DOI', '')).strip()
                title = str(row.get('Title', '')).strip()

                if not doi: # Title can be generated from DOI if empty
                    print(f"Advertencia: Fila {index+2} (archivo original) ignorada por DOI vacío.")
                    failed_downloads.append({'title': title or "Título Desconocido", 'doi': "DOI Desconocido", 'reason': "DOI vacío"})
                    continue
                
                effective_title = title if title else doi
                clean_title_for_filename = clean_filename(effective_title)
                pdf_filename_in_zip = clean_title_for_filename[:150] + ".pdf"
                
                print(f"\nProcesando artículo ({index + 1}/{total_articles}): {effective_title}")
                
                full_sci_hub_url = f"{sci_hub_base_url}{doi}"
                pdf_content = None
                download_attempt_reason = ""

                # Primary Download Attempt using extracted link
                print(f"Intentando extraer enlace PDF desde: {full_sci_hub_url}")
                actual_pdf_download_url = extract_pdf_link_from_html(full_sci_hub_url, session)

                if actual_pdf_download_url:
                    print(f"Intentando descargar PDF desde enlace extraído: {actual_pdf_download_url}")
                    try:
                        response = session.get(actual_pdf_download_url, timeout=60) # Increased timeout for direct PDF
                        response.raise_for_status()
                        content_type = response.headers.get('Content-Type', '').lower()
                        if 'application/pdf' in content_type:
                            pdf_content = response.content
                            print(f"ÉXITO (enlace extraído): PDF obtenido para '{effective_title}'.")
                            download_attempt_reason = "Éxito vía enlace extraído"
                        else:
                            print(f"FALLO (enlace extraído): Contenido no es PDF desde {actual_pdf_download_url}. Content-Type: {content_type}")
                            download_attempt_reason = f"Contenido no PDF desde enlace extraído ({content_type})"
                    except requests.exceptions.RequestException as e:
                        print(f"FALLO (enlace extraído): No se pudo descargar desde {actual_pdf_download_url}. Error: {e}")
                        download_attempt_reason = f"Error descargando de enlace extraído: {e}"
                    except Exception as e:
                        print(f"FALLO INESPERADO (enlace extraído): Error descargando desde {actual_pdf_download_url}. Error: {e}")
                        download_attempt_reason = f"Error inesperado (enlace extraído): {e}"

                # Fallback Download Attempt (direct access to Sci-Hub URL)
                if not pdf_content:
                    if actual_pdf_download_url: # Fallback because previous attempt failed
                        print(f"Fallback: Intentando descargar directamente desde {full_sci_hub_url} porque el enlace extraído falló.")
                    else: # Fallback because no link was extracted
                        print(f"Fallback: No se pudo extraer enlace PDF. Intentando descargar directamente desde {full_sci_hub_url}")
                    
                    try:
                        response = session.get(full_sci_hub_url, timeout=30)
                        response.raise_for_status()
                        content_type = response.headers.get('Content-Type', '').lower()
                        if 'application/pdf' in content_type:
                            pdf_content = response.content
                            print(f"ÉXITO (directo): PDF obtenido para '{effective_title}'.")
                            if not download_attempt_reason: # Only set if primary was not attempted or failed silently before this
                                download_attempt_reason = "Éxito vía acceso directo"
                        else:
                            print(f"FALLO (directo): Contenido no es PDF desde {full_sci_hub_url}. Content-Type: {content_type}. Sci-Hub podría haber mostrado una página HTML.")
                            if not download_attempt_reason or "Éxito" not in download_attempt_reason :
                                download_attempt_reason = f"Contenido no PDF desde acceso directo ({content_type})"
                    except requests.exceptions.RequestException as e:
                        print(f"FALLO (directo): No se pudo descargar desde {full_sci_hub_url}. Error: {e}")
                        if not download_attempt_reason or "Éxito" not in download_attempt_reason:
                            download_attempt_reason = f"Error descargando de acceso directo: {e}"
                    except Exception as e:
                        print(f"FALLO INESPERADO (directo): Error descargando desde {full_sci_hub_url}. Error: {e}")
                        if not download_attempt_reason or "Éxito" not in download_attempt_reason:
                            download_attempt_reason = f"Error inesperado (directo): {e}"
                
                # Process pdf_content if obtained
                if pdf_content:
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
                    final_reason = download_attempt_reason if download_attempt_reason else "No se pudo obtener contenido PDF"
                    print(f"FALLO FINAL: No se pudo descargar '{effective_title}' (DOI: {doi}). Razón: {final_reason}")
                    failed_downloads.append({'title': effective_title, 'doi': doi, 'reason': final_reason})
                
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

if __name__ == "__main__":
    # Cambié el nombre de la función principal para reflejar que ahora maneja más que solo Excel
    download_pdfs_from_file()
    print("\nScript finalizado.")
