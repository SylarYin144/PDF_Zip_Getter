import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import pandas as pd
import requests
import zipfile
import os
import re

def clean_filename(title):
    """
    Limpia el título para que sea un nombre de archivo válido.
    Reemplaza caracteres no permitidos con guiones bajos.
    """
    return re.sub(r'[\\/*?:"<>|]', '_', title)

def download_pdfs_from_file():
    """
    Función principal para descargar PDFs desde un archivo Excel o CSV de DOIs.
    """
    # Inicializar Tkinter (necesario para los diálogos)
    root = tk.Tk()
    root.withdraw()  # Ocultar la ventana principal de Tkinter

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
    total_downloaded_size_bytes = 0 # 3. Para calcular el tamaño total

    # Crear archivo ZIP
    try:
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            total_articles = len(df)
            for index, row in df.iterrows():
                doi = str(row.get('DOI', '')).strip() # Usar .get para evitar KeyError si la columna no existe (aunque ya se verificó)
                title = str(row.get('Title', '')).strip()
                
                if not doi or not title:
                    print(f"Advertencia: Fila {index+2} (archivo original) ignorada por DOI o Título vacío.")
                    failed_downloads.append({'title': title or "Título Desconocido", 'doi': doi or "DOI Desconocido", 'reason': "DOI o Título vacío"})
                    continue

                print(f"\nProcesando artículo ({index + 1}/{total_articles}): {title if title else 'Sin Título'}")
                
                # Si el título está vacío después de todo, usar el DOI para el nombre del archivo
                # Esto no debería ocurrir si la comprobación anterior de 'not title' funciona.
                # Pero como defensa adicional:
                effective_title = title if title else doi 
                clean_title_for_filename = clean_filename(effective_title)
                
                pdf_filename_in_zip = clean_title_for_filename[:150] + ".pdf" 
                full_sci_hub_url = f"{sci_hub_base_url}{doi}"

                print(f"Intentando descargar desde: {full_sci_hub_url}")

                try:
                    # Intentar descargar el PDF
                    response = requests.get(full_sci_hub_url, timeout=30, headers={'User-Agent': 'Mozilla/5.0'})
                    response.raise_for_status() 

                    content_type = response.headers.get('Content-Type', '').lower()
                    
                    if 'application/pdf' in content_type:
                        # Guardar temporalmente el PDF
                        # Crear un directorio temporal si no existe para los PDFs
                        temp_dir = "temp_scihub_pdfs"
                        if not os.path.exists(temp_dir):
                            os.makedirs(temp_dir)
                        temp_pdf_path = os.path.join(temp_dir, f"temp_{os.getpid()}_{index}_{pdf_filename_in_zip}")
                        
                        with open(temp_pdf_path, 'wb') as f:
                            f.write(response.content)
                        
                        # 3. Obtener tamaño del archivo y sumarlo
                        try:
                            pdf_size_bytes = os.path.getsize(temp_pdf_path)
                            total_downloaded_size_bytes += pdf_size_bytes
                        except OSError as e:
                            print(f"Advertencia: No se pudo obtener el tamaño del archivo temporal {temp_pdf_path}: {e}")


                        # Agregar al ZIP
                        zf.write(temp_pdf_path, arcname=pdf_filename_in_zip)
                        temp_pdf_paths.append(temp_pdf_path) 
                        
                        print(f"ÉXITO: '{title if title else doi}' descargado y agregado al ZIP como '{pdf_filename_in_zip}'.")
                        successful_downloads += 1
                    else:
                        print(f"FALLO: El contenido para '{title if title else doi}' (DOI: {doi}) no es un PDF. Content-Type: {content_type}. Sci-Hub podría haber mostrado una página HTML.")
                        failed_downloads.append({'title': title if title else "Título Desconocido", 'doi': doi, 'reason': 'No es un archivo PDF'})

                except requests.exceptions.RequestException as e:
                    print(f"FALLO: No se pudo descargar '{title if title else doi}' (DOI: {doi}). Error: {e}")
                    failed_downloads.append({'title': title if title else "Título Desconocido", 'doi': doi, 'reason': str(e)})
                except Exception as e:
                    print(f"FALLO INESPERADO: Ocurrió un error procesando '{title if title else doi}' (DOI: {doi}). Error: {e}")
                    failed_downloads.append({'title': title if title else "Título Desconocido", 'doi': doi, 'reason': f"Error inesperado: {str(e)}"})

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
