import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import threading
import queue
import platform
import subprocess
from downloader_class import Downloader, DEFAULT_SCI_HUB_MIRRORS
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from collections import Counter

class ConfigFrame(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.input_file_path = ""
        self.zip_file_path = ""

        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # --- INPUT/OUTPUT SECTION ---
        io_frame = ctk.CTkFrame(self)
        io_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        io_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(io_frame, text="1. Archivos de Entrada y Salida", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="w")

        self.select_file_button = ctk.CTkButton(io_frame, text="Seleccionar Archivo (.xlsx, .csv)", command=self.select_input_file)
        self.select_file_button.grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.file_info_label = ctk.CTkLabel(io_frame, text="No se ha cargado ningún archivo.")
        self.file_info_label.grid(row=1, column=1, padx=10, pady=10, sticky="w")

        self.select_zip_button = ctk.CTkButton(io_frame, text="Definir Ubicación del ZIP", command=self.select_zip_location)
        self.select_zip_button.grid(row=2, column=0, padx=10, pady=10, sticky="w")
        self.zip_path_label = ctk.CTkLabel(io_frame, text="No se ha seleccionado la ubicación.")
        self.zip_path_label.grid(row=3, column=0, columnspan=2, padx=10, pady=(0, 10), sticky="w")

        # --- SOURCES SECTION ---
        sources_frame = ctk.CTkFrame(self)
        sources_frame.grid(row=1, column=0, padx=20, pady=20, sticky="nsew")

        ctk.CTkLabel(sources_frame, text="2. Fuentes de Descarga", font=ctk.CTkFont(weight="bold")).pack(padx=10, pady=(10, 5), anchor="w")
        self.check_scihub = ctk.CTkCheckBox(sources_frame, text="Usar Sci-Hub"); self.check_scihub.pack(padx=10, pady=5, anchor="w"); self.check_scihub.select(); self.check_scihub.configure(state="disabled")
        self.check_gscholar = ctk.CTkCheckBox(sources_frame, text="Usar Google Scholar"); self.check_gscholar.pack(padx=10, pady=5, anchor="w"); self.check_gscholar.select()
        self.check_pmc = ctk.CTkCheckBox(sources_frame, text="Usar PubMed Central (PMC)"); self.check_pmc.pack(padx=10, pady=(5, 10), anchor="w"); self.check_pmc.select()

        # --- ADVANCED CONFIG SECTION ---
        adv_frame = ctk.CTkFrame(self)
        adv_frame.grid(row=0, column=1, rowspan=2, padx=20, pady=20, sticky="nsew")
        adv_frame.grid_columnconfigure(0, weight=1)
        adv_frame.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(adv_frame, text="3. Configuración Avanzada", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="w")
        ctk.CTkLabel(adv_frame, text="Mirrors de Sci-Hub (separados por coma):").grid(row=1, column=0, columnspan=2, padx=10, pady=(5,0), sticky="w")
        self.mirrors_textbox = ctk.CTkTextbox(adv_frame)
        self.mirrors_textbox.grid(row=2, column=0, columnspan=2, padx=10, pady=5, sticky="nsew")
        self.mirrors_textbox.insert("1.0", ",\n".join(DEFAULT_SCI_HUB_MIRRORS))
        ctk.CTkLabel(adv_frame, text="Tiempo de espera entre DOIs (s):").grid(row=3, column=0, padx=10, pady=(10, 0), sticky="w")
        self.delay_entry = ctk.CTkEntry(adv_frame, width=80); self.delay_entry.grid(row=3, column=1, padx=10, pady=(10,0), sticky="e"); self.delay_entry.insert(0, "2")

        ctk.CTkLabel(adv_frame, text="Timeout Carga de Página (s):").grid(row=4, column=0, padx=10, pady=(10, 0), sticky="w")
        self.page_load_timeout_entry = ctk.CTkEntry(adv_frame, width=80); self.page_load_timeout_entry.grid(row=4, column=1, padx=10, pady=(10,0), sticky="e"); self.page_load_timeout_entry.insert(0, "60")

        ctk.CTkLabel(adv_frame, text="Timeout Búsqueda Elemento (s):").grid(row=5, column=0, padx=10, pady=(10, 0), sticky="w")
        self.element_wait_timeout_entry = ctk.CTkEntry(adv_frame, width=80); self.element_wait_timeout_entry.grid(row=5, column=1, padx=10, pady=(10,0), sticky="e"); self.element_wait_timeout_entry.insert(0, "20")

        # --- REPORT SECTION ---
        report_frame = ctk.CTkFrame(self)
        report_frame.grid(row=2, column=0, columnspan=2, padx=20, pady=20, sticky="nsew")
        self.report_checkbox = ctk.CTkCheckBox(report_frame, text="Generar reporte detallado en Excel"); self.report_checkbox.pack(padx=10, pady=10, anchor="w"); self.report_checkbox.select()

        # --- ACTION BUTTON ---
        self.start_button = ctk.CTkButton(self, text="🚀 Iniciar Descarga", command=self.start_download, font=ctk.CTkFont(size=16)); self.start_button.grid(row=3, column=0, columnspan=2, padx=20, pady=20, sticky="ew"); self.start_button.configure(state="disabled")

    def select_input_file(self):
        path = filedialog.askopenfilename(title="Seleccionar archivo con DOIs", filetypes=(("Archivos Excel", "*.xlsx *.xls"), ("Archivos CSV", "*.csv")))
        if not path: return
        self.input_file_path = path
        try:
            df = pd.read_csv(path, dtype=str) if path.endswith('.csv') else pd.read_excel(path, dtype=str)
            df.fillna('', inplace=True)
            if 'DOI' not in df.columns and 'doi' not in df.columns:
                 raise ValueError("El archivo no contiene la columna 'DOI'.")
            if 'doi' in df.columns and 'DOI' not in df.columns:
                df.rename(columns={'doi': 'DOI'}, inplace=True)
            self.file_info_label.configure(text=f"{os.path.basename(path)} ({len(df)} artículos detectados)")
            self.controller.input_df = df
        except Exception as e:
            self.file_info_label.configure(text=f"Error al leer archivo: {e}", text_color="red"); self.input_file_path = ""
        self._check_start_conditions()

    def select_zip_location(self):
        path = filedialog.asksaveasfilename(title="Guardar archivo ZIP como...", defaultextension=".zip", filetypes=(("Archivos ZIP", "*.zip"),))
        if not path: return
        self.zip_file_path = path
        self.zip_path_label.configure(text=f"Se guardará en: {path}")
        self._check_start_conditions()

    def _check_start_conditions(self):
        if self.input_file_path and self.zip_file_path and self.controller.input_df is not None: self.start_button.configure(state="normal")
        else: self.start_button.configure(state="disabled")

    def start_download(self):
        mirrors = [m.strip() for m in self.mirrors_textbox.get("1.0", "end-1c").split(',') if m.strip()]
        if not mirrors:
            self.controller.show_error("Lista de mirrors vacía", "Por favor, especifique al menos un mirror de Sci-Hub.")
            return

        try:
            delay = int(self.delay_entry.get())
            page_load_timeout = int(self.page_load_timeout_entry.get())
            element_wait_timeout = int(self.element_wait_timeout_entry.get())
        except ValueError:
            self.controller.show_error("Valor inválido", "Los tiempos de espera deben ser números enteros.")
            return

        config = {
            'input_df': self.controller.input_df,
            'zip_path': self.zip_file_path,
            'excel_report_path': self.zip_file_path.replace('.zip', '_reporte.xlsx') if self.report_checkbox.get() else None,
            'sci_hub_mirrors': mirrors,
            'use_google_scholar': bool(self.check_gscholar.get()),
            'use_pmc': bool(self.check_pmc.get()),
            'inter_doi_delay': delay,
            'page_load_timeout': page_load_timeout,
            'element_wait_timeout': element_wait_timeout,
        }
        self.controller.start_download_process(config)


class ProgressFrame(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        top_frame = ctk.CTkFrame(self); top_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        top_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)
        self.progress_bar = ctk.CTkProgressBar(top_frame); self.progress_bar.grid(row=0, column=0, columnspan=4, padx=10, pady=10, sticky="ew"); self.progress_bar.set(0)
        self.progress_label = ctk.CTkLabel(top_frame, text="Procesando 0 de 0... (0%)"); self.progress_label.grid(row=1, column=0, columnspan=4, padx=10, pady=(0, 10))

        kpi_frame = ctk.CTkFrame(self); kpi_frame.grid(row=1, column=0, padx=10, pady=10, sticky="ew")
        kpi_frame.grid_columnconfigure((0,1,2), weight=1)
        self.kpi_obtained = ctk.CTkLabel(kpi_frame, text="🟩 OBTENIDOS\n0", font=ctk.CTkFont(size=16)); self.kpi_obtained.grid(row=0, column=0)
        self.kpi_failed = ctk.CTkLabel(kpi_frame, text="🟥 FALLIDOS\n0", font=ctk.CTkFont(size=16)); self.kpi_failed.grid(row=0, column=1)
        self.kpi_pending = ctk.CTkLabel(kpi_frame, text="🟦 PENDIENTES\n0", font=ctk.CTkFont(size=16)); self.kpi_pending.grid(row=0, column=2)

        self.article_list_frame = ctk.CTkScrollableFrame(self, label_text="Lista de Artículos"); self.article_list_frame.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")
        self.article_widgets = {}

        bottom_frame = ctk.CTkFrame(self); bottom_frame.grid(row=3, column=0, padx=10, pady=10, sticky="ew")
        bottom_frame.grid_columnconfigure(0, weight=3); bottom_frame.grid_columnconfigure(1, weight=1)
        self.log_textbox = ctk.CTkTextbox(bottom_frame, height=150); self.log_textbox.grid(row=0, column=0, padx=(10,5), pady=10, sticky="ew"); self.log_textbox.configure(state="disabled")

        controls_frame = ctk.CTkFrame(bottom_frame); controls_frame.grid(row=0, column=1, padx=(5,10), pady=10, sticky="ns")
        self.pause_button = ctk.CTkButton(controls_frame, text="Pausar", command=self.controller.toggle_pause); self.pause_button.pack(padx=10, pady=10, expand=True, fill="x")
        self.cancel_button = ctk.CTkButton(controls_frame, text="Cancelar", fg_color="red", hover_color="darkred", command=self.controller.cancel_download); self.cancel_button.pack(padx=10, pady=10, expand=True, fill="x")

    def reset_ui(self, total_articles):
        self.progress_bar.set(0)
        self.progress_label.configure(text=f"Procesando 0 de {total_articles}... (0%)")
        self.kpi_obtained.configure(text="🟩 OBTENIDOS\n0")
        self.kpi_failed.configure(text="🟥 FALLIDOS\n0")
        self.kpi_pending.configure(text=f"🟦 PENDIENTES\n{total_articles}")
        self.log_textbox.configure(state="normal"); self.log_textbox.delete("1.0", "end"); self.log_textbox.configure(state="disabled")
        for widget in self.article_list_frame.winfo_children(): widget.destroy()
        self.article_widgets = {}
        self.pause_button.configure(text="Pausar", state="normal")
        self.cancel_button.configure(state="normal")

    def populate_initial_articles(self, df):
        self.article_list_frame.grid_columnconfigure(1, weight=1)
        for index, row in df.iterrows():
            doi = row.get("DOI", "N/A")
            title = row.get("Title", doi)
            status_label = ctk.CTkLabel(self.article_list_frame, text="⏳ En cola", width=150); status_label.grid(row=index, column=0, padx=5, pady=2, sticky="w")
            title_label = ctk.CTkLabel(self.article_list_frame, text=title, anchor="w"); title_label.grid(row=index, column=1, padx=5, pady=2, sticky="ew")
            self.article_widgets[doi] = {'status': status_label, 'title': title_label}

    def update_article_status(self, doi, success, source, reason):
        if doi in self.article_widgets:
            status_text = f"✅ Obtenido ({source})" if success else f"❌ Fallido"
            color = "#34A853" if success else "#EA4335"
            self.article_widgets[doi]['status'].configure(text=status_text, text_color=color)

    def add_log_message(self, message):
        self.log_textbox.configure(state="normal"); self.log_textbox.insert("end", message + "\n"); self.log_textbox.see("end"); self.log_textbox.configure(state="disabled")

    def finalize_ui(self):
        self.pause_button.configure(state="disabled")
        self.cancel_button.configure(state="disabled")


class ResultsFrame(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.results_data = {}
        self.grid_columnconfigure(0, weight=1); self.grid_rowconfigure(2, weight=1)

        self.summary_label = ctk.CTkLabel(self, text="", font=ctk.CTkFont(size=20, weight="bold")); self.summary_label.grid(row=0, column=0, padx=20, pady=20)

        top_content_frame = ctk.CTkFrame(self, fg_color="transparent"); top_content_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        top_content_frame.grid_columnconfigure(0, weight=1); top_content_frame.grid_columnconfigure(1, weight=1)

        kpi_frame = ctk.CTkFrame(top_content_frame); kpi_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        kpi_frame.grid_columnconfigure((0, 1), weight=1)
        self.kpi_obtained = ctk.CTkLabel(kpi_frame, text="🟩 OBTENIDOS\n0", font=ctk.CTkFont(size=16)); self.kpi_obtained.grid(row=0, column=0, pady=10)
        self.kpi_failed = ctk.CTkLabel(kpi_frame, text="🟥 FALLIDOS\n0", font=ctk.CTkFont(size=16)); self.kpi_failed.grid(row=0, column=1, pady=10)

        self.chart_frame = ctk.CTkFrame(top_content_frame); self.chart_frame.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        self.chart_canvas = None

        self.tab_view = ctk.CTkTabview(self); self.tab_view.grid(row=2, column=0, padx=20, pady=10, sticky="nsew")
        self.tab_view.add("Obtenidos"); self.tab_view.add("Fallidos")
        self.obtained_frame = ctk.CTkScrollableFrame(self.tab_view.tab("Obtenidos")); self.obtained_frame.pack(fill="both", expand=True)
        self.failed_frame = ctk.CTkScrollableFrame(self.tab_view.tab("Fallidos")); self.failed_frame.pack(fill="both", expand=True)

        self.copy_failed_button = ctk.CTkButton(self.tab_view.tab("Fallidos"), text="Copiar DOIs fallidos", command=self.copy_failed_dois); self.copy_failed_button.pack(pady=10)

        actions_frame = ctk.CTkFrame(self); actions_frame.grid(row=3, column=0, padx=20, pady=20, sticky="ew")
        actions_frame.grid_columnconfigure((0,1,2), weight=1)
        self.open_folder_button = ctk.CTkButton(actions_frame, text="📂 Abrir Carpeta de Descarga", command=self.open_download_folder); self.open_folder_button.grid(row=0, column=0, padx=5, pady=10)
        self.open_report_button = ctk.CTkButton(actions_frame, text="🧾 Ver Reporte en Excel", command=self.open_excel_report); self.open_report_button.grid(row=0, column=1, padx=5, pady=10)
        self.new_task_button = ctk.CTkButton(actions_frame, text="🔄 Iniciar una Nueva Tarea", command=lambda: controller.show_frame("ConfigFrame")); self.new_task_button.grid(row=0, column=2, padx=5, pady=10)

    def update_results(self, results):
        self.results_data = results
        total = results['total_articles']; success = results['successful_count']
        msg = "¡Proceso Completado!" if not results['was_cancelled'] else "Proceso Cancelado"
        self.summary_label.configure(text=f"{msg} Se obtuvieron {success} de {total} artículos.")
        self.kpi_obtained.configure(text=f"🟩 OBTENIDOS\n{success}")
        self.kpi_failed.configure(text=f"🟥 FALLIDOS\n{results['failed_count']}")

        for widget in self.obtained_frame.winfo_children(): widget.destroy()
        for item in results['successful_articles']:
            ctk.CTkLabel(self.obtained_frame, text=f"✅ {item['data'].get('Title', item['data']['DOI'])} (Fuente: {item['source']})").pack(anchor="w", padx=5)

        for widget in self.failed_frame.winfo_children(): widget.destroy()
        for item in results['failed_articles']:
            ctk.CTkLabel(self.failed_frame, text=f"❌ {item['data'].get('Title', item['data']['DOI'])} (Razón: {item['reason']})").pack(anchor="w", padx=5)
        self.copy_failed_button.pack_forget() if not results['failed_articles'] else self.copy_failed_button.pack(pady=10)

        self.open_report_button.configure(state="normal" if results['excel_report_path'] else "disabled")
        self._create_donut_chart(results)

    def _create_donut_chart(self, results):
        if self.chart_canvas: self.chart_canvas.get_tk_widget().destroy()

        sources = [item['source'] for item in results['successful_articles']]
        source_counts = Counter(sources)

        if not source_counts:
            no_data_label = ctk.CTkLabel(self.chart_frame, text="No hay datos para el gráfico.")
            no_data_label.pack(expand=True)
            return

        labels = list(source_counts.keys())
        sizes = list(source_counts.values())

        is_dark = ctk.get_appearance_mode() == "Dark"
        bg_color = self.chart_frame.cget("fg_color")[1]
        text_color = "#FFFFFF" if is_dark else "#000000"

        fig = Figure(figsize=(4, 2.5), dpi=100, facecolor=bg_color)
        ax = fig.add_subplot(111)

        wedges, texts, autotexts = ax.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90,
                                          wedgeprops=dict(width=0.4), textprops={'color': text_color})

        ax.set_title("Fuentes de Éxito", color=text_color)
        fig.tight_layout()

        self.chart_canvas = FigureCanvasTkAgg(fig, master=self.chart_frame)
        self.chart_canvas.draw()
        self.chart_canvas.get_tk_widget().pack(side=ctk.TOP, fill=ctk.BOTH, expand=True)

    def copy_failed_dois(self):
        failed_dois = [item['data']['DOI'] for item in self.results_data.get('failed_articles', []) if item['data']['DOI']]
        if failed_dois:
            self.clipboard_clear(); self.clipboard_append("\n".join(failed_dois))
            self.controller.show_info("Copiado", f"{len(failed_dois)} DOIs fallidos copiados al portapapeles.")

    def open_path(self, path):
        if not path or not os.path.exists(path):
            self.controller.show_error("Error", f"La ruta no existe o no fue especificada:\n{path}"); return
        try:
            if platform.system() == "Windows": os.startfile(os.path.dirname(path) if os.path.isfile(path) else path)
            elif platform.system() == "Darwin": subprocess.run(["open", path if os.path.isdir(path) else os.path.dirname(path)])
            else: subprocess.run(["xdg-open", path if os.path.isdir(path) else os.path.dirname(path)])
        except Exception as e: self.controller.show_error("Error al abrir", f"No se pudo abrir la ruta:\n{e}")

    def open_download_folder(self): self.open_path(self.results_data.get('zip_path'))
    def open_excel_report(self): self.open_path(self.results_data.get('excel_report_path'))


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Sci-Hub Downloader Pro"); self.geometry("1200x800")
        self.input_df = None; self.downloader_thread = None
        self.progress_queue = queue.Queue()
        self.cancel_event = threading.Event(); self.pause_event = threading.Event()

        ctk.set_appearance_mode("System"); ctk.set_default_color_theme("blue")
        container = ctk.CTkFrame(self); container.pack(side="top", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1); container.grid_columnconfigure(0, weight=1)

        self.frames = {}
        for F in (ConfigFrame, ProgressFrame, ResultsFrame):
            frame = F(parent=container, controller=self); self.frames[F.__name__] = frame; frame.grid(row=0, column=0, sticky="nsew")
        self.show_frame("ConfigFrame")

    def show_frame(self, frame_name): self.frames[frame_name].tkraise()

    def start_download_process(self, config):
        self.cancel_event.clear(); self.pause_event.clear()
        progress_frame = self.frames['ProgressFrame']; progress_frame.reset_ui(len(config['input_df'])); progress_frame.populate_initial_articles(config['input_df'])
        self.show_frame("ProgressFrame")
        downloader = Downloader(config, self.progress_queue, self.cancel_event, self.pause_event)
        self.downloader_thread = threading.Thread(target=downloader.run); self.downloader_thread.start()
        self.after(100, self.process_queue)

    def process_queue(self):
        try:
            message = self.progress_queue.get_nowait()
            progress_frame = self.frames['ProgressFrame']
            if message['type'] == 'log': progress_frame.add_log_message(message['message'])
            elif message['type'] == 'kpi':
                 progress_frame.kpi_obtained.configure(text=f"🟩 OBTENIDOS\n{message['obtained']}")
                 progress_frame.kpi_failed.configure(text=f"🟥 FALLIDOS\n{message['failed']}")
                 progress_frame.kpi_pending.configure(text=f"🟦 PENDIENTES\n{message['pending']}")
                 total = message['obtained'] + message['failed'] + message['pending']
                 current = message['obtained'] + message['failed']
                 if total > 0: progress_frame.progress_bar.set(current / total); progress_frame.progress_label.configure(text=f"Procesando {current} de {total}... ({current/total*100:.1f}%)")
            elif message['type'] == 'article_result': progress_frame.update_article_status(message['doi'], message['success'], message.get('source'), message.get('reason'))
            elif message['type'] == 'finished':
                progress_frame.finalize_ui()
                self.frames['ResultsFrame'].update_results(message['results'])
                self.show_frame("ResultsFrame")
                return
        except queue.Empty: pass
        finally:
            if self.downloader_thread and self.downloader_thread.is_alive(): self.after(100, self.process_queue)
            elif not self.progress_queue.empty(): self.after(10, self.process_queue)

    def toggle_pause(self):
        if self.pause_event.is_set(): self.pause_event.clear(); self.frames['ProgressFrame'].pause_button.configure(text="Pausar")
        else: self.pause_event.set(); self.frames['ProgressFrame'].pause_button.configure(text="Reanudar")

    def cancel_download(self): self.cancel_event.set(); self.frames['ProgressFrame'].finalize_ui()

    def show_error(self, title, message): messagebox.showerror(title, message)
    def show_info(self, title, message): messagebox.showinfo(title, message)

if __name__ == "__main__":
    app = App()
    app.mainloop()
