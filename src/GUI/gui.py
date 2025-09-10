import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pyodbc
import threading
import sys
from PIL import Image, ImageTk
import os
import subprocess
import pythoncom
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from main import ejecutar_proceso_desde_gui
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Generador de Reporte de Lecturas")
        self.configure(bg="#f7f7f7")
        self.minsize(820, 760)
        self.resizable(True, True)
        self.iconbitmap(default="data/gui/icono.ico") 

        self.rutas_lecturas = [] 
        self.ruta_bd_maestro = tk.StringVar()
        self.ruta_bd_extra = tk.StringVar()
        self.tabla_maestro = tk.StringVar()
        self.tabla_extra = tk.StringVar()

        self._crear_widgets()

    def _crear_widgets(self):

        try:
            img = Image.open("data/gui/logo.png")
            img = img.resize((140, 60))
            logo = ImageTk.PhotoImage(img)
            tk.Label(self, image=logo, bg="#f7f7f7").pack(pady=5)
            self.logo_img = logo  
        except:
            tk.Label(self, text="🔧 Reporte de Lecturas", font=("Segoe UI", 16, "bold"), bg="#f7f7f7").pack(pady=10)

        padx, pady = 8, 4

        # === Selección de Archivos ===
        frm = tk.Frame(self, bg="#f7f7f7")
        frm.pack(padx=10, pady=5, fill="x")
        frm.grid_columnconfigure(1, weight=1)

        tk.Label(frm, text="📘 Archivo de Lecturas:").grid(row=0, column=0, sticky="w")
        tk.Label(frm, text="(Seleccionados en consola)").grid(row=0, column=1, sticky="w")
        tk.Button(frm, text="Buscar", command=self._buscar_lecturas).grid(row=0, column=2)

        tk.Label(frm, text="📙 BD Access (Maestro):").grid(row=1, column=0, sticky="w")
        tk.Entry(frm, textvariable=self.ruta_bd_maestro, width=70).grid(row=1, column=1)
        tk.Button(frm, text="Buscar", command=lambda: self._buscar_bd('maestro')).grid(row=1, column=2)

        tk.Label(frm, text="📗 BD Access (Secundaria):").grid(row=2, column=0, sticky="w")
        tk.Entry(frm, textvariable=self.ruta_bd_extra, width=70).grid(row=2, column=1)
        tk.Button(frm, text="Buscar", command=lambda: self._buscar_bd('extra')).grid(row=2, column=2)

        tk.Label(frm, text="📋 Tabla de Maestro:").grid(row=3, column=0, sticky="w")
        self.combo_maestro = ttk.Combobox(frm, textvariable=self.tabla_maestro, width=67, state="readonly")
        self.combo_maestro.grid(row=3, column=1, columnspan=2, sticky="w")

        tk.Label(frm, text="📋 Tabla de BD Extra:").grid(row=4, column=0, sticky="w")
        self.combo_extra = ttk.Combobox(frm, textvariable=self.tabla_extra, width=67, state="readonly")
        self.combo_extra.grid(row=4, column=1, columnspan=2, sticky="w")

        tk.Button(self, text="▶ Ejecutar Proceso", bg="green", fg="white", command=self._ejecutar_thread).pack(pady=10)

        log_frame = tk.Frame(self, bg="#f7f7f7")
        log_frame.pack(padx=padx, pady=pady, fill="both", expand=True)

        self.text_log = tk.Text(log_frame, height=10, bg="#f0f0f0", font=("Consolas", 10), wrap="word")
        ysb = ttk.Scrollbar(log_frame, orient="vertical", command=self.text_log.yview)
        self.text_log.configure(yscrollcommand=ysb.set)

        self.text_log.pack(side="left", fill="both", expand=True)
        ysb.pack(side="right", fill="y")

        # === Barra de Progreso ===
        self.progress_var = tk.IntVar(value=0)
        self.progress = ttk.Progressbar(self, orient="horizontal", mode="determinate", variable=self.progress_var)
        self.progress.pack(padx=10, pady=(5, 6), fill="x")

        self.progress_label = tk.Label(self, text="Listo", anchor="w", bg="#f7f7f7")
        self.progress_label.pack(padx=10, pady=(0, 10), fill="x")
        
        # --- Firma / footer ---
        footer = tk.Frame(self, bg="#f7f7f7")
        footer.pack(side="bottom", fill="x")

        self.author_label = tk.Label(
            footer,
            text="Autor: Daniel Paredes",
            bg="#f7f7f7",       # usa el mismo fondo de la ventana
            fg="#6b7280",       # gris suave
            font=("Segoe UI", 9, "italic")
        )
        self.author_label.pack(side="right", padx=10, pady=6)

    def _buscar_lecturas(self):
        paths = filedialog.askopenfilenames(filetypes=[("Archivos Excel", "*.xlsx")])
        if paths:
            nuevos = [p for p in paths if p not in self.rutas_lecturas]  # Evita duplicados
            self.rutas_lecturas.extend(nuevos)
            self.log(f"🗂 Total de archivos seleccionados: {len(self.rutas_lecturas)}")
            for p in self.rutas_lecturas:
                self.log(f"  📄 {os.path.basename(p)}")

    def _buscar_bd(self, tipo):
        path = filedialog.askopenfilename(filetypes=[("Bases de datos Access", "*.mdb *.accdb")])
        if path:
            if tipo == 'maestro':
                self.ruta_bd_maestro.set(path)
                self._cargar_tablas(path, self.combo_maestro)
            elif tipo == 'extra':
                self.ruta_bd_extra.set(path)
                self._cargar_tablas(path, self.combo_extra)

    def _cargar_tablas(self, ruta_bd, combo):
        try:
            conn = pyodbc.connect(
                r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + ruta_bd + ';', timeout=5)
            cursor = conn.cursor()
            tablas = [row.table_name for row in cursor.tables(tableType='TABLE')]
            combo['values'] = tablas
            if tablas:
                combo.current(0)
            conn.close()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron cargar las tablas: {e}")



    def log(self, mensaje):
        def _append():
            self.text_log.insert(tk.END, mensaje + "\n")
            self.text_log.see(tk.END)
        self.after(0, _append)

    def _ejecutar_thread(self):
        hilo = threading.Thread(target=self._ejecutar_proceso)
        hilo.start()

    def _ejecutar_proceso(self):
        
        try:
            pythoncom.CoInitialize() 

            self.set_progress(0, "Iniciando...")

            ejecutar_proceso_desde_gui(
                rutas_lecturas=self.rutas_lecturas,
                ruta_bd_maestro=self.ruta_bd_maestro.get(),
                tabla_maestro=self.tabla_maestro.get(),
                ruta_bd_extra=self.ruta_bd_extra.get(),      
                tabla_extra=self.tabla_extra.get(),        
                ruta_excel_final=os.path.join("data", "formato", "BE FORMATO.xlsx"),
                ruta_reporte_final_dir=os.path.join("data", "output"),
                logger=self.log,
                progress_cb=self.set_progress
            )
            self.set_progress(100, "Completado")
            carpeta = os.path.abspath(os.path.join("data", "output"))
            subprocess.Popen(f'explorer "{carpeta}"')
        except Exception as e:
            self.log(f"❌ Error: {e}")
            messagebox.showerror("Error", str(e))
        finally:
            try:
                pythoncom.CoUninitialize()
            except:
                pass

    def set_progress(self, value: int, text: str = None):
        def _set():
            v = max(0, min(100, int(value)))
            self.progress_var.set(v)
            if text is not None:
                self.progress_label.config(text=f"{v}% - {text}")
            self.update_idletasks()
        self.after(0, _set)

if __name__ == "__main__":
    app = App()
    app.mainloop()
