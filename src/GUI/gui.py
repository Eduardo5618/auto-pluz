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
from GUI.gui_fotos import abrir_popup_fotos


class App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Generador de Reporte de Lecturas")
        self.configure(bg="#f7f7f7")

        
        w, h = 620, 660
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (w // 2)
        y = (self.winfo_screenheight() // 2) - (h // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")

        self._setup_style() 
        self.resizable(False, False)
        self.iconbitmap(default="data/gui/icono.ico") 

        self.rutas_lecturas = [] 
        self.ruta_bd_maestro = tk.StringVar()
        self.ruta_bd_extra = tk.StringVar()
        self.tabla_maestro = tk.StringVar()
        self.tabla_extra = tk.StringVar()

        self._crear_widgets()

        self.protocol("WM_DELETE_WINDOW", self._on_close)
    
    def _on_close(self):
        try:
            self.destroy()    
        finally:
            os._exit(0)     
            
    def _crear_widgets(self):
        
        bg = "#f7f7f7"
        self.configure(bg=bg)

        # --- Encabezado ---
        header = tk.Frame(self, bg=bg)
        header.pack(fill="x", padx=10, pady=(8, 4))

        try:
            img = Image.open("data/gui/logo.png").resize((140, 60))
            logo = ImageTk.PhotoImage(img)
            tk.Label(header, image=logo, bg=bg).pack()
            self.logo_img = logo
        except:
            tk.Label(self, text="🔧 Reporte de Lecturas", font=("Segoe UI", 16, "bold"), bg="#f7f7f7").pack(pady=10)

        
        # --- Cuerpo ---
        body = tk.Frame(self, bg=bg)
        body.pack(fill="both", expand=True, padx=10, pady=4)

        # ====== Sección: Fuentes ======
        sec1 = ttk.Labelframe(body, text="Fuentes")
        sec1.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        sec1.grid_columnconfigure(1, weight=1)
        

        ttk.Label(sec1, text="Archivo de Lecturas:").grid(row=0, column=0, sticky="w", pady=3)
        ttk.Label(sec1, text="(seleccionados en consola)").grid(row=0, column=1, sticky="w", pady=3)
        ttk.Button(sec1, text="Buscar", command=self._buscar_lecturas).grid(row=0, column=2, padx=(6,0))

        ttk.Label(sec1, text="BD Access (Maestro):").grid(row=1, column=0, sticky="w", pady=3)
        ttk.Entry(sec1, textvariable=self.ruta_bd_maestro).grid(row=1, column=1, sticky="ew", pady=3)
        ttk.Button(sec1, text="Buscar", command=lambda: self._buscar_bd('maestro')).grid(row=1, column=2, padx=(6,0))

        ttk.Label(sec1, text="BD Access (Secundaria):").grid(row=2, column=0, sticky="w", pady=3)
        ttk.Entry(sec1, textvariable=self.ruta_bd_extra).grid(row=2, column=1, sticky="ew", pady=3)
        ttk.Button(sec1, text="Buscar", command=lambda: self._buscar_bd('extra')).grid(row=2, column=2, padx=(6,0))

        ttk.Separator(body).grid(row=1, column=0, sticky="ew", pady=4)


        # ====== Sección: Tablas ======
        sec2 = ttk.Labelframe(body, text="Tablas")
        sec2.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        sec2.grid_columnconfigure(1, weight=1)

        ttk.Label(sec2, text="Tabla de Maestro:").grid(row=0, column=0, sticky="w", pady=3)
        self.combo_maestro = ttk.Combobox(sec2, textvariable=self.tabla_maestro, state="readonly")
        self.combo_maestro.grid(row=0, column=1, sticky="ew", pady=3)

        ttk.Label(sec2, text="Tabla de BD Extra:").grid(row=1, column=0, sticky="w", pady=3)
        self.combo_extra = ttk.Combobox(sec2, textvariable=self.tabla_extra, state="readonly")
        self.combo_extra.grid(row=1, column=1, sticky="ew", pady=3)

        ttk.Separator(body).grid(row=3, column=0, sticky="ew", pady=4)


        # ====== Sección: Acciones ======
        actions = ttk.Frame(body)
        actions.grid(row=4, column=0, sticky="ew", pady=(0, 8))
        actions.grid_columnconfigure(0, weight=1)

        bar = ttk.Frame(actions)
        bar.grid(row=0, column=0, sticky="w")

        # Ejecutar
        tk.Button(bar, text="▶ Ejecutar Proceso",
                bg="#22a22a", fg="white", font=("Segoe UI", 10, "bold"),
                padx=14, pady=6,
                command=self._ejecutar_thread).pack(side="left", padx=(0, 8))
        
        # Manejo Fotos
        tk.Button(bar, text="📷 Manejo Fotos…",
                bg="#4da6ff", fg="white", font=("Segoe UI", 10, "bold"),
                padx=12, pady=6,
                command=self._abrir_popup_fotos).pack(side="left")
        
                # Manejo Fotos
        tk.Button(bar, text="🧹 Limpiar registro",
                bg="#FF0000", fg="white", font=("Segoe UI", 10, "bold"),
                padx=12, pady=6,
                activebackground="#be123c", activeforeground="white",
                command=self._limpiar_registro
                ).pack(side="left", padx=(8, 0))
        
        # ====== Sección: Registro ======
        sec4 = ttk.Labelframe(body, text="Registro")
        sec4.grid(row=5, column=0, sticky="nsew")
        body.grid_rowconfigure(5, weight=1)

        log_frame = tk.Frame(sec4)
        log_frame.pack(fill="both", expand=True, padx=6, pady=6)

        self.text_log = tk.Text(log_frame, height=10, bg="#f0f0f0", font=("Consolas", 10), wrap="word")
        ysb = ttk.Scrollbar(log_frame, orient="vertical", command=self.text_log.yview)
        self.text_log.configure(yscrollcommand=ysb.set)
        self.text_log.pack(side="left", fill="both", expand=True)
        ysb.pack(side="right", fill="y")

        # ====== Barra de estado ======
        status = tk.Frame(self, bg=bg)
        status.pack(fill="x", padx=10, pady=(6, 8))

        self.progress_var = tk.IntVar(value=0)
        self.progress = ttk.Progressbar(status, orient="horizontal", mode="determinate",
                                        variable=self.progress_var, length=220)
        self.progress.pack(side="left", padx=(0, 10))

        self.progress_label = ttk.Label(status, text="Listo")
        self.progress_label.pack(side="left")

        ttk.Label(status, text="Autor: Daniel Paredes", foreground="#6b7280").pack(side="right")

    
    def _limpiar_registro(self):
        if not messagebox.askyesno("Confirmar", "¿Limpiar todos los campos y el registro?"):
            return

        # Limpia listas/vars
        self.rutas_lecturas.clear()
        self.ruta_bd_maestro.set("")
        self.ruta_bd_extra.set("")
        self.tabla_maestro.set("")
        self.tabla_extra.set("")

        # Resetea combos
        try:
            self.combo_maestro.set("")
            self.combo_maestro['values'] = ()
        except Exception:
            pass
        try:
            self.combo_extra.set("")
            self.combo_extra['values'] = ()
        except Exception:
            pass

        # Limpia log y progreso
        self.text_log.delete("1.0", tk.END)
        self.set_progress(0, "Listo")

        # Mensaje en log
        self.log("🧹 Registro limpiado. Archivos seleccionados: 0")

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

    def _setup_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")  # aspecto limpio y consistente
        except:
            pass

        # Botón primario (verde) para acciones principales
        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), padding=6)
        style.map("Primary.TButton",
                foreground=[("!disabled", "white")])

        # Campos
        style.configure("TEntry", padding=3)
        style.configure("TLabelframe.Label", font=("Segoe UI", 10, "bold"))

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

    def _abrir_popup_fotos(self):
    # El popup actualizará la barra de estado usando este callback
        abrir_popup_fotos(
            parent=self,
            status_cb=lambda msg: self.set_progress(self.progress_var.get(), msg)
        )

if __name__ == "__main__":
    app = App()
    app.mainloop()
