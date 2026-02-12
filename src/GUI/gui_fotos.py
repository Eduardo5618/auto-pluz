import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from mov_img.insertar_imagenes import  procesar_fotos_predefinidos

def abrir_popup_fotos(parent, status_cb=lambda s: None):
    win = tk.Toplevel(parent)
    win.title("Copiado/Inserción de Fotos (por patrón)")
    win.resizable(False, False)

    win.transient(parent);win.grab_set();win.focus_force();win.lift()

    win.update_idletasks()
    px, py = parent.winfo_rootx(), parent.winfo_rooty()
    pw, ph = parent.winfo_width(), parent.winfo_height()
    ww, wh = 720, 460
    x = px + (pw - ww)//2
    y = py + (ph - wh)//2
    win.geometry(f"{ww}x{wh}+{x}+{y}")

    # Vars
    v_excel = tk.StringVar()
    v_ini   = tk.StringVar()
    v_cie   = tk.StringVar()
    v_limpiar = tk.BooleanVar(value=False)
    v_do_ini = tk.BooleanVar(value=True)   
    v_do_cie = tk.BooleanVar(value=True) 

    # UI
    frm = ttk.Frame(win, padding=12); frm.grid(row=0, column=0, sticky="nsew")

    def pick_excel():
        f = filedialog.askopenfilename(
            title="Selecciona archivo Excel destino",
            filetypes=[("Excel .xlsx","*.xlsx")]
        )
        if f: v_excel.set(f)

    def pick_dir(var):
        d = filedialog.askdirectory(title="Selecciona carpeta")
        if d: var.set(d)

    def log(msg):
        txt.configure(state="normal")
        txt.insert("end", msg + "\n")
        txt.see("end")
        txt.configure(state="disabled")

    # Excel
    ttk.Label(frm, text="Archivo destino (.xlsx):").grid(row=0, column=0, sticky="w")
    ttk.Entry(frm, textvariable=v_excel, width=62).grid(row=0, column=1, padx=6)
    ttk.Button(frm, text="Buscar", command=pick_excel).grid(row=0, column=2)

    ttk.Separator(frm).grid(row=1, column=0, columnspan=3, sticky="ew", pady=8)

    # INICIO
    ttk.Checkbutton(frm, text="Procesar RT_1 (INICIO)", variable=v_do_ini)\
        .grid(row=2, column=0, columnspan=3, sticky="w", pady=(0, 2))

    ttk.Label(frm, text="RT_1 (Carpeta de INICIO):").grid(row=3, column=0, sticky="w")
    ttk.Entry(frm, textvariable=v_ini, width=62).grid(row=3, column=1, padx=6)
    ttk.Button(frm, text="Carpeta", command=lambda: pick_dir(v_ini)).grid(row=3, column=2)

    ttk.Separator(frm).grid(row=4, column=0, columnspan=3, sticky="ew", pady=8)

    # CIERRE
    ttk.Checkbutton(frm, text="Procesar RT_2 (CIERRE)", variable=v_do_cie)\
        .grid(row=5, column=0, columnspan=3, sticky="w", pady=(0, 2))

    ttk.Label(frm, text="RT_2 (Carpeta de CIERRE):").grid(row=6, column=0, sticky="w")
    ttk.Entry(frm, textvariable=v_cie, width=62).grid(row=6, column=1, padx=6)
    ttk.Button(frm, text="Carpeta", command=lambda: pick_dir(v_cie)).grid(row=6, column=2)

    ttk.Checkbutton(
        frm,
        text="Limpiar imágenes previas en celdas destino",
        variable=v_limpiar
    ).grid(row=7, column=0, columnspan=3, sticky="w", pady=(6, 0))

    ttk.Separator(frm).grid(row=8, column=0, columnspan=3, sticky="ew", pady=8)

    btn = ttk.Button(frm, text="Iniciar proceso")
    btn.grid(row=9, column=0, columnspan=3, pady=4)

    txt = tk.Text(frm, height=10, width=90, state="disabled")
    txt.grid(row=10, column=0, columnspan=3, pady=(8, 0))

    def run():
        try:
            xlsx = v_excel.get().strip()
            ini  = v_ini.get().strip()
            cie  = v_cie.get().strip()
            
            do_ini = bool(v_do_ini.get())
            do_cie = bool(v_do_cie.get())


            if not xlsx:
                messagebox.showwarning("Campos incompletos", "Selecciona el Excel destino.")
                return

            if not do_ini and not do_cie:
                messagebox.showwarning("Nada seleccionado", "Activa RT_1 y/o RT_2 para ejecutar.")
                return

            if do_ini and not ini:
                messagebox.showwarning("Campos incompletos", "Selecciona carpeta de INICIO (RT_1).")
                return

            if do_cie and not cie:
                messagebox.showwarning("Campos incompletos", "Selecciona carpeta de CIERRE (RT_2).")
                return

            btn.configure(state="disabled")
            txt.configure(state="normal"); txt.delete("1.0","end"); txt.configure(state="disabled")

            res = procesar_fotos_predefinidos(
                ruta_excel=xlsx,    
                carpeta_inicio=ini if do_ini else None,
                carpeta_cierre=cie if do_cie else None,
                procesar_inicio=do_ini,
                procesar_cierre=do_cie,
                limpiar_previas=bool(v_limpiar.get()),
                img_w=157,
                img_h=210,
                logger=log
            )
            

            msg = (f"🏁 RT_1: {res.get('RT_1', {}).get('ok', 0)} ok / {res.get('RT_1', {}).get('err', 0)} err | "
                f"RT_2: {res.get('RT_2', {}).get('ok', 0)} ok / {res.get('RT_2', {}).get('err', 0)} err")
            
            log(msg)
            status_cb(msg)
            messagebox.showinfo("Proceso finalizado", msg)

        except Exception as e:
            log(f"⚠️ Error: {e}")
            messagebox.showerror("Error", str(e))
        finally:
            btn.configure(state="normal")

    btn.configure(command=run)