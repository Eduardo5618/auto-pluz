# src/GUI/gui_fotos.py
import os
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
    ttk.Label(frm, text="RT_1 (Carpeta de INICIO):").grid(row=2, column=0, sticky="w")
    ttk.Entry(frm, textvariable=v_ini, width=62).grid(row=2, column=1, padx=6)
    ttk.Button(frm, text="Carpeta", command=lambda: pick_dir(v_ini)).grid(row=2, column=2)

    ttk.Separator(frm).grid(row=3, column=0, columnspan=3, sticky="ew", pady=8)

    # CIERRE
    ttk.Label(frm, text="RT_2 (Carpeta de CIERRE):").grid(row=4, column=0, sticky="w")
    ttk.Entry(frm, textvariable=v_cie, width=62).grid(row=4, column=1, padx=6)
    ttk.Button(frm, text="Carpeta", command=lambda: pick_dir(v_cie)).grid(row=4, column=2)  

    ttk.Checkbutton(
        frm,
        text="Limpiar imágenes previas en celdas destino",
        variable=v_limpiar
    ).grid(row=5, column=0, columnspan=3, sticky="w", pady=(6,0))

    ttk.Separator(frm).grid(row=6, column=0, columnspan=3, sticky="ew", pady=8)

    btn = ttk.Button(frm, text="Iniciar proceso")
    btn.grid(row=7, column=0, columnspan=3, pady=4)

    txt = tk.Text(frm, height=10, width=90, state="disabled")
    txt.grid(row=8, column=0, columnspan=3, pady=(8,0))

    def run():
        try:
            xlsx, ini, cie = v_excel.get().strip(), v_ini.get().strip(), v_cie.get().strip()
            if not (xlsx and ini and cie):
                messagebox.showwarning("Campos incompletos", "Completa Excel e INICIO/CIERRE.")
                return

            btn.configure(state="disabled")
            txt.configure(state="normal"); txt.delete("1.0","end"); txt.configure(state="disabled")
            res = procesar_fotos_predefinidos(
                ruta_excel=xlsx,
                carpeta_inicio=ini,
                carpeta_cierre=cie,
                limpiar_previas=bool(v_limpiar.get()),
                img_w=157,
                img_h=210,
                logger=log
            )
            

            msg = (f"🏁 RT_1: {res['RT_1']['ok']} ok / {res['RT_1']['err']} err | "
                   f"RT_2: {res['RT_2']['ok']} ok / {res['RT_2']['err']} err")
            log(msg)
            status_cb(msg)
            messagebox.showinfo("Proceso finalizado", msg)

        except Exception as e:
            log(f"⚠️ Error: {e}")
            messagebox.showerror("Error", str(e))
        finally:
            btn.configure(state="normal")

    btn.configure(command=run)