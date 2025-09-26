import os
import shutil
import time
import win32com.client as win32
import pythoncom

def _retry_com(callable_fn, *args, _retries=6, _sleep=0.5, **kwargs):

    for i in range(_retries):
        try:
            return callable_fn(*args, **kwargs)
        except Exception as e:
            msg = str(e)
            if "-2147418111" in msg or "rechazada" in msg or "rejected by the callee" in msg:
                time.sleep(_sleep)
                continue
            raise
    return callable_fn(*args, **kwargs)


def insertar_datos_en_excel_existente(
    ruta_plantilla_excel: str,
    hoja_destino: str,
    df_datos,
    mapeo_columnas,
    fila_inicio: int=12,
    ruta_salida: str | None=None
):
    
    if not os.path.exists(ruta_plantilla_excel):
        raise FileNotFoundError(f"❌ No se encontró el archivo: {ruta_plantilla_excel}")
    
    # Si vamos a escribir a otra ruta, hacemos copia para no tocar el original
    ruta_final = ruta_salida if ruta_salida and ruta_salida != ruta_plantilla_excel else ruta_plantilla_excel
    if ruta_final != ruta_plantilla_excel:
        os.makedirs(os.path.dirname(ruta_final), exist_ok=True)
        shutil.copy(ruta_plantilla_excel, ruta_final)

    # Limpieza de NaN/'nan'
    df = df_datos.copy()
    if 'Item' not in df.columns:
        df.insert(0, 'Item', range(1, len(df) + 1))
    df = df.fillna("").astype(str).replace("nan", "", regex=False)

    nrows = len(df)
    print(f"🧮 Filas a escribir: {nrows}")
    if nrows == 0:
        print("⚠️ No hay filas para escribir. Salgo.")
        return

    pythoncom.CoInitialize()
    xl = None
    wb = None
    
    try:
        xl = win32.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.EnableEvents = False

        ruta_abs = os.path.abspath(ruta_final)

        wb = _retry_com(
            xl.Workbooks.Open,
            ruta_abs,
            UpdateLinks=False, ReadOnly=False, Notify=False
        )

        # Hoja destino
        nombres_hojas = [ws.Name for ws in wb.Worksheets]
        if hoja_destino not in nombres_hojas:
            raise ValueError(f"❌ La hoja '{hoja_destino}' no existe en '{os.path.basename(ruta_abs)}'.")
        ws = wb.Worksheets(hoja_destino)

        # Encabezados en fila (fila_inicio - 1)
        fila_enc = fila_inicio - 1
        xlDirectionLeft = -4159 
        last_col = ws.Cells(fila_enc, ws.Columns.Count).End(xlDirectionLeft).Column
        if ws.Cells(fila_enc, last_col).Value is None:
            last_col = 0

        encabezado_excel = [ws.Cells(fila_enc, c).Value for c in range(1, last_col + 1)]
        mapa_destino = {nombre: idx + 1 for idx, nombre in enumerate(encabezado_excel) if nombre}

        # Intento por Nombre Definido si no está en encabezado
        for col_destino in mapeo_columnas.values():
            if col_destino not in mapa_destino:
                try:
                    nm = wb.Names(col_destino)  # puede fallar si no existe
                    ref_rng = nm.RefersToRange
                    if ref_rng.Worksheet.Name == hoja_destino:
                        mapa_destino[col_destino] = ref_rng.Column
                except Exception:
                    pass

        # Agregar columnas dinámicas (mensuales) si no existen
        patrones = ("Ultima Lectura -", "Fecha Ultima Lectura -", "Ultimo Consumo -")
        columnas_dinamicas = [c for c in df.columns if c.startswith(patrones)]
        col_actual = last_col
        for col in columnas_dinamicas:
            if col not in mapa_destino:
                col_actual += 1
                ws.Cells(fila_enc, col_actual).Value = col
                mapa_destino[col] = col_actual
                print(f"➕ Dinámica '{col}' en columna {col_actual}")

        # Writer robusto (con retry)
        def write_col(col_name_df, etiqueta_destino):
            if etiqueta_destino not in mapa_destino or col_name_df not in df.columns:
                return
            col_idx_excel = mapa_destino[etiqueta_destino]
            vals = df[col_name_df].tolist()
            data_tuple = tuple((v if v != "" else None,) for v in vals)
            start = ws.Cells(fila_inicio, col_idx_excel)
            end = ws.Cells(fila_inicio + nrows - 1, col_idx_excel)
            rng = ws.Range(start, end)
            _retry_com(setattr, rng, "Value", data_tuple)
            print(f"✍️  Escrito '{etiqueta_destino}' → col {col_idx_excel}, {nrows} filas")

        # Escribe columnas mapeadas
        for col_origen, col_destino in mapeo_columnas.items():
            write_col(col_origen, col_destino)

        # Escribe dinámicas
        for col in columnas_dinamicas:
            write_col(col, col)

        _retry_com(wb.Save)
        print(f"✅ Datos insertados correctamente en {hoja_destino}")
    finally:
        try:
            if wb is not None:
                _retry_com(wb.Close,SaveChanges=True)            
        except Exception:
            pass

        try:
            if xl is not None:
                _retry_com(xl.Quit())
        except Exception:
            pass
        pythoncom.CoUninitialize()
