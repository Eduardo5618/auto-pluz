import os
import shutil
import win32com.client as win32


def insertar_datos_en_excel_existente(
    ruta_plantilla_excel,
    hoja_destino,
    df_datos,
    mapeo_columnas,
    fila_inicio=12,
    ruta_salida=None
):
    
    if not os.path.exists(ruta_plantilla_excel):
        raise FileNotFoundError(f"❌ No se encontró el archivo: {ruta_plantilla_excel}")

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

    xl = win32.Dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    
    try:
        wb = xl.Workbooks.Open(os.path.abspath(ruta_final))
        # Verifica hoja
        nombres_hojas = [ws.Name for ws in wb.Worksheets]
        if hoja_destino not in nombres_hojas:
            raise ValueError(f"❌ La hoja '{hoja_destino}' no existe.")
        ws = wb.Worksheets(hoja_destino)

        # Encabezados en fila (fila_inicio - 1)
        fila_enc = fila_inicio - 1
        last_col = ws.Cells(fila_enc, ws.Columns.Count).End(-4159).Column
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

        # Función para escribir UNA columna con tupla de tuplas
        def write_col(col_name_df, etiqueta_destino):
            if etiqueta_destino not in mapa_destino or col_name_df not in df.columns:
                return
            col_idx_excel = mapa_destino[etiqueta_destino]
            vals = df[col_name_df].tolist()
            # Excel quiere tupla de tuplas para (n x 1)
            data_tuple = tuple((v if v != "" else None,) for v in vals)
            start = ws.Cells(fila_inicio, col_idx_excel)
            end = ws.Cells(fila_inicio + nrows - 1, col_idx_excel)
            rng = ws.Range(start, end)
            rng.Value = data_tuple
            print(f"✍️  Escrito '{etiqueta_destino}' → col {col_idx_excel}, {nrows} filas")

        # Escribe columnas mapeadas
        for col_origen, col_destino in mapeo_columnas.items():
            write_col(col_origen, col_destino)

        # Escribe dinámicas
        for col in columnas_dinamicas:
            write_col(col, col)

        wb.Save()
    finally:
        try:
            wb.Close(SaveChanges=True)
        except Exception:
            pass
        xl.Quit()
        xl = None

    print(f"✅ Datos insertados correctamente en '{hoja_destino}' desde fila {fila_inicio}.")