import pandas as pd
import os
import warnings
from utils.helpers import(
    encontrar_item_index,
    limpiar_columnas,
    limpiar_valores_texto,
    extraer_bloque_agregares
)
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

def extraer_lecturas(path_excel: str) -> pd.DataFrame:

    df_all = pd.read_excel(path_excel, header=None)

    idx_main_row, idx_main_col = encontrar_item_index(df_all)
    if idx_main_row is None:
        raise ValueError("❌ No se encontró 'Item' en el archivo.")

    # localizar bloque 'AGREGARES'
    idx_agregares = df_all.iloc[idx_main_row + 1:].apply(
        lambda row: row.astype(str).str.upper().str.contains("AGREGARES", na=False).any(),
        axis=1
    )
    idx_agregares = idx_agregares[idx_agregares].index
    idx_agregares = idx_agregares[0] if not idx_agregares.empty else len(df_all)   

    # header base
    header = df_all.iloc[idx_main_row, idx_main_col:].astype(str)
    header = (
        header
        .str.strip()
        .str.replace(r'\s+', ' ', regex=True)
        .str.replace('\n', ' ')
        .str.strip()
        .tolist()
    )
    # bloque principal
    main_block = df_all.iloc[idx_main_row + 1 : idx_agregares, idx_main_col : idx_main_col + len(header)]
    main_block = main_block.dropna(how='all')
    # quitar filas que son solo números (totales)
    main_block = main_block[~main_block.apply(
        lambda row: row.dropna().astype(str).str.fullmatch(r'\d+(\.\d+)?').all(), axis=1
    )]

    #⚠️ SIEMPRE define df_main (aunque no haya filas)
    if main_block.empty:
        df_main = pd.DataFrame(columns=header)
    else:
        df_main = pd.DataFrame(main_block.values, columns=header)

     # --- sanea df_main ---
    df_main.columns = df_main.columns.map(lambda x: str(x).strip())
    df_main = df_main.dropna(axis=1, how='all')
    # si hay columnas duplicadas, conserva la primera
    if df_main.columns.duplicated().any():
        df_main = df_main.loc[:, ~df_main.columns.duplicated(keep='first')]

    df_agregares = extraer_bloque_agregares(df_all, idx_agregares, header)
    if not df_agregares.empty:
        df_agregares.columns = df_agregares.columns.map(lambda x: str(x).strip())
        df_agregares = df_agregares.dropna(axis=1, how='all')
        if df_agregares.columns.duplicated().any():
            df_agregares = df_agregares.loc[:, ~df_agregares.columns.duplicated(keep='first')]
    

    df_final = pd.concat([df_main, df_agregares], ignore_index=True, sort=False)

     # --- sanea df_final ---
    df_final.columns = df_final.columns.map(lambda x: str(x).strip())
    df_final = df_final.dropna(axis=1, how='all')
    if df_final.columns.duplicated().any():
        df_final = df_final.loc[:, ~df_final.columns.duplicated(keep='first')]

    df_final = limpiar_columnas(df_final)
    df_final = limpiar_valores_texto(df_final)

    
    for col in df_final.select_dtypes(include='object').columns:
        df_final[col] = df_final[col].astype(str).str.strip()

    columnas_deseadas = [
        'Item', 'SED', 'Cuenta', 'Nombre', 'Direccion',
        'Medidor', 'Marca', 'Fase', 'Factor', 'Cod Act Comerc',
        'Lectura 1', 'Lectura 2', 'Consumo 1 y 2', 'Observaciones',
        'Giro de Negocio', 'Código de Giro', 'Tipo de Medidor', 'Tipo Acomet'
    ]

    columnas_existentes = [col for col in columnas_deseadas if col in df_final.columns]
    df_final = df_final[columnas_existentes].reset_index(drop=True)
    return df_final
