import pandas as pd

def encontrar_item_index(df):
    for row_idx in df.index:
        for col_idx in df.columns:
            if str(df.loc[row_idx, col_idx]).strip().upper() == 'ITEM':
                return row_idx, col_idx
    return None, None

def extraer_bloque_agregares(df_all, idx_agregares, header):
    df_agregares = pd.DataFrame()
    if idx_agregares < len(df_all):
        agregares_raw = df_all.iloc[idx_agregares + 1:].dropna(how='all').reset_index(drop=True)
        if not agregares_raw.empty:
            offset_col = agregares_raw.apply(lambda row: row.first_valid_index(), axis=1).mode()[0]
            agregares_block = agregares_raw.iloc[:, offset_col : offset_col + len(header)]
            while agregares_block.shape[1] < len(header):
                agregares_block[agregares_block.shape[1]] = None
            agregares_block = agregares_block.iloc[:, :len(header)]
            agregares_block.columns = header
            df_agregares = agregares_block
    return df_agregares.dropna(axis=1, how='all')

def limpiar_columnas(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
        .str.replace(r'\s+', ' ', regex=True)
        .str.replace('\n', ' ')
    )
    return df

def limpiar_valores_texto(df: pd.DataFrame) -> pd.DataFrame:
    for col in df.select_dtypes(include='object').columns:
        df[col] = df[col].astype(str).str.strip()
    return df

def normalizar_columnas_finales(df):
    renombres = {
        'Cuenta': 'Suministro',
        'Nombre': 'Nombre',
        'Direccion': 'Direccion',
        'Distrito_x': 'Distrito',
        'Estado_Cliente_x': 'Estado_Cliente',
        'Medidor_x': 'Medidor',
        'Marca_x': 'Marca',
        'Fase_x': 'Fase',
        'Factor_x': 'Factor',
        'Cod Act Comerc': 'Giro_de_Sistema',
        'Tarifa': 'Tarifa',
        'Tipo_Acomet': 'Tipo_Acomet',
        'Llave': 'Llave',
        'Ultima Lectura Terreno': 'Ultima_Lectura',
        'Fecha Ultima Lectura Terreno': 'Fecha_Ultima_Lectura',
        'Ultimo Consumo': 'Ultimo_Consumo',
        'Clave_Ult_LectFact': 'Clave_Ult_LectFact',
        'Consumo_8': 'Consumo_9',
        'Consumo_7': 'Consumo_8',
        'Consumo_6': 'Consumo_7',
        'Consumo_5': 'Consumo_6',
        'Consumo_4': 'Consumo_5',
        'Consumo_3': 'Consumo_4',
        'Consumo_2': 'Consumo_3',
        'Consumo_1': 'Consumo_2',
        'Consumo_0': 'Consumo_1',
        'Lectura 1': 'Lectura_1',
        'Lectura 2': 'Lectura_2',
        'Observaciones': 'Observaciones',
        'Giro de Negocio': 'Causal',
        'Código de Giro': 'Giro_de_Sistema_1',
        'Tipo Acomet': 'Tipo_Acomet_1',
        'Tipo de Medidor': 'Tipo_de_Medidor'
    }
    return df.rename(columns=renombres)