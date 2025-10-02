import pandas as pd
import pyodbc
import os
import warnings
warnings.filterwarnings("ignore", message="pandas only supports SQLAlchemy connectable", category=UserWarning)


def obtener_campos_deseados_access():
    return [
        'Cuenta','Distrito','Tarifa', 'Tipo_Acomet', 'Estado_Cliente', 'Medidor', 'Marca','Fase', 'Factor', 
        'CodActComerc', 'ActComercial', 'Sector','Zona','Correlativo','Ultima_Lectura_Terreno', 'Fecha_Ultima_Lectura_Terreno',
        'Ultimo_Consumo', 'Clave_Ult_LectFact','Consumo_8', 'Consumo_7', 'Consumo_6', 'Consumo_5', 'Consumo_4',
        'Consumo_3', 'Consumo_2', 'Consumo_1', 'Consumo_0','Llave'
    ]

def leer_tabla_access(ruta_bd: str, tabla: str, campos:list) -> pd.DataFrame:
    conn_str = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        rf'DBQ={ruta_bd};'
    )
    try:
        conn = pyodbc.connect(conn_str, timeout=5)
        campos_sql = ', '.join(f"[{c}]" for c in campos)
        query = f"SELECT {campos_sql} FROM [{tabla}]"
        df = pd.read_sql(query, conn)
        conn.close()
        duplicated = df.columns[df.columns.duplicated()].tolist()
        if duplicated:
            raise ValueError(f"❌ La base de datos devolvió columnas duplicadas: {duplicated}")
        return df
    except Exception as e:
        raise RuntimeError(f"❌ Error leyendo la base de datos Access: {e}")
    

def hacer_join(df_excel, df_access, clave_excel, clave_access, columnas_deseadas):

    #df_excel[clave_excel] = df_excel[clave_excel].astype(str).str.strip()
    df_excel[clave_excel]  = pd.to_numeric(df_excel[clave_excel],  errors="coerce").astype("Int64")
    #df_access[clave_access] = df_access[clave_access].astype(str).str.strip()
    df_access[clave_access]= pd.to_numeric(df_access[clave_access],errors="coerce").astype("Int64")

    df_access = df_access[df_access[clave_access].notna()]
    campos_unicos = [clave_access] + [c for c in columnas_deseadas if c != clave_access]
    df_access = df_access[campos_unicos]
    #df_excel[clave_excel] = df_excel[clave_excel].str.extract(r'(\d+)')
    #df_access[clave_access] = df_access[clave_access].str.extract(r'(\d+)')

    df_joined = df_excel.merge(df_access, left_on=clave_excel, right_on=clave_access, how="left", suffixes=('', '_access'))
    if clave_access in df_joined.columns and clave_access != clave_excel:
        df_joined.drop(columns=[clave_access], inplace=True)
        
    return df_joined
    '''
    df_access = df_access[df_access[clave_access].notna() & (df_access[clave_access] != 'nan')]

    campos_unicos = [clave_access] + [c for c in columnas_deseadas if c != clave_access]
    df_access = df_access[campos_unicos]

    df_joined = pd.merge(df_excel,
                        df_access, 
                        left_on=clave_excel,
                        right_on=clave_access,
                        how='left',
                        suffixes=('', '_access')
                        )

    if clave_access in df_joined.columns and clave_access != clave_excel:
        df_joined.drop(columns=[clave_access], inplace=True)

    return df_joined
    '''

def crear_indice_si_no_existe(ruta_bd, tabla, columna, nombre_indice="idx_auto"):

    try:
        conn = pyodbc.connect(
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + ruta_bd + ';'
        )
        cursor = conn.cursor()

        # Intentamos crear el índice (si ya existe, se ignora)
        sql = f"CREATE INDEX {nombre_indice}_{columna} ON [{tabla}] ([{columna}]);"
        cursor.execute(sql)
        conn.commit()
        conn.close()
        print(f"✅ Índice '{nombre_indice}_{columna}' creado en tabla '{tabla}'.")
    except Exception as e:
        if "ya tiene un índice llamado" not in str(e):
            raise 