import os
import pandas as pd

from extraccion.extraccion import extraer_lecturas
from access.access_reader import leer_tabla_access, hacer_join, obtener_campos_deseados_access,crear_indice_si_no_existe
from export.export import insertar_datos_en_excel_existente
from export.export_datos_final import extraer_y_pegar
from utils.helpers import normalizar_columnas_finales

ruta_excel_final = '../data/output/BE FORMATO (2).xlsx'
hoja_destino = "LECTURAS"
ruta_reporte_final = '../data/output/reporte_final.xlsx'

def ejecutar_proceso_desde_gui(
    rutas_lecturas, 
    ruta_bd_maestro, 
    tabla_maestro,
    ruta_bd_extra,
    tabla_extra, 
    ruta_excel_final, 
    ruta_reporte_final_dir, 
    logger,
    progress_cb=None,
    usar_analisis_extra: bool = False 
):
    from utils.config import mapeo_columnas

    campos_deseados_access = obtener_campos_deseados_access()
    total_files = len(rutas_lecturas) or 1
    steps_per_file = 7
    total_steps = total_files * steps_per_file
    step_count = 0

    def _bump(msg):
        nonlocal step_count
        step_count += 1
        if progress_cb:
            pct = min(100, int(step_count * 100 / total_steps))
            progress_cb(pct, msg)
    
    if progress_cb:
        progress_cb(0, "Preparando...")

    for i,ruta_lecturas in enumerate(rutas_lecturas):

        nombre_archivo = os.path.splitext(os.path.basename(ruta_lecturas))[0]

        logger("\n" + "🔽" * 40)
        logger(f"📄 PROCESANDO ARCHIVO: {nombre_archivo.upper()}")
        logger("🔽" * 40 + "\n")
        
        try:
            # === 1. Leer lecturas desde Excel ===
            logger("📥 Extrayendo datos del Excel...")
            df_lecturas = extraer_lecturas(ruta_lecturas)
            _bump(f"Leyendo Excel ({i+1}/{total_files})")

            cuentas_vacias = df_lecturas[df_lecturas['Cuenta'].isna()]
            if not cuentas_vacias.empty:
                logger(f"⚠️ Atención: {len(cuentas_vacias)} fila(s) sin número de cuenta. Serán procesadas pero sin datos del maestro.")
            logger("✅ Lecturas extraídas\n")

            # === 2. Leer Access Maestro ===
            logger("📥 Conectando a base de datos del Maestro...")
            crear_indice_si_no_existe(ruta_bd_maestro, tabla_maestro, "Cuenta")
            df_access = leer_tabla_access(ruta_bd_maestro, tabla_maestro, campos_deseados_access)
            _bump(f"Maestro leído ({i+1}/{total_files})")
            logger("✅ Conexión con Access exitosa\n")
            
            # === 3. Hacer JOIN con Maestro ===
            logger("🔗 Haciendo JOIN con la BD Maestro...")
            df_final = hacer_join(df_lecturas, df_access, 'Cuenta', 'Cuenta', campos_deseados_access)
            _bump(f"JOIN Maestro ({i+1}/{total_files})")
            logger("✅ JOIN realizado\n")

            # === 4. Ruta de Lectura ===
            logger("🧱 Generando Ruta de Lectura...")
            for c in ["Sector","Zona","Correlativo"]:
                df_final[c] = pd.to_numeric(df_final[c], errors="coerce").fillna(0).astype(int).astype(str)
            df_final["Ruta_de_Lectura"] = (
                df_final["Sector"].fillna(0).astype(int).astype(str) + "-" +
                df_final["Zona"].fillna(0).astype(int).astype(str) + "-" +
                df_final["Correlativo"].fillna(0).astype(int).astype(str)
            )
            _bump(f"Ruta de lectura ({i+1}/{total_files})")
            logger("✅ Ruta de Lectura generada\n")

            # === 5. Mes actual (OPCIONAL - análisis extra) ===
            col_lectura = 'Ultima_Lectura_Terreno'
            col_fecha = 'Fecha_Ultima_Lectura_Terreno'
            col_consumo = 'Ultimo_Consumo'

            MESES_ES = {1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
                        7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"}

            if usar_analisis_extra:
                if all(c in df_final.columns for c in [col_lectura, col_fecha, col_consumo]):

                    fecha_mode = pd.to_datetime(df_final[col_fecha], errors='coerce').dt.month.mode()
                    if not fecha_mode.empty:
                        mes_actual = int(fecha_mode[0])
                        nombre_mes = MESES_ES.get(mes_actual, str(mes_actual))

                        logger(f"📅 (Extra) Revisión actual activada: {nombre_mes}")

                        df_final[f'Ultima Lectura - {nombre_mes}'] = df_final[col_lectura]
                        df_final[f'Fecha Ultima Lectura - {nombre_mes}'] = df_final[col_fecha]
                        df_final[f'Ultimo Consumo - {nombre_mes}'] = df_final[col_consumo]
                    else:
                        logger("⚠️ (Extra) No se pudo detectar el mes de revisión actual.")
                else:
                    logger("⚠️ (Extra) Faltan columnas para crear revisión actual.")
            else:
                logger("ℹ️ (Extra) Análisis desactivado: no se agregan columnas de mes actual.")
            _bump(f"Mes actual ({i+1}/{total_files})")

            # === 6. Agregar columnas desde segunda BD Access (revisión anterior) ===
            campos_extra = ['Cuenta', col_lectura, col_fecha, col_consumo]

            usar_bd_extra = (
                usar_analisis_extra and
                bool(ruta_bd_extra and str(ruta_bd_extra).strip()) and
                bool(tabla_extra and str(tabla_extra).strip())
            )

            if usar_bd_extra: 
                logger("📥 Leyendo revisión anterior desde segunda BD...")
                crear_indice_si_no_existe(ruta_bd_extra, tabla_extra, "Cuenta")
                df_extra = leer_tabla_access(ruta_bd_extra, tabla_extra, campos_extra)

                df_extra['Cuenta'] = df_extra['Cuenta'].astype(str).str.extract(r'(\d+)')
                df_final['Cuenta'] = df_final['Cuenta'].astype(str).str.extract(r'(\d+)')

                df_final = df_final.merge(
                    df_extra,
                    left_on='Cuenta',
                    right_on='Cuenta',
                    how='left',
                    suffixes=('', '_extra')
                )

                fecha_extra = pd.to_datetime(df_final[f'{col_fecha}_extra'], errors='coerce')
                mes_extra = fecha_extra.dt.month.mode()

                if not mes_extra.empty:
                    mes_2 = int(mes_extra[0])
                    nombre_mes_2 = MESES_ES.get(mes_2, str(mes_2))
                    logger(f"📅 Detectado mes de revisión anterior: {nombre_mes_2}")

                    df_final[f'Ultima Lectura - {nombre_mes_2}'] = df_final[f'{col_lectura}_extra']
                    df_final[f'Fecha Ultima Lectura - {nombre_mes_2}'] = df_final[f'{col_fecha}_extra']
                    df_final[f'Ultimo Consumo - {nombre_mes_2}'] = df_final[f'{col_consumo}_extra']
                else:
                    logger("⚠️ No se detectó mes válido en revisión anterior")

                drop_cols = [c for c in df_final.columns if c.endswith('_extra') or c == 'SUM']
                if drop_cols:
                    df_final.drop(columns=drop_cols, inplace=True, errors='ignore')
                    logger("✅ Columnas de revisión comparativa añadidas\n")
                    _bump(f"Revisión anterior ({i+1}/{total_files})")
                else:
                    if usar_analisis_extra:
                        logger("ℹ️ (Extra) BD secundaria no seleccionada. Se omite revisión anterior.\n")
                    else:
                        logger("ℹ️ (Extra) Análisis desactivado: no se consulta BD secundaria.\n")
                    _bump(f"Sin revisión anterior ({i+1}/{total_files})")
            else:
                logger("ℹ️ BD secundaria no seleccionada. Se omite revisión anterior.\n")
                _bump(f"Sin revisión anterior ({i+1}/{total_files})")

            # === 7. Renombrado final de columnas + CSV intermedio ===
            logger("🧽 Renombrando columnas finales...")

            if 'Llave' in df_final.columns:
                df_final['Llave'] = pd.to_numeric(df_final['Llave'], errors='coerce').fillna(0).astype(int)
            else:
                df_final['Llave'] = 0

            df_final = df_final.astype("string").fillna("")

            # Normaliza nombres
            df_final = normalizar_columnas_finales(df_final)

            df_final = df_final.replace({"nan": "", "<NA>": ""}, regex=False)

            logger("✅ Datos listos\n")

            # === 8. Exportar a Excel ===
            logger("📤 Insertando datos en Excel final...")

            if usar_analisis_extra:
                for col in df_final.columns:
                    if col.startswith("Ultima Lectura -") or \
                    col.startswith("Fecha Ultima Lectura -") or \
                    col.startswith("Ultimo Consumo -"):
                        if col not in mapeo_columnas:
                            mapeo_columnas[col] = col

            # al final, guarda un archivo de salida único por input
            ruta_reporte_final = os.path.join(
                ruta_reporte_final_dir,
                f"BE-{nombre_archivo}.xlsx"
            )

            insertar_datos_en_excel_existente(
                ruta_plantilla_excel=ruta_excel_final,
                hoja_destino="LECTURAS",
                df_datos=df_final,
                mapeo_columnas= mapeo_columnas,
                fila_inicio=12,
                ruta_salida= ruta_reporte_final
            )

            logger(f"✅ Reporte generado: {ruta_reporte_final}\n")


            logger("📤 Moviendo datos al excel Final...")

            mapeo_celdas = {
                "fecha": "G13",  
                "sed":   "D7",  
                "TOTALIZADOR": {"FACTOR":   "F22","LECTURA 1":"G22","LECTURA 2":"H22"},
                "ALP1": {"FACTOR":   "F25","LECTURA 1":"G25","LECTURA 2":"H25"}
            }
                            
            logger(f"🧭 extraer_y_pegar(input='{ruta_lecturas}', output='{ruta_reporte_final}')")
            extraer_y_pegar(
                ruta_input = ruta_lecturas,
                hoja_input = "Lecturas",
                ruta_output= ruta_reporte_final,
                hoja_output="BALANCE KWH",
                mapeo_celdas = mapeo_celdas
            )


            logger("✅ Datos movidos con exito")

        except Exception as e:
            logger(f"❌ Error procesando {nombre_archivo}: {e}")

    logger("✅ Todos los archivos fueron procesados correctamente.\n")