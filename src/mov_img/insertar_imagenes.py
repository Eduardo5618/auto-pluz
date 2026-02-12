import os
import unicodedata
from difflib import SequenceMatcher
from typing import Callable, Dict, List, Optional
import pythoncom, win32com.client as win32

try:
    from PIL import Image, ImageOps
    _PIL_OK = True
except Exception:
    _PIL_OK = False

# ======== Config por defecto ========
IMG_WIDTH  = 157
IMG_HEIGHT = 210

EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff"}

DESTINOS_DEF: Dict[str, List[str]] = {
    "RT_1": ["G3", "G5", "F3", "F5"],
    "RT_2": ["G3", "G5", "F3", "F5"],
}

OBJETIVOS_DEF: Dict[str, List[str]] = {
    "RT_1": ["carga primaria r", "carga primaria t", "carga secundaria r", "carga secundaria t"],
    "RT_2": ["carga primaria r", "carga primaria t", "carga secundaria r", "carga secundaria t"],
}

# ======== Utilidades básicas ========
def _normalize(s: str) -> str:
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.lower()
    for ch in (" ", "-", "_", ".", ","):
        s = s.replace(ch, "")
    return s

def _score(a: str, b: str) -> float:
    na, nb = _normalize(a), _normalize(b)
    if nb in na:
        return 1.0
    return SequenceMatcher(None, na, nb).ratio()

def _listar_imagenes(carpeta: str) -> List[str]:
    return sorted(
        os.path.join(carpeta, fn)
        for fn in os.listdir(carpeta)
        if os.path.splitext(fn.lower())[1] in EXTS
    )

def _abrir_corrigiendo_exif(ruta: str):
    if not _PIL_OK:
        return None
    try:
        img = Image.open(ruta)
        return ImageOps.exif_transpose(img)
    except Exception:
        return None

def _mejores_coincidencias(carpeta: str, patron: str, k: int = 4, threshold: float = 0.45) -> List[str]:
    files = _listar_imagenes(carpeta)
    scores = [(_score(os.path.basename(p), patron), p) for p in files]
    scores.sort(key=lambda x: x[0], reverse=True)
    return [p for sc, p in scores if sc >= threshold][:k]

def _seleccionar_por_objetivos(carpeta: str, objetivos: List[str], threshold: float = 0.45) -> List[Optional[str]]:
    files = _listar_imagenes(carpeta)
    usados = set()
    seleccion: List[Optional[str]] = []
    for obj in objetivos:
        if not obj:
            seleccion.append(None); continue
        candidatos = []
        for p in files:
            if p in usados: continue
            candidatos.append((_score(os.path.basename(p), obj), p))
        if not candidatos:
            seleccion.append(None); continue
        candidatos.sort(key=lambda x: x[0], reverse=True)
        best_sc, best_path = candidatos[0]
        if best_sc >= threshold:
            usados.add(best_path)
            seleccion.append(best_path)
        else:
            seleccion.append(None)
    return seleccion

# ======== Backend COM ========
def _with_excel(ruta_excel, worker):

    pythoncom.CoInitialize()

    xl = wb = None
    try:
        xl = win32.DispatchEx("Excel.Application")
        xl.DisplayAlerts  = False
        xl.ScreenUpdating = False
        xl.EnableEvents   = False

        prev_calc = None

        try:
            prev_calc = xl.Application.Calculation
            xl.Application.Calculation = -4135  
        except Exception:
            prev_calc = None

        wb = xl.Workbooks.Open(os.path.abspath(ruta_excel), ReadOnly=False, UpdateLinks=0)
        result = worker(xl, wb)
        wb.Save()

        if prev_calc is not None:
            try:
                xl.Application.Calculation = prev_calc
            except Exception:
                pass
        return result
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=True)
        except Exception:
            pass
        try:
            if xl is not None:
                xl.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()

def _px_to_pt(px: float) -> float:
    return float(px) * 0.75

def _insertar_img_en_celda_com(ws, celda, ruta_img, img_w, img_h, autorotar=True):

    from tempfile import NamedTemporaryFile
    src = os.path.abspath(ruta_img)
    tmpfile = None


    w_pt = _px_to_pt(img_w)
    h_pt = _px_to_pt(img_h)

    r = ws.Range(celda)
    try:

        _limpiar_imagenes_en_celdas_com(ws, [celda])
    except Exception:
        pass

    if autorotar and _PIL_OK:
        try:
            pil = _abrir_corrigiendo_exif(ruta_img)
            if pil is not None:
                tmp = NamedTemporaryFile(delete=False, suffix=".png")
                pil.save(tmp.name, format="PNG")
                tmpfile = tmp.name
                src = tmpfile
        except Exception:
            pass

    shp = ws.Shapes.AddPicture(
        Filename=src,
        LinkToFile=False,
        SaveWithDocument=True,
        Left=r.Left,
        Top=r.Top,
        Width=w_pt,
        Height=h_pt
    )

    try:
        shp.Name = f"FOTO__{ws.Name}__{celda}"
    except Exception:
        pass

    try:
        shp.LockAspectRatio = 0 
    except Exception:
        pass
    try:
        shp.Width = w_pt
        shp.Height = h_pt
    except Exception:
        try:
            if shp.Width and shp.Height:
                shp.ScaleWidth(w_pt / float(shp.Width), 0)   
                shp.ScaleHeight(h_pt / float(shp.Height), 0)
                shp.Width = w_pt
                shp.Height = h_pt
        except Exception:
            pass
    try:
        shp.Placement = 2 
    except Exception:
        pass

    if tmpfile:
        try:
            os.unlink(tmpfile)
        except Exception:
            pass


def _limpiar_imagenes_en_celdas_com(ws, celdas: List[str]):
    targets = []
    for c in celdas:
        try:
            r = ws.Range(c)
            targets.append((round(r.Left, 1), round(r.Top, 1)))
        except Exception:
            continue
    to_delete = []
    for shp in ws.Shapes:
        try:
            if (round(shp.Left, 1), round(shp.Top, 1)) in targets:
                to_delete.append(shp)
        except Exception:
            continue
    for shp in to_delete:
        try: shp.Delete()
        except Exception: pass

def procesar_fotos_por_patron(
    ruta_excel: str,
    carpeta_inicio: str,
    patron_inicio: str,
    carpeta_cierre: str,
    patron_cierre: str,
    destinos: Optional[Dict[str, List[str]]] = None,
    img_w: int = IMG_WIDTH,
    img_h: int = IMG_HEIGHT,
    limpiar_previas: bool = False,
    logger: Callable[[str], None] = print,
    autorotar: bool = True,
) -> Dict[str, Dict[str, int]]:

    if destinos is None:
        destinos = DESTINOS_DEF

    if not os.path.isfile(ruta_excel):
        raise FileNotFoundError("Excel destino no existe.")
    if not os.path.isdir(carpeta_inicio):
        raise FileNotFoundError("Carpeta de INICIO inválida.")
    if not os.path.isdir(carpeta_cierre):
        raise FileNotFoundError("Carpeta de CIERRE inválida.")

    def _worker(xl, wb):
        sheetnames = [s.Name for s in wb.Sheets]
        if "RT_1" not in sheetnames: raise ValueError("No existe hoja 'RT_1'.")
        if "RT_2" not in sheetnames: raise ValueError("No existe hoja 'RT_2'.")
        ws1, ws2 = wb.Worksheets("RT_1"), wb.Worksheets("RT_2")

        dest1 = destinos.get("RT_1", [])[:4]
        dest2 = destinos.get("RT_2", [])[:4]

        if limpiar_previas:
            _limpiar_imagenes_en_celdas_com(ws1, dest1)
            _limpiar_imagenes_en_celdas_com(ws2, dest2)
            logger("🧹 Imágenes previas removidas en celdas destino.")

        logger(f"🔎 INICIO: buscando patrón “{patron_inicio}”…")
        sel_ini = _mejores_coincidencias(carpeta_inicio, patron_inicio, k=len(dest1))
        ok1 = err1 = 0
        for ruta, celda in zip(sel_ini, dest1):
            try:
                if not ruta:
                    logger(f"⚠️ Sin coincidencia para RT_1:{celda}"); continue
                _insertar_img_en_celda_com(ws1, celda, ruta, img_w, img_h, autorotar=autorotar)
                ok1 += 1; logger(f"✅ {os.path.basename(ruta)} → RT_1:{celda}")
            except Exception as e:
                err1 += 1; logger(f"❌ {os.path.basename(ruta)} → RT_1:{celda}: {e}")

        logger(f"🔎 CIERRE: buscando patrón “{patron_cierre}”…")
        sel_cie = _mejores_coincidencias(carpeta_cierre, patron_cierre, k=len(dest2))
        ok2 = err2 = 0
        for ruta, celda in zip(sel_cie, dest2):
            try:
                if not ruta:
                    logger(f"⚠️ Sin coincidencia para RT_2:{celda}"); continue
                _insertar_img_en_celda_com(ws2, celda, ruta, img_w, img_h, autorotar=autorotar)
                ok2 += 1; logger(f"✅ {os.path.basename(ruta)} → RT_2:{celda}")
            except Exception as e:
                err2 += 1; logger(f"❌ {os.path.basename(ruta)} → RT_2:{celda}: {e}")

        return {"RT_1": {"ok": ok1, "err": err1}, "RT_2": {"ok": ok2, "err": err2}}

    return _with_excel(ruta_excel, _worker)

def procesar_fotos_por_objetivos(
    ruta_excel: str,
    carpeta_inicio: str,
    objetivos_inicio: List[str],
    carpeta_cierre: str,
    objetivos_cierre: List[str],
    destinos: Optional[Dict[str, List[str]]] = None,
    img_w: int = IMG_WIDTH,
    img_h: int = IMG_HEIGHT,
    limpiar_previas: bool = False,
    logger: Callable[[str], None] = print,
    autorotar: bool = True,
) -> Dict[str, Dict[str, int]]:

    if destinos is None:
        destinos = DESTINOS_DEF

    if not os.path.isfile(ruta_excel):
        raise FileNotFoundError("Excel destino no existe.")
    if not os.path.isdir(carpeta_inicio):
        raise FileNotFoundError("Carpeta de INICIO inválida.")
    if not os.path.isdir(carpeta_cierre):
        raise FileNotFoundError("Carpeta de CIERRE inválida.")

    def _worker(xl, wb):
        sheetnames = [s.Name for s in wb.Sheets]
        if "RT_1" not in sheetnames: raise ValueError("No existe hoja 'RT_1'.")
        if "RT_2" not in sheetnames: raise ValueError("No existe hoja 'RT_2'.")
        ws1, ws2 = wb.Worksheets("RT_1"), wb.Worksheets("RT_2")

        dest1 = (destinos.get("RT_1", []) or [])[:len(objetivos_inicio)]
        dest2 = (destinos.get("RT_2", []) or [])[:len(objetivos_cierre)]

        if limpiar_previas:
            _limpiar_imagenes_en_celdas_com(ws1, dest1)
            _limpiar_imagenes_en_celdas_com(ws2, dest2)
            logger("🧹 Imágenes previas removidas en celdas destino.")

        logger("🔎 INICIO: seleccionando por objetivos…")
        rutas_ini = _seleccionar_por_objetivos(carpeta_inicio, objetivos_inicio)
        ok1 = err1 = 0
        for ruta, celda in zip(rutas_ini, dest1):
            try:
                if not ruta:
                    logger(f"⚠️ Sin coincidencia para RT_1:{celda}"); continue
                _insertar_img_en_celda_com(ws1, celda, ruta, img_w, img_h, autorotar=autorotar)
                ok1 += 1; logger(f"✅ {os.path.basename(ruta)} → RT_1:{celda}")
            except Exception as e:
                err1 += 1; logger(f"❌ {os.path.basename(ruta)} → RT_1:{celda}: {e}")

        logger("🔎 CIERRE: seleccionando por objetivos…")
        rutas_cie = _seleccionar_por_objetivos(carpeta_cierre, objetivos_cierre)
        ok2 = err2 = 0
        for ruta, celda in zip(rutas_cie, dest2):
            try:
                if not ruta:
                    logger(f"⚠️ Sin coincidencia para RT_2:{celda}"); continue
                _insertar_img_en_celda_com(ws2, celda, ruta, img_w, img_h, autorotar=autorotar)
                ok2 += 1; logger(f"✅ {os.path.basename(ruta)} → RT_2:{celda}")
            except Exception as e:
                err2 += 1; logger(f"❌ {os.path.basename(ruta)} → RT_2:{celda}: {e}")

        return {"RT_1": {"ok": ok1, "err": err1}, "RT_2": {"ok": ok2, "err": err2}}

    return _with_excel(ruta_excel, _worker)

def procesar_fotos_predefinidos(
    ruta_excel: str,
    carpeta_inicio: str | None,
    carpeta_cierre: str | None,
    destinos: Optional[Dict[str, List[str]]] = None,
    img_w: int = IMG_WIDTH,
    img_h: int = IMG_HEIGHT,
    limpiar_previas: bool = False,
    logger: Callable[[str], None] = print,
    autorotar: bool = True,
    procesar_inicio: bool = True,
    procesar_cierre: bool = True,
) -> Dict[str, Dict[str, int]]:
    

    if destinos is None:
        destinos = DESTINOS_DEF

    if not os.path.isfile(ruta_excel):
        raise FileNotFoundError("Excel destino no existe.")

    if not procesar_inicio and not procesar_cierre:
        return {}

    if procesar_inicio:
        if not carpeta_inicio or not os.path.isdir(carpeta_inicio):
            raise FileNotFoundError("Carpeta de INICIO inválida.")
    if procesar_cierre:
        if not carpeta_cierre or not os.path.isdir(carpeta_cierre):
            raise FileNotFoundError("Carpeta de CIERRE inválida.")

    def _worker(xl, wb):
        sheetnames = [s.Name for s in wb.Sheets]

        if procesar_inicio and "RT_1" not in sheetnames:
            raise ValueError("No existe hoja 'RT_1'.")
        if procesar_cierre and "RT_2" not in sheetnames:
            raise ValueError("No existe hoja 'RT_2'.")

        out = {"RT_1": {"ok": 0, "err": 0}, "RT_2": {"ok": 0, "err": 0}}

        if procesar_inicio:
            ws1 = wb.Worksheets("RT_1")
            dest1 = (destinos.get("RT_1", []) or [])[:len(OBJETIVOS_DEF["RT_1"])]

            if limpiar_previas:
                _limpiar_imagenes_en_celdas_com(ws1, dest1)
                logger("🧹 (RT_1) Imágenes previas removidas en celdas destino.")

            logger("🔎 RT_1: seleccionando por objetivos…")
            rutas_ini = _seleccionar_por_objetivos(carpeta_inicio, OBJETIVOS_DEF["RT_1"])
            ok1 = err1 = 0

            for ruta, celda in zip(rutas_ini, dest1):
                try:
                    if not ruta:
                        logger(f"⚠️ Sin coincidencia para RT_1:{celda}")
                        continue
                    _insertar_img_en_celda_com(ws1, celda, ruta, img_w, img_h, autorotar=autorotar)
                    ok1 += 1
                    logger(f"✅ {os.path.basename(ruta)} → RT_1:{celda}")
                except Exception as e:
                    err1 += 1
                    logger(f"❌ {os.path.basename(ruta) if ruta else 'N/A'} → RT_1:{celda}: {e}")

            out["RT_1"] = {"ok": ok1, "err": err1}


        if procesar_cierre:
            ws2 = wb.Worksheets("RT_2")
            dest2 = (destinos.get("RT_2", []) or [])[:len(OBJETIVOS_DEF["RT_2"])]

            if limpiar_previas:
                _limpiar_imagenes_en_celdas_com(ws2, dest2)
                logger("🧹 (RT_2) Imágenes previas removidas en celdas destino.")

            logger("🔎 RT_2: seleccionando por objetivos…")
            rutas_cie = _seleccionar_por_objetivos(carpeta_cierre, OBJETIVOS_DEF["RT_2"])
            ok2 = err2 = 0

            for ruta, celda in zip(rutas_cie, dest2):
                try:
                    if not ruta:
                        logger(f"⚠️ Sin coincidencia para RT_2:{celda}")
                        continue
                    _insertar_img_en_celda_com(ws2, celda, ruta, img_w, img_h, autorotar=autorotar)
                    ok2 += 1
                    logger(f"✅ {os.path.basename(ruta)} → RT_2:{celda}")
                except Exception as e:
                    err2 += 1
                    logger(f"❌ {os.path.basename(ruta) if ruta else 'N/A'} → RT_2:{celda}: {e}")

            out["RT_2"] = {"ok": ok2, "err": err2}

        return out

    return _with_excel(ruta_excel, _worker)
