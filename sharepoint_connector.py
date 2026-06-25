"""
sharepoint_connector.py  v5
Estrategia de carga con fallback automático:
  1. Intenta descargar desde SharePoint (links en config.toml)
  2. Si falla por cualquier motivo → usa archivos locales en datos/
  3. Si tampoco hay locales → muestra error claro

Así el dashboard siempre funciona, con o sin conexión a SharePoint.
"""

import requests, io, toml, os
import pandas as pd
import streamlit as st
from datetime import datetime

MESES_MAP = {
    'enero':1,'febrero':2,'marzo':3,'abril':4,'mayo':5,'junio':6,
    'julio':7,'agosto':8,'septiembre':9,'octubre':10,'noviembre':11,'diciembre':12
}
MESES_12 = {1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
             7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",
             11:"Noviembre",12:"Diciembre"}

# Nombres de los archivos locales en la carpeta datos/
ARCHIVOS_LOCALES = {
    "facturas": "PROTOTIPO_Homologacion_CRAFT_v6_con_datos.xlsx",
    "ppto_adm": "4__ADM_PPTO_gastos_año_2026_BR.xlsx",
    "ppto_gh":  "5__GH_PPTO_gastos_año_2026_BR.xlsx",
    "ppto_it":  "6__IT_PPTO_gastos_año_2026_BR.xlsx",
}

# Valores de respaldo si ningún archivo de presupuesto se puede leer
PPTO_FALLBACK = {
    "ADM": 5_020_600_000.0,
    "IT":  473_000_000.0,
    "GH":  314_225_957.73,
}

# ── CONFIG ────────────────────────────────────────────────────────
def cargar_config():
    try:
        return toml.load("config.toml")
    except FileNotFoundError:
        st.error("❌ No se encontró config.toml.")
        st.stop()
    except Exception as e:
        st.error(f"Error leyendo config.toml: {e}")
        st.stop()

def validar_config(config):
    faltantes = [k for k, v in config.get("archivos", {}).items()
                 if str(v).startswith("PEGA_AQUI")]
    if faltantes:
        st.warning(f"⚙️ Links no configurados en config.toml: **{', '.join(faltantes)}**. "
                   "Se usarán los archivos locales.")
        return False
    return True

# ── DESCARGA SHAREPOINT ───────────────────────────────────────────
def _link_descarga(link):
    if not link or str(link).startswith("PEGA_AQUI"):
        return None
    sep = "&" if "?" in link else "?"
    return f"{link}{sep}download=1"

def _descargar_sharepoint(link, nombre):
    """
    Intenta descargar desde SharePoint.
    Retorna (bytes, None) si OK, o (None, mensaje_error) si falla.
    """
    url = _link_descarga(link)
    if not url:
        return None, "link no configurado"
    try:
        r = requests.get(url,
                         headers={"User-Agent": "CraftDashboard/5.0"},
                         timeout=20,
                         allow_redirects=True)
        r.raise_for_status()
        if r.content[:4] == b'PK\x03\x04':
            return r.content, None
        return None, "la respuesta no es un archivo Excel válido"
    except requests.exceptions.Timeout:
        return None, "tiempo de espera agotado"
    except requests.exceptions.HTTPError as e:
        return None, f"HTTP {e.response.status_code}"
    except Exception as e:
        return None, str(e)

# ── CARGA LOCAL ───────────────────────────────────────────────────
def _cargar_local(key, nombre):
    """
    Lee el archivo local como bytes.
    Retorna (bytes, None) si OK, o (None, mensaje_error) si falla.
    """
    ruta = ARCHIVOS_LOCALES.get(key)
    if not ruta or not os.path.exists(ruta):
        return None, f"archivo no encontrado en {ruta}"
    try:
        with open(ruta, "rb") as f:
            return f.read(), None
    except Exception as e:
        return None, str(e)

# ── OBTENER BYTES (SharePoint → local → error) ───────────────────
def _obtener_bytes(key, nombre, link):
    """
    Intenta SharePoint primero. Si falla, usa local.
    Retorna (bytes, fuente) donde fuente es 'sharepoint' o 'local'.
    """
    # 1. Intentar SharePoint
    if link and not str(link).startswith("PEGA_AQUI"):
        b, err = _descargar_sharepoint(link, nombre)
        if b:
            return b, "sharepoint"
        # SharePoint falló → avisar silenciosamente y usar local
        st.toast(f"📁 {nombre}: SharePoint no disponible, usando archivo local.",
                 icon="ℹ️")

    # 2. Usar archivo local
    b, err = _cargar_local(key, nombre)
    if b:
        return b, "local"

    # 3. Nada funcionó
    st.error(f"❌ No se pudo cargar **{nombre}**.\n"
             f"- SharePoint: sin acceso\n"
             f"- Local: {err}\n\n"
             "Verifica que el archivo esté en la carpeta `datos/`.")
    return None, None

# ── PARSEO ────────────────────────────────────────────────────────
def _parsear_mes(texto):
    if pd.isna(texto):
        return 0, 0, "Sin mes"
    partes = str(texto).lower().strip().split()
    num  = MESES_MAP.get(partes[0], 0)
    anio = int(partes[1]) if len(partes) > 1 and partes[1].isdigit() else 0
    nombre = f"{partes[0].capitalize()} {anio}" if anio else partes[0].capitalize()
    return num, anio, nombre

# ── LECTURA DE GASTOS ─────────────────────────────────────────────
def _limpiar_registro(b):
    raw = pd.read_excel(io.BytesIO(b), sheet_name='REGISTRO', header=1)
    raw.columns = raw.iloc[0].tolist()
    data = raw.iloc[1:].copy().reset_index(drop=True)
    data['Monto']      = pd.to_numeric(data.get('MONTO FACTURA', 0), errors='coerce')
    data['Gasto_Neto'] = pd.to_numeric(data.get('GASTO NETO', data['Monto']), errors='coerce')
    data = data[data['Monto'] > 0].copy()
    data['Fuente']    = 'REGISTRO'
    data['Proveedor'] = data['PROVEEDOR'].fillna('Sin proveedor')
    data['Area']      = data['ÁREA'].fillna('Sin clasificar')
    data['Sucursal']  = data['SUCURSAL'].fillna('Sin clasificar')
    data['Cuenta']    = data['CUENTA CONTABLE'].fillna('Sin clasificar')
    data['Tipo']      = data['TIPO'].fillna('Sin clasificar')
    data['Servicio']  = data['SERVICIO'].fillna('Sin clasificar')
    data['Detalle']   = data['DETALLE DEL GASTO'].fillna('')
    parsed = data['MES EJECUCIÓN'].apply(_parsear_mes)
    data['Mes_num'] = parsed.apply(lambda x: x[0])
    data['Anio']    = parsed.apply(lambda x: x[1])
    data['Mes']     = parsed.apply(lambda x: x[2])
    return data[['Fuente','Proveedor','Area','Sucursal','Cuenta','Tipo',
                 'Servicio','Detalle','Monto','Gasto_Neto','Mes_num','Anio','Mes']]

def _limpiar_no_pagar(b):
    raw = pd.read_excel(io.BytesIO(b), sheet_name='REGISTRO NO PAGAR', header=1)
    raw.columns = raw.iloc[0].tolist()
    data = raw.iloc[1:].copy().reset_index(drop=True)
    data['Monto']      = pd.to_numeric(data.get('MONTO', 0), errors='coerce')
    data['Gasto_Neto'] = data['Monto'].copy()
    data = data[data['Monto'] > 0].copy()
    data['Fuente']    = 'NO PAGAR'
    data['Proveedor'] = data['PROVEEDOR'].fillna('Sin proveedor')
    data['Area']      = data['ÁREA'].fillna('Sin clasificar')
    data['Sucursal']  = data['SUCURSAL'].fillna('Sin clasificar')
    data['Cuenta']    = data['CUENTA CONTABLE'].fillna('Sin clasificar')
    data['Tipo']      = data['TIPO'].fillna('Sin clasificar')
    data['Servicio']  = data['SERVICIO'].fillna('Sin clasificar')
    data['Detalle']   = data['DETALLE DEL GASTO'].fillna('')
    parsed = data['MES EJECUCIÓN'].apply(_parsear_mes)
    data['Mes_num'] = parsed.apply(lambda x: x[0])
    data['Anio']    = parsed.apply(lambda x: x[1])
    data['Mes']     = parsed.apply(lambda x: x[2])
    return data[['Fuente','Proveedor','Area','Sucursal','Cuenta','Tipo',
                 'Servicio','Detalle','Monto','Gasto_Neto','Mes_num','Anio','Mes']]

# ── LECTURA DE PRESUPUESTOS ───────────────────────────────────────
def _ppto_adm(b):
    rows, err = [], None
    try:
        xl   = pd.ExcelFile(io.BytesIO(b))
        hoja = next((s for s in xl.sheet_names
                     if 'mensual' in s.lower() and 'adm' in s.lower()), None)
        if not hoja:
            hoja = next((s for s in xl.sheet_names
                         if 'presupuesto' in s.lower()), xl.sheet_names[0])
        adm = pd.read_excel(io.BytesIO(b), sheet_name=hoja, header=1)
        adm.columns = ["skip","Item","Ppto_Mes","USD_Mes"]
        adm["Ppto_Mes"] = pd.to_numeric(adm["Ppto_Mes"], errors="coerce")
        adm = adm[adm["Ppto_Mes"] > 0].dropna(subset=["Ppto_Mes"])
        for _, r in adm.iterrows():
            for mn, mn_n in MESES_12.items():
                rows.append({"Area":"ADM","Item":r["Item"],
                             "Mes_num":mn,"Mes":mn_n,"Ppto_Mes":r["Ppto_Mes"]})
    except Exception as e:
        err = str(e)
    return rows, err

def _ppto_it(b):
    rows, err = [], None
    try:
        xl   = pd.ExcelFile(io.BytesIO(b))
        hoja = next((s for s in xl.sheet_names if 'it' in s.lower()), xl.sheet_names[0])
        it   = pd.read_excel(io.BytesIO(b), sheet_name=hoja, header=None)
        for _, r in it.iloc[3:, :].iterrows():
            item = str(r.iloc[0]).strip()
            if not item or item in ["nan","None"] or "TOTAL" in item.upper():
                continue
            for col_idx, (mn, mn_n) in enumerate(MESES_12.items(), start=1):
                try:
                    val = pd.to_numeric(r.iloc[col_idx], errors="coerce")
                    if pd.notna(val) and val > 0:
                        rows.append({"Area":"IT","Item":item,
                                     "Mes_num":mn,"Mes":mn_n,"Ppto_Mes":val})
                except IndexError:
                    pass
    except Exception as e:
        err = str(e)
    return rows, err

def _ppto_gh(b):
    rows, err = [], None
    try:
        xl   = pd.ExcelFile(io.BytesIO(b))
        hoja = next((s for s in xl.sheet_names
                     if 'hr' in s.lower() or 'gh' in s.lower()), xl.sheet_names[0])
        gh   = pd.read_excel(io.BytesIO(b), sheet_name=hoja, header=0)
        n    = len(gh.columns)
        base = ["skip","Proceso","Actividad","Cantidad","Costo_ind",
                "Total_Año","Total_mes","Frecuencia","Frecuencia2","IPC"]
        gh.columns = base[:n] + [f"c{i}" for i in range(max(0, n-10))]
        gh["Total_mes"] = pd.to_numeric(gh["Total_mes"], errors="coerce")
        gh_clean = gh[gh["Total_mes"] > 0].dropna(subset=["Total_mes"])
        for _, r in gh_clean.iterrows():
            for mn, mn_n in MESES_12.items():
                rows.append({"Area":"GH","Item":str(r["Actividad"]),
                             "Mes_num":mn,"Mes":mn_n,"Ppto_Mes":r["Total_mes"]})
    except Exception as e:
        err = str(e)
    return rows, err

def _construir_ppto(rows_adm, rows_it, rows_gh,
                    err_adm, err_it, err_gh):
    avisos = []
    if not rows_adm:
        val = PPTO_FALLBACK["ADM"] / 12
        for mn, mn_n in MESES_12.items():
            rows_adm.append({"Area":"ADM","Item":"Presupuesto ADM",
                             "Mes_num":mn,"Mes":mn_n,"Ppto_Mes":val})
        avisos.append("ADM")
    if not rows_it:
        val = PPTO_FALLBACK["IT"] / 12
        for mn, mn_n in MESES_12.items():
            rows_it.append({"Area":"IT","Item":"Presupuesto IT",
                            "Mes_num":mn,"Mes":mn_n,"Ppto_Mes":val})
        avisos.append("IT")
    if not rows_gh:
        val = PPTO_FALLBACK["GH"] / 12
        for mn, mn_n in MESES_12.items():
            rows_gh.append({"Area":"GH","Item":"Presupuesto GH",
                            "Mes_num":mn,"Mes":mn_n,"Ppto_Mes":val})
        avisos.append("GH")
    if avisos:
        st.info(f"ℹ️ Presupuesto de **{', '.join(avisos)}** usando valores de referencia "
                "(no se pudo leer el archivo). Los totales son aproximados.")
    df = pd.concat([pd.DataFrame(rows_adm),
                    pd.DataFrame(rows_it),
                    pd.DataFrame(rows_gh)], ignore_index=True)
    ppto_mensual = (df.groupby(["Area","Mes_num","Mes"])["Ppto_Mes"]
                    .sum().reset_index().sort_values("Mes_num"))
    ppto_area    = (df.groupby("Area")["Ppto_Mes"]
                    .sum().reset_index()
                    .rename(columns={"Ppto_Mes":"Ppto_Anual"}))
    return ppto_mensual, ppto_area

# ── FUNCIÓN PRINCIPAL ─────────────────────────────────────────────
@st.cache_data(ttl=1800, show_spinner=False)
def obtener_datos_sharepoint(facturas_link, ppto_adm_link,
                              ppto_gh_link,  ppto_it_link):
    out     = {}
    fuentes = {}   # registra de dónde vino cada archivo

    # ── Facturas ──────────────────────────────────────────────────
    with st.spinner("📥 Cargando facturas..."):
        b_fact, f_fact = _obtener_bytes("facturas", "Facturas", facturas_link)
        fuentes["facturas"] = f_fact

    if b_fact:
        try:
            df_reg = _limpiar_registro(b_fact)
        except Exception as e:
            st.error(f"Error leyendo hoja REGISTRO: {e}")
            df_reg = pd.DataFrame()
        try:
            df_np = _limpiar_no_pagar(b_fact)
        except Exception as e:
            st.error(f"Error leyendo hoja REGISTRO NO PAGAR: {e}")
            df_np = pd.DataFrame()

        frames = [x for x in [df_reg, df_np] if not x.empty]
        df = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
        out["df"]         = df
        out["n_sin_area"] = int((df["Area"] == "Sin clasificar").sum()) if not df.empty else 0
        out["n_sin_suc"]  = int((df["Sucursal"] == "Sin clasificar").sum()) if not df.empty else 0
    else:
        out["df"]         = pd.DataFrame()
        out["n_sin_area"] = 0
        out["n_sin_suc"]  = 0

    # ── Presupuestos ──────────────────────────────────────────────
    with st.spinner("📥 Cargando presupuesto ADM..."):
        b_adm, f_adm = _obtener_bytes("ppto_adm", "Presupuesto ADM", ppto_adm_link)
        fuentes["ppto_adm"] = f_adm
    with st.spinner("📥 Cargando presupuesto IT..."):
        b_it,  f_it  = _obtener_bytes("ppto_it",  "Presupuesto IT",  ppto_it_link)
        fuentes["ppto_it"] = f_it
    with st.spinner("📥 Cargando presupuesto GH..."):
        b_gh,  f_gh  = _obtener_bytes("ppto_gh",  "Presupuesto GH",  ppto_gh_link)
        fuentes["ppto_gh"] = f_gh

    rows_adm, err_adm = _ppto_adm(b_adm) if b_adm else ([], "no disponible")
    rows_it,  err_it  = _ppto_it(b_it)   if b_it  else ([], "no disponible")
    rows_gh,  err_gh  = _ppto_gh(b_gh)   if b_gh  else ([], "no disponible")

    ppto_mensual, ppto_area = _construir_ppto(
        rows_adm, rows_it, rows_gh, err_adm, err_it, err_gh
    )
    out["ppto_mensual"] = ppto_mensual
    out["ppto_area"]    = ppto_area
    out["ultima_act"]   = datetime.now().strftime("%d/%m/%Y %H:%M")
    out["fuentes"]      = fuentes
    return out
