"""
Dashboard Presupuesto – Craft Logistics  v7
· Facturas: SharePoint (se actualiza automáticamente)
· Presupuesto: valores fijos en código (no cambia frecuentemente)
· Ejecutar: streamlit run app.py
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests, io, toml

# ── PÁGINA ────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Dashboard Presupuesto · Craft",
    page_icon="📊", layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
  .stApp { background-color: #F8FAFC; }
  #MainMenu, footer { visibility: hidden; }
  .kpi-box {
    background: #FFFFFF; border-radius: 12px; padding: 18px 20px;
    box-shadow: 0 1px 6px rgba(0,0,0,.07); margin-bottom: 4px;
  }
  .kpi-label { font-size: 11px; font-weight: 600; color: #64748B;
               text-transform: uppercase; letter-spacing: .06em; }
  .kpi-value { font-size: 26px; font-weight: 700; margin: 6px 0 2px; line-height: 1; }
  .kpi-sub   { font-size: 11px; color: #94A3B8; }
  .sec { font-size: 11px; font-weight: 700; color: #64748B;
         text-transform: uppercase; letter-spacing: .1em;
         margin: 24px 0 12px; border-bottom: 1px solid #E2E8F0; padding-bottom: 6px; }
</style>
""", unsafe_allow_html=True)

PAL      = ["#1A56DB","#7C3AED","#059669","#D97706","#DC2626",
            "#0891B2","#DB2777","#65A30D","#EA580C","#0369A1"]
COL_AREA = {"ADM":"#1A56DB","IT":"#7C3AED","GH":"#059669","Sin clasificar":"#94A3B8"}
MESES_ORD= ["Enero","Febrero","Marzo","Abril","Mayo","Junio",
            "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]
MESES_MAP= {'enero':1,'febrero':2,'marzo':3,'abril':4,'mayo':5,'junio':6,
            'julio':7,'agosto':8,'septiembre':9,'octubre':10,'noviembre':11,'diciembre':12}
MESES_12 = {v:k.capitalize() for k,v in MESES_MAP.items()}

# ── PRESUPUESTO ANUAL (fijo — actualizar aquí si cambia) ──────────
PPTO_ANUAL = {
    "ADM": 5_020_600_000.0,   # Presupuesto mensual ADM × 12
    "IT":  473_000_000.0,     # Total anual IT
    "GH":  314_225_957.73,    # Total mes GH × 12
}

# ── FORMATO COP ───────────────────────────────────────────────────
def cop(v):
    """$1.681.485  (puntos como miles, sin decimales)"""
    if pd.isna(v) or v == 0: return "$0"
    return "${:,.0f}".format(float(v)).replace(",", ".")

# ── CARGA DESDE SHAREPOINT ────────────────────────────────────────
def cargar_config():
    try:
        return toml.load("config.toml")
    except:
        return {"archivos": {"facturas": ""}}

def descargar_sharepoint(link, nombre):
    """
    Descarga un Excel desde SharePoint.
    Retorna bytes si OK, None si falla (con mensaje de error en UI).
    """
    if not link or link.startswith("PEGA_AQUI") or link == "":
        return None, "link no configurado"
    try:
        sep = "&" if "?" in link else "?"
        url = f"{link}{sep}download=1"
        r   = requests.get(url,
                           headers={"User-Agent": "Mozilla/5.0 CraftDashboard/7.0"},
                           timeout=25, allow_redirects=True)
        # Verificar que no sea una página de login de SharePoint
        ct = r.headers.get("Content-Type", "")
        if "text/html" in ct:
            return None, "SharePoint devolvió página de login (sin acceso)"
        r.raise_for_status()
        # Verificar magic bytes de ZIP/xlsx
        if r.content[:2] != b'PK':
            return None, "la respuesta no es un archivo Excel"
        return r.content, None
    except requests.exceptions.Timeout:
        return None, "tiempo de espera agotado (25s)"
    except requests.exceptions.HTTPError as e:
        return None, f"HTTP {e.response.status_code}"
    except Exception as e:
        return None, str(e)

def parse_mes(txt):
    if pd.isna(txt): return 0, 0, "Sin mes", "Sin mes"
    p    = str(txt).lower().strip().split()
    num  = MESES_MAP.get(p[0], 0)
    anio = int(p[1]) if len(p) > 1 and p[1].isdigit() else 0
    nom  = f"{p[0].capitalize()} {anio}" if anio else p[0].capitalize()
    base = p[0].capitalize()
    return num, anio, nom, base

def procesar_facturas(b):
    """Procesa bytes del Excel homologado → DataFrame limpio."""
    # REGISTRO
    raw = pd.read_excel(io.BytesIO(b), sheet_name='REGISTRO', header=1)
    raw.columns = raw.iloc[0].tolist()
    reg = raw.iloc[1:].copy().reset_index(drop=True)
    reg['Monto']      = pd.to_numeric(reg['MONTO FACTURA'], errors='coerce')
    reg['Gasto_Neto'] = pd.to_numeric(reg['GASTO NETO'],    errors='coerce')
    reg = reg[reg['Monto'] > 0].copy()
    reg['Fuente'] = 'REGISTRO'

    # NO PAGAR
    raw2 = pd.read_excel(io.BytesIO(b), sheet_name='REGISTRO NO PAGAR', header=1)
    raw2.columns = raw2.iloc[0].tolist()
    nop = raw2.iloc[1:].copy().reset_index(drop=True)
    nop['Monto']      = pd.to_numeric(nop['MONTO'], errors='coerce')
    nop['Gasto_Neto'] = nop['Monto']
    nop = nop[nop['Monto'] > 0].copy()
    nop['Fuente'] = 'NO PAGAR'

    cols_comunes = ['Fuente','PROVEEDOR','ÁREA','SUCURSAL',
                    'CUENTA CONTABLE','TIPO','Monto','Gasto_Neto','MES EJECUCIÓN']
    df = pd.concat([reg[cols_comunes], nop[cols_comunes]], ignore_index=True)
    df = df.rename(columns={
        'PROVEEDOR':'Proveedor','ÁREA':'Area','SUCURSAL':'Sucursal',
        'CUENTA CONTABLE':'Cuenta','TIPO':'Tipo','MES EJECUCIÓN':'Mes_raw'
    })
    df['Proveedor'] = df['Proveedor'].fillna('Sin proveedor')
    df['Area']      = df['Area'].fillna('Sin clasificar')
    df['Sucursal']  = df['Sucursal'].fillna('Sin clasificar')
    df['Cuenta']    = df['Cuenta'].fillna('Sin clasificar')
    df['Tipo']      = df['Tipo'].fillna('Sin clasificar')

    parsed          = df['Mes_raw'].apply(parse_mes)
    df['Mes_num']   = parsed.apply(lambda x: x[0])
    df['Anio']      = parsed.apply(lambda x: x[1])
    df['Mes_nom']   = parsed.apply(lambda x: x[2])   # "Enero 2025"
    df['Mes_base']  = parsed.apply(lambda x: x[3])   # "Enero"

    return df

# ── CARGA PRINCIPAL (cache 30 min) ────────────────────────────────
@st.cache_data(ttl=1800, show_spinner=False)
def cargar_datos(link_facturas):
    b, err = descargar_sharepoint(link_facturas, "Facturas")
    if b:
        df     = procesar_facturas(b)
        fuente = "☁️ SharePoint"
    else:
        st.error(f"❌ No se pudo cargar el archivo de facturas desde SharePoint: **{err}**\n\n"
                 "Verifica que el link en `config.toml` sea correcto y sea de tipo "
                 "'Cualquier persona con el vínculo'.")
        st.stop()
    return df, fuente

# ── LEER CONFIG Y DATOS ───────────────────────────────────────────
config         = cargar_config()
link_facturas  = config.get("archivos", {}).get("facturas", "")

if not link_facturas or link_facturas.startswith("PEGA_AQUI"):
    st.error("⚙️ Configura el link de facturas en `config.toml` antes de continuar.")
    st.stop()

with st.spinner("📥 Cargando facturas desde SharePoint..."):
    df, fuente_datos = cargar_datos(link_facturas)

# ── PRESUPUESTO ───────────────────────────────────────────────────
ppto_area = pd.DataFrame([
    {"Area": k, "Ppto_Anual": v} for k, v in PPTO_ANUAL.items()
])

# ── SIDEBAR ───────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🎛️ Filtros")

    anios_disp = sorted([a for a in df['Anio'].unique() if a > 0])
    sel_anio   = st.multiselect("📆 Año", anios_disp, default=anios_disp)

    meses_disp = [m for m in MESES_ORD if m in df['Mes_base'].values]
    sel_mes    = st.multiselect("📅 Mes", ["Todos"] + meses_disp, default=["Todos"])

    areas_disp = sorted(df['Area'].unique())
    sel_area   = st.multiselect("🏢 Área", areas_disp, default=areas_disp)

    sucs_disp  = sorted(df['Sucursal'].unique())
    sel_suc    = st.multiselect("📍 Sucursal", sucs_disp, default=sucs_disp)

    sel_fuente = st.multiselect("📂 Fuente",
                                 ["REGISTRO","NO PAGAR"],
                                 default=["REGISTRO","NO PAGAR"])

    st.divider()
    st.markdown("### 📋 Vista Presupuesto")
    vista_ppto = st.radio("",
        ["Anual", "Mensual (meses seleccionados)"], index=0)

    st.divider()
    st.markdown("### 🔄 Datos")
    st.markdown(f"**{fuente_datos}**")
    if st.button("🔄 Forzar actualización"):
        cargar_datos.clear()
        st.rerun()
    st.caption("Se actualiza automáticamente cada 30 min")

# ── FILTRADO ──────────────────────────────────────────────────────
f = df.copy()
if sel_anio:
    f = f[f['Anio'].isin(sel_anio)]
if "Todos" not in sel_mes and sel_mes:
    f = f[f['Mes_base'].isin(sel_mes)]
if sel_area:
    f = f[f['Area'].isin(sel_area)]
if sel_suc:
    f = f[f['Sucursal'].isin(sel_suc)]
if sel_fuente:
    f = f[f['Fuente'].isin(sel_fuente)]

meses_activos = [m for m in sel_mes if m != "Todos"]

# ── PRESUPUESTO SEGÚN VISTA ───────────────────────────────────────
ppto_f = ppto_area[ppto_area['Area'].isin(sel_area)].copy() if sel_area else ppto_area.copy()

if vista_ppto == "Mensual (meses seleccionados)" and meses_activos:
    n_m = len(meses_activos)
    ppto_f = ppto_f.copy()
    ppto_f['Ppto_Anual'] = ppto_f['Ppto_Anual'] / 12 * n_m
    label_p = f"Presupuesto ({n_m} mes{'es' if n_m>1 else ''})"
else:
    label_p = "Presupuesto Anual"

total_ppto  = ppto_f['Ppto_Anual'].sum()
total_gasto = f['Monto'].sum()
saldo       = total_ppto - total_gasto
pct_ejec    = (total_gasto / total_ppto * 100) if total_ppto > 0 else 0
n_fact      = len(f)
mes_v       = f[f['Mes_num'] > 0].groupby('Mes_num')['Monto'].sum()
prom_mes    = mes_v.mean() if len(mes_v) > 0 else 0

# ── HEADER ────────────────────────────────────────────────────────
st.markdown("""
<div style="background:linear-gradient(135deg,#0F1C3F,#1A3A6B);
     padding:20px 28px;border-radius:14px;margin-bottom:22px">
  <h1 style="color:#FFF;font-size:22px;margin:0;font-weight:700">
    📊 Dashboard Presupuesto · Craft Logistics
  </h1>
  <p style="color:#94AECF;font-size:13px;margin:4px 0 0">
    REGISTRO + NO PAGAR · ADM · IT · GH · Valores en COP
  </p>
</div>
""", unsafe_allow_html=True)

# ── KPIs ──────────────────────────────────────────────────────────
st.markdown('<div class="sec">Indicadores Clave</div>', unsafe_allow_html=True)
k1,k2,k3,k4,k5 = st.columns(5)

ce = "#059669" if pct_ejec < 80 else "#D97706" if pct_ejec < 95 else "#DC2626"
cs = "#059669" if saldo >= 0 else "#DC2626"

for col, lbl, val, sub, clr in [
    (k1, "Total Gastado",    cop(total_gasto),    "COP acumulado",   "#1A56DB"),
    (k2, label_p,            cop(total_ppto),      "ADM + IT + GH",   "#7C3AED"),
    (k3, "% Ejecución",      f"{pct_ejec:.1f}%",  "del presupuesto", ce),
    (k4, "Saldo Disponible", cop(abs(saldo)),      "disponible ✅" if saldo>=0 else "⚠️ excedido", cs),
    (k5, "Registros",        str(n_fact),          f"Prom {cop(prom_mes)}/mes", "#0891B2"),
]:
    with col:
        st.markdown(f"""<div class="kpi-box" style="border-left:4px solid {clr}">
        <div class="kpi-label">{lbl}</div>
        <div class="kpi-value" style="color:{clr}">{val}</div>
        <div class="kpi-sub">{sub}</div></div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── GASTO MENSUAL + PPTO vs REAL ─────────────────────────────────
st.markdown('<div class="sec">Evolución Mensual & Comparativo por Área</div>', unsafe_allow_html=True)
g1, g2 = st.columns([3, 2])

with g1:
    md = (f[f['Mes_num'] > 0]
          .groupby(['Mes_num','Mes_base'])['Monto'].sum()
          .reset_index().sort_values('Mes_num'))

    fig_mes = go.Figure()
    fig_mes.add_trace(go.Bar(
        x=md['Mes_base'], y=md['Monto'],
        name="Gasto mensual", marker_color="#1A56DB",
        text=[cop(v) for v in md['Monto']],
        textposition="outside", textfont=dict(size=9)
    ))
    fig_mes.add_trace(go.Scatter(
        x=md['Mes_base'], y=md['Monto'],
        name="Tendencia", mode="lines+markers",
        line=dict(color="#D97706", width=2),
        marker=dict(size=6, color="#D97706")
    ))
    fig_mes.update_layout(
        title="Gasto Mensual (COP)",
        plot_bgcolor="white", paper_bgcolor="white",
        yaxis=dict(tickformat="$,.0f"),
        legend=dict(orientation="h", y=-0.2),
        margin=dict(t=45,b=10,l=10,r=10), height=340
    )
    st.plotly_chart(fig_mes, use_container_width=True)

with g2:
    ag   = f.groupby('Area')['Monto'].sum().reset_index()
    ag.columns = ['Area','Gastado']
    comp = ppto_f.merge(ag, on='Area', how='left').fillna(0)
    comp['Pct'] = (comp['Gastado'] / comp['Ppto_Anual'] * 100).round(1)

    fig_comp = go.Figure()
    fig_comp.add_trace(go.Bar(
        name=label_p, x=comp['Area'], y=comp['Ppto_Anual'],
        marker_color="rgba(26,86,219,0.15)",
        marker_line=dict(color="#1A56DB", width=1.5)
    ))
    fig_comp.add_trace(go.Bar(
        name="Gastado", x=comp['Area'], y=comp['Gastado'],
        marker_color=[COL_AREA.get(a,"#94A3B8") for a in comp['Area']],
        text=[f"{p}%" for p in comp['Pct']], textposition="outside"
    ))
    fig_comp.update_layout(
        title=f"{label_p} vs Ejecutado",
        barmode="group", plot_bgcolor="white", paper_bgcolor="white",
        yaxis=dict(tickformat="$,.0f"),
        legend=dict(orientation="h", y=-0.2),
        margin=dict(t=45,b=10,l=10,r=10), height=340
    )
    st.plotly_chart(fig_comp, use_container_width=True)

# ── SUCURSALES + TIPO DE GASTO ────────────────────────────────────
st.markdown('<div class="sec">Por Sucursal & Tipo de Gasto</div>', unsafe_allow_html=True)
g3, g4 = st.columns([3, 2])

with g3:
    sd = (f.groupby('Sucursal')['Monto'].sum()
          .reset_index().sort_values('Monto', ascending=True))
    sd['Pct'] = (sd['Monto'] / sd['Monto'].sum() * 100).round(1)

    fig_suc = go.Figure(go.Bar(
        x=sd['Monto'], y=sd['Sucursal'], orientation="h",
        marker_color=[PAL[i%len(PAL)] for i in range(len(sd))],
        text=[f"{cop(v)} ({p}%)" for v,p in zip(sd['Monto'], sd['Pct'])],
        textposition="outside"
    ))
    fig_suc.update_layout(
        title="Gasto por Sucursal (COP)",
        plot_bgcolor="white", paper_bgcolor="white",
        xaxis=dict(tickformat="$,.0f"),
        margin=dict(t=45,b=10,l=10,r=10), height=360
    )
    st.plotly_chart(fig_suc, use_container_width=True)

with g4:
    cd = (f[f['Cuenta'] != 'Sin clasificar']
          .groupby('Cuenta')['Monto'].sum()
          .reset_index().sort_values('Monto', ascending=False).head(8))

    fig_pie = px.pie(
        cd, values='Monto', names='Cuenta',
        title="Distribución por Tipo de Gasto",
        color_discrete_sequence=PAL, hole=0.45
    )
    fig_pie.update_traces(
        textposition="inside", textinfo="percent+label", textfont_size=10
    )
    fig_pie.update_layout(
        showlegend=False, plot_bgcolor="white", paper_bgcolor="white",
        margin=dict(t=45,b=10,l=10,r=10), height=360
    )
    st.plotly_chart(fig_pie, use_container_width=True)

# ── TOP PROVEEDORES + ÁREA × MES ──────────────────────────────────
st.markdown('<div class="sec">Top Proveedores & Gasto por Área y Mes</div>', unsafe_allow_html=True)
g5, g6 = st.columns([2, 3])

with g5:
    pd_ = (f[f['Proveedor'] != 'Sin proveedor']
           .groupby('Proveedor')['Monto'].sum()
           .reset_index().sort_values('Monto', ascending=False).head(10))
    pd_['Pct'] = (pd_['Monto'] / f['Monto'].sum() * 100).round(1)

    fig_prov = go.Figure(go.Bar(
        x=pd_['Monto'], y=pd_['Proveedor'].str[:32],
        orientation="h", marker_color=PAL[0],
        text=[f"{p}%" for p in pd_['Pct']], textposition="outside"
    ))
    fig_prov.update_layout(
        title="Top 10 Proveedores",
        plot_bgcolor="white", paper_bgcolor="white",
        xaxis=dict(tickformat="$,.0f"),
        yaxis=dict(autorange="reversed"),
        margin=dict(t=45,b=10,l=10,r=10), height=380
    )
    st.plotly_chart(fig_prov, use_container_width=True)

with g6:
    am = (f[f['Mes_num'] > 0]
          .groupby(['Area','Mes_num','Mes_base'])['Monto']
          .sum().reset_index().sort_values('Mes_num'))

    fig_am = px.bar(
        am, x='Mes_base', y='Monto', color='Area',
        color_discrete_map=COL_AREA, barmode="group",
        title="Gasto Mensual por Área (COP)",
        labels={"Monto":"COP","Mes_base":""},
        category_orders={"Mes_base": MESES_ORD}
    )
    fig_am.update_layout(
        plot_bgcolor="white", paper_bgcolor="white",
        yaxis=dict(tickformat="$,.0f"),
        legend=dict(orientation="h", y=-0.15),
        margin=dict(t=45,b=10,l=10,r=10), height=380
    )
    st.plotly_chart(fig_am, use_container_width=True)

# ── SOBRE / SUBEJECUCIÓN ──────────────────────────────────────────
st.markdown('<div class="sec">🔴 Top 10 Sobreejecutados & 🟢 Top 10 Subejecuados por Tipo de Gasto</div>',
            unsafe_allow_html=True)
st.caption("Último mes con datos vs promedio de meses anteriores — detecta qué categorías se desvían.")

_mes_max = int(f[f['Mes_num'] > 0]['Mes_num'].max()) if not f.empty and f[f['Mes_num']>0].shape[0]>0 else 0

if _mes_max > 1:
    _mn = MESES_12.get(_mes_max, str(_mes_max))
    _fc = f[(f['Cuenta'] != 'Sin clasificar') & (f['Mes_num'] > 0)]
    _pm = _fc.groupby(['Cuenta','Area','Mes_num'])['Monto'].sum().reset_index()

    _ul = (_pm[_pm['Mes_num'] == _mes_max]
           .rename(columns={'Monto':'Gasto_Ultimo'}))
    _pr = (_pm[_pm['Mes_num'] < _mes_max]
           .groupby(['Cuenta','Area'])['Monto'].mean().reset_index()
           .rename(columns={'Monto':'Prom_Ant'}))
    _co = _ul.merge(_pr, on=['Cuenta','Area'], how='inner')
    _co['Desv']     = _co['Gasto_Ultimo'] - _co['Prom_Ant']
    _co['Desv_Pct'] = (_co['Desv'] / _co['Prom_Ant'] * 100).round(1)

    _sobre = _co[_co['Desv'] > 0].sort_values('Desv', ascending=False).head(10)
    _sub   = _co[_co['Desv'] < 0].sort_values('Desv').head(10)

    cs1, cs2 = st.columns(2)

    for col, titulo, data, color_bar, color_line, signo in [
        (cs1, f"🔴 Sobreejecutados — {_mn}", _sobre, "rgba(220,38,38,0.2)", "#DC2626", "+"),
        (cs2, f"🟢 Subejecuados — {_mn}",    _sub,   "rgba(5,150,105,0.2)", "#059669", ""),
    ]:
        with col:
            st.markdown(f"#### {titulo}")
            if not data.empty:
                fig_ = go.Figure()
                fig_.add_trace(go.Bar(
                    y=data['Cuenta'].str[:30], x=data['Prom_Ant'],
                    name="Promedio histórico", orientation="h",
                    marker_color=color_bar,
                    marker_line=dict(color=color_line, width=1.5)
                ))
                fig_.add_trace(go.Bar(
                    y=data['Cuenta'].str[:30], x=data['Gasto_Ultimo'],
                    name="Último mes", orientation="h",
                    marker_color=color_line,
                    text=[f"{signo}{cop(v)} ({p:+.0f}%)"
                          for v,p in zip(data['Desv'], data['Desv_Pct'])],
                    textposition="outside"
                ))
                fig_.update_layout(
                    barmode="overlay", plot_bgcolor="white", paper_bgcolor="white",
                    xaxis=dict(tickformat="$,.0f"),
                    yaxis=dict(autorange="reversed"),
                    legend=dict(orientation="h", y=-0.18),
                    margin=dict(t=10,b=10,l=10,r=100),
                    height=max(300, len(data)*42)
                )
                st.plotly_chart(fig_, use_container_width=True)

                tbl = data[['Cuenta','Area','Gasto_Ultimo','Prom_Ant','Desv','Desv_Pct']].copy()
                for c in ['Gasto_Ultimo','Prom_Ant','Desv']:
                    tbl[c] = tbl[c].apply(cop)
                tbl['Desv_Pct'] = tbl['Desv_Pct'].apply(lambda x: f"{x:+.1f}%")
                tbl.columns = ['Tipo de Gasto','Área',f'Gasto {_mn}',
                               'Prom. Histórico','Desviación COP','Desv. %']
                st.dataframe(tbl, use_container_width=True, hide_index=True)
            else:
                st.info("No hay datos suficientes en el período seleccionado.")
else:
    st.info("Selecciona al menos 2 meses para ver el análisis de sobre/subejecución.")

# ── TABLA DETALLE ─────────────────────────────────────────────────
st.markdown('<div class="sec">Detalle de Registros</div>', unsafe_allow_html=True)

tbl = f[['Fuente','Proveedor','Area','Sucursal','Cuenta','Tipo',
          'Mes_nom','Monto','Gasto_Neto']].copy()
tbl['Monto']      = tbl['Monto'].apply(cop)
tbl['Gasto_Neto'] = tbl['Gasto_Neto'].apply(cop)
tbl = tbl.rename(columns={
    'Area':'Área','Mes_nom':'Mes','Tipo':'Fijo/Variable',
    'Monto':'Monto Factura (COP)','Gasto_Neto':'Gasto Neto (COP)'
})

st.dataframe(
    tbl.style.apply(
        lambda r: ["background-color:#FFF5F5;color:#991B1B"]*len(r)
                  if r['Fuente']=='NO PAGAR' else [""]*len(r), axis=1
    ),
    use_container_width=True, height=360, hide_index=True
)
st.caption(
    f"🔵 REGISTRO: {len(f[f['Fuente']=='REGISTRO'])} · "
    f"🔴 NO PAGAR: {len(f[f['Fuente']=='NO PAGAR'])} · "
    f"Total: {len(tbl)} registros · Montos en COP"
)

st.divider()
st.caption("📊 Craft Logistics · Presupuesto hardcodeado (ADM/IT/GH) · Facturas: SharePoint")
