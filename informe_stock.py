import base64
import os
import shutil
import tempfile
from pathlib import Path
import pandas as pd
import requests
import streamlit as st
import streamlit_antd_components as sac

st.set_page_config(
    page_title="Stock Doina — Ventas",
    page_icon="🥩",
    layout="wide",
    initial_sidebar_state="collapsed",
)

BASE_DIR = Path(__file__).resolve().parent
ARCHIVO  = str(BASE_DIR / "stock doina.xlsx")
LOGO     = str(BASE_DIR / "DOINA Marca - negro.png")

# URL de OneDrive cuando corre en la nube (se configura en Streamlit Cloud secrets)
try:
    EXCEL_URL = st.secrets["EXCEL_URL"]
except Exception:
    EXCEL_URL = os.environ.get("EXCEL_URL", "")

def leer_excel(sheet_name, **kwargs):
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp.close()
    try:
        if EXCEL_URL:
            r = requests.get(EXCEL_URL, timeout=30)
            r.raise_for_status()
            with open(tmp.name, "wb") as f:
                f.write(r.content)
        else:
            shutil.copy2(ARCHIVO, tmp.name)
        return pd.read_excel(tmp.name, sheet_name=sheet_name, **kwargs)
    finally:
        try:
            os.unlink(tmp.name)
        except Exception:
            pass

HOJAS_PRODUCTO = [
    "jamon parma", "jamon iberico", "SALAME FUET",
    "PANCETA SALADA", "BONDIOLA CURADA", "LOMO CURADO", "RAPIDISIMOS",
]

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

  html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

  .stApp, [data-testid="stAppViewContainer"] { background-color: #111318; }
  [data-testid="stHeader"] { background: transparent; }
  section[data-testid="stSidebar"] { background: #1a1c23; }

  /* ── Header ── */
  .header-box {
    background: #1a1c23;
    border-radius: 20px;
    padding: 22px 32px;
    margin-bottom: 24px;
    border: 1px solid #2a2c35;
    box-shadow: 0 4px 24px rgba(232,119,34,0.12);
    display: flex;
    align-items: center;
    gap: 20px;
  }
  .header-title { color: #fff; font-size: 1.7rem; font-weight: 700; margin: 0; letter-spacing: -0.5px; }
  .header-sub   { color: #888; font-size: 0.85rem; margin: 3px 0 0; }
  .header-badge {
    margin-left: auto;
    background: rgba(232,119,34,0.15);
    border: 1px solid rgba(232,119,34,0.4);
    color: #e87722;
    padding: 6px 14px;
    border-radius: 20px;
    font-size: 0.8rem;
    font-weight: 600;
    white-space: nowrap;
  }

  /* ── Cards de datos ── */
  .card {
    background: #1a1c23;
    border: 1px solid #2a2c35;
    border-radius: 16px;
    padding: 20px 24px;
    margin-bottom: 16px;
  }
  .card-title {
    color: #888;
    font-size: 0.72rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    margin: 0 0 14px;
    display: flex;
    align-items: center;
    gap: 8px;
  }
  .card-title::before {
    content: '';
    display: inline-block;
    width: 3px; height: 14px;
    background: #e87722;
    border-radius: 2px;
  }

  /* ── Mini métricas ── */
  .mini-metric {
    background: #22242d;
    border: 1px solid #2d3040;
    border-radius: 12px;
    padding: 14px 18px;
    text-align: center;
  }
  .mini-metric .label { color: #777; font-size: 0.72rem; text-transform: uppercase; letter-spacing: 0.06em; }
  .mini-metric .value { color: #fff; font-size: 1.5rem; font-weight: 700; margin: 4px 0 0; }

  /* ── Tablas ── */
  [data-testid="stDataFrame"] {
    border-radius: 12px !important;
    overflow: hidden;
    border: 1px solid #2a2c35 !important;
  }
  [data-testid="stDataFrame"] th {
    background: #22242d !important;
    color: #999 !important;
    font-size: 0.75rem !important;
    font-weight: 600 !important;
    text-transform: uppercase !important;
    letter-spacing: 0.05em !important;
  }

  /* ── Search input ── */
  [data-testid="stTextInput"] input {
    background: #1e2028 !important;
    border: 1px solid #2d3040 !important;
    border-radius: 10px !important;
    color: #e0e0e0 !important;
    font-family: 'Inter', sans-serif !important;
  }
  [data-testid="stTextInput"] input:focus {
    border-color: #e87722 !important;
    box-shadow: 0 0 0 2px rgba(232,119,34,0.2) !important;
  }

  /* ── Selectbox ── */
  [data-testid="stSelectbox"] > div > div {
    background: #1e2028 !important;
    border: 1px solid #2d3040 !important;
    border-radius: 10px !important;
    color: #e0e0e0 !important;
  }

  /* ── Checkbox ── */
  [data-testid="stCheckbox"] label { color: #aaa !important; font-size: 0.88rem !important; }

  /* ── Scrollbar ── */
  ::-webkit-scrollbar { width: 6px; height: 6px; }
  ::-webkit-scrollbar-track { background: #1a1c23; }
  ::-webkit-scrollbar-thumb { background: #3a3d4a; border-radius: 3px; }
  ::-webkit-scrollbar-thumb:hover { background: #e87722; }

  /* ── antd tabs override ── */
  .ant-tabs-tab { font-family: 'Inter', sans-serif !important; font-size: 0.9rem !important; }
  .ant-tabs-ink-bar { background: #e87722 !important; }
  .ant-tabs-tab-active .ant-tabs-tab-btn { color: #e87722 !important; }

  h2, h3 { color: #e0e0e0 !important; font-family: 'Inter', sans-serif !important; }
</style>
""", unsafe_allow_html=True)

# ── Helpers ───────────────────────────────────────────────────────────────────
def logo_b64():
    with open(LOGO, "rb") as f:
        return base64.b64encode(f.read()).decode()

def fmt_u(n):
    try:    return f"{int(n):,}".replace(",", ".")
    except: return "—"

def fmt_kg(n):
    try:    return f"{float(n):,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except: return "—"

def tabla_lotes(df):
    cols = [c for c in ["N° LOTE", "FECHA TERMINACION", "UNIDADES", "KILOS", "TAMAÑO"] if c in df.columns]
    t = df[cols].copy()
    if "FECHA TERMINACION" in t.columns:
        t = t.sort_values("FECHA TERMINACION")
        t["FECHA TERMINACION"] = t["FECHA TERMINACION"].dt.strftime("%d/%m/%Y")
    if "UNIDADES" in t.columns:
        t["UNIDADES"] = t["UNIDADES"].apply(lambda x: fmt_u(int(x)) if pd.notna(x) else "—")
    if "KILOS" in t.columns:
        t["KILOS"] = t["KILOS"].apply(lambda x: fmt_kg(x) if pd.notna(x) else "—")
    t.columns = [
        c.replace("FECHA TERMINACION", "Disponible el")
         .replace("UNIDADES", "Unidades").replace("KILOS", "KG").replace("TAMAÑO", "Tamaño")
        for c in t.columns
    ]
    return t

def mostrar_mini_metricas(lotes, unidades, kg):
    cols = st.columns(3)
    datos = [("Lotes", lotes), ("Unidades", unidades), ("KG", kg)]
    for col, (label, val) in zip(cols, datos):
        col.markdown(f"""
        <div class="mini-metric">
          <div class="label">{label}</div>
          <div class="value">{val}</div>
        </div>""", unsafe_allow_html=True)

def card(titulo, contenido_fn):
    st.markdown(f'<div class="card"><div class="card-title">{titulo}</div></div>', unsafe_allow_html=True)
    contenido_fn()

# ── ETL ───────────────────────────────────────────────────────────────────────
@st.cache_data(ttl=300)
def cargar_lotes(hoja):
    raw = leer_excel( sheet_name=hoja, header=None)
    header_row = None
    for i, row in raw.iterrows():
        if row.astype(str).str.contains("FECHA SALAZON", case=False, na=False).any():
            header_row = i
            break
    if header_row is None:
        return pd.DataFrame()
    df = leer_excel( sheet_name=hoja, header=header_row)
    df.columns = (
        df.columns.astype(str).str.strip()
        .str.replace("N\x83 LOTE",  "N° LOTE",  regex=False)
        .str.replace("TAMA\x83 O", "TAMAÑO",   regex=False)
        .str.replace("N\x86 LOTE",  "N° LOTE",  regex=False)
        .str.replace("N\xb0 LOTE",  "N° LOTE",  regex=False)
        .str.replace("N� LOTE", "N° LOTE",  regex=False)
        .str.replace("TAMA� O", "TAMAÑO",   regex=False)
        .str.replace("N\x83 LOTE",  "N° LOTE",  regex=False)
    )
    # normalización extra por si quedan caracteres raros
    df.columns = [c if "LOTE" not in c or "°" in c else "N° LOTE" if "LOTE" in c and c.startswith("N") else c for c in df.columns]
    df = df.dropna(how="all")
    for col in ["KILOS", "UNIDADES"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    for col in ["FECHA SALAZON", "FECHA TERMINACION"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    df = df.drop(columns=[c for c in df.columns if "Unnamed" in c], errors="ignore")
    df = df.dropna(subset=["KILOS"])
    df["PRODUCTO"] = hoja.upper()
    return df.reset_index(drop=True)

@st.cache_data(ttl=300)
def cargar_todos_los_lotes():
    dfs = [cargar_lotes(h) for h in HOJAS_PRODUCTO]
    return pd.concat([d for d in dfs if not d.empty], ignore_index=True)

@st.cache_data(ttl=300)
def cargar_expedicion_detalle():
    raw = leer_excel( sheet_name="stock expedicion", header=None)
    # detectar fila de encabezado: primera fila con al menos 4 celdas no nulas
    header_row = 0
    for i, row in raw.iterrows():
        if row.notna().sum() >= 4:
            header_row = i
            break
    df = leer_excel( sheet_name="stock expedicion", header=header_row)
    df.columns = df.columns.astype(str).str.strip()
    df = df.dropna(how="all")
    # la columna H es el índice 7 dentro de las columnas del df
    return df.reset_index(drop=True)

@st.cache_data(ttl=300)
def cargar_stock_expedicion():
    raw = leer_excel(sheet_name="stock expedicion", header=None)
    header_row = None
    for i, row in raw.iterrows():
        if row.astype(str).str.contains("Descripcion", case=False, na=False).any():
            header_row = i
            break
    if header_row is None:
        return pd.DataFrame(columns=["Producto", "Unidades", "KG"])
    df = leer_excel(sheet_name="stock expedicion", header=header_row)
    df.columns = df.columns.astype(str).str.strip()
    df = df.dropna(how="all")
    for col in ["Uni", "Kg"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    result = df.groupby("Descripcion", as_index=False).agg(
        Unidades=("Uni", "sum"), KG=("Kg", "sum")
    ).rename(columns={"Descripcion": "Producto"})
    return result[result["Unidades"] > 0].sort_values("Unidades", ascending=False).reset_index(drop=True)

@st.cache_data(ttl=300)
def cargar_stock_produccion():
    raw = leer_excel(sheet_name="stock terminado produccion", header=None)
    header_row = None
    for i, row in raw.iterrows():
        vals = row.astype(str).str.strip().tolist()
        if "Producto" in vals and "Unidades" in vals:
            header_row = i
            break
    if header_row is None:
        return pd.DataFrame(columns=["Producto", "Unidades", "KG"])
    df = leer_excel(sheet_name="stock terminado produccion", header=header_row)
    df.columns = df.columns.astype(str).str.strip()
    df = df.dropna(how="all")
    df["Unidades"] = pd.to_numeric(df.get("Unidades", pd.Series(dtype=float)), errors="coerce").fillna(0)
    kg_neto  = pd.to_numeric(df.get("Kg Neto",  pd.Series(dtype=float)), errors="coerce").fillna(0)
    kg_bruto = pd.to_numeric(df.get("Kg Bruto", pd.Series(dtype=float)), errors="coerce").fillna(0)
    df["KG_val"] = kg_neto.where(kg_neto > 0, kg_bruto)
    result = df.groupby("Producto", as_index=False).agg(Unidades=("Unidades", "sum"), KG=("KG_val", "sum"))
    return result[result["Unidades"] > 0].sort_values("Unidades", ascending=False).reset_index(drop=True)

@st.cache_data(ttl=300)
def cargar_stock_proceso():
    df_all = cargar_todos_los_lotes()
    if df_all.empty:
        return pd.DataFrame(columns=["Producto", "Unidades", "KG"])
    result = df_all.groupby("PRODUCTO", as_index=False).agg(
        Unidades=("UNIDADES", "sum"), KG=("KILOS", "sum")
    ).rename(columns={"PRODUCTO": "Producto"})
    return result[result["Unidades"] > 0].sort_values("Unidades", ascending=False).reset_index(drop=True)

MAPA_KEYWORDS = {
    "jamon parma":     ["parma"],
    "jamon iberico":   ["iberico"],
    "SALAME FUET":     ["fuet", "salame"],
    "PANCETA SALADA":  ["panceta"],
    "BONDIOLA CURADA": ["bondiola"],
    "LOMO CURADO":     ["lomo"],
    "RAPIDISIMOS":     ["pernil", "delipork", "rapidisimo"],
}

def encontrar_hoja(nombre_resumen):
    nombre_up = nombre_resumen.upper()
    for hoja, keywords in MAPA_KEYWORDS.items():
        if any(kw.upper() in nombre_up for kw in keywords):
            return hoja
    return None

# ── Carga ─────────────────────────────────────────────────────────────────────
df_lotes      = cargar_todos_los_lotes()
df_expedicion = cargar_stock_expedicion()
df_produccion = cargar_stock_produccion()
df_proceso    = cargar_stock_proceso()
hoy           = pd.Timestamp.today().normalize()
try:
    _raw_fecha = leer_excel(sheet_name="stock expedicion", header=None)
    fecha_str  = pd.to_datetime(_raw_fecha.iloc[0, 1]).strftime("%d/%m/%Y")
except Exception:
    fecha_str  = pd.Timestamp.fromtimestamp(os.path.getmtime(ARCHIVO)).strftime("%d/%m/%Y")

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="header-box">
  <img src="data:image/png;base64,{logo_b64()}" style="height:64px; border-radius:50%; border:2px solid #e87722;">
  <div>
    <p class="header-title">Stock — Ventas</p>
    <p class="header-sub">Doina · Pasión desde la crianza</p>
  </div>
  <div class="header-badge">📅 Actualizado al {fecha_str}</div>
</div>
""", unsafe_allow_html=True)

# ── Navegación principal con antd ─────────────────────────────────────────────
tab = sac.tabs([
    sac.TabsItem("Producción",    icon="building-gear"),
    sac.TabsItem("Expedición",    icon="truck"),
    sac.TabsItem("Resumen",       icon="clipboard2-data"),
    sac.TabsItem("En Proceso",    icon="hourglass-split"),
], color="#e87722", size="md", align="start", use_container_width=False, return_index=True)

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 0 — PRODUCCIÓN
# ═══════════════════════════════════════════════════════════════════════════════
if tab == 0:
    sac.divider(label="Stock en Producción", icon="building-gear", color="#e87722")
    buscar = st.text_input("🔍 Buscar producto", placeholder="Escribí parte del nombre...", key="buscar_prod", label_visibility="collapsed")
    df_p = df_produccion.copy()
    if buscar:
        df_p = df_p[df_p["Producto"].str.contains(buscar, case=False, na=False)]
    if df_p.empty:
        sac.result(label="Sin resultados", description="No hay productos que coincidan con la búsqueda.", status="info")
    else:
        df_p_show = df_p.copy()
        df_p_show["Unidades"] = df_p_show["Unidades"].apply(fmt_u)
        df_p_show["KG"]       = df_p_show["KG"].apply(fmt_kg)
        st.markdown('<div class="card"><div class="card-title">Productos disponibles en producción</div>', unsafe_allow_html=True)
        st.dataframe(df_p_show, width="stretch", hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 1 — EXPEDICIÓN
# ═══════════════════════════════════════════════════════════════════════════════
elif tab == 1:
    sac.divider(label="Stock en Expedición", icon="truck", color="#e87722")
    buscar = st.text_input("🔍 Buscar producto", placeholder="Escribí parte del nombre...", key="buscar_exp", label_visibility="collapsed")
    df_e = df_expedicion.copy()
    if buscar:
        df_e = df_e[df_e["Producto"].str.contains(buscar, case=False, na=False)]
    if df_e.empty:
        sac.result(label="Sin resultados", description="No hay productos que coincidan con la búsqueda.", status="info")
    else:
        df_e_show = df_e.copy()
        df_e_show["Unidades"] = df_e_show["Unidades"].apply(fmt_u)
        df_e_show["KG"]       = df_e_show["KG"].apply(fmt_kg)
        st.markdown('<div class="card"><div class="card-title">Productos disponibles en expedición</div>', unsafe_allow_html=True)
        st.dataframe(df_e_show, width="stretch", hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)

    sac.divider(label="Detalle", icon="table", color="#555")
    df_mis = cargar_expedicion_detalle()
    if df_mis.empty:
        sac.result(label="Sin datos", description="No se encontraron datos en la hoja de expedición.", status="info")
    else:
        buscar_mis = st.text_input("🔍 Buscar", placeholder="Escribí parte del nombre...", key="buscar_mis", label_visibility="collapsed")
        col_h_name = df_mis.columns[7] if len(df_mis.columns) > 7 else None
        df_f = df_mis.copy()
        if buscar_mis:
            mask = df_f.apply(lambda c: c.astype(str).str.contains(buscar_mis, case=False, na=False)).any(axis=1)
            df_f = df_f[mask]
        def resaltar_vida_util(row):
            if col_h_name:
                val = pd.to_numeric(row[col_h_name], errors="coerce")
                if pd.notna(val) and val < 80:
                    return ["background-color:#e87722;color:#fff;font-weight:600"] * len(row)
            return [""] * len(row)
        styled = df_f.style.apply(resaltar_vida_util, axis=1)
        st.markdown('<div class="card"><div class="card-title">Stock expedición · <span style="color:#e87722">naranja</span> = vida útil &lt; 80%</div>', unsafe_allow_html=True)
        st.dataframe(styled, width="stretch", hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 2 — RESUMEN
# ═══════════════════════════════════════════════════════════════════════════════
elif tab == 2:
    sac.divider(label="Resumen General — Producción + Expedición", icon="clipboard2-data", color="#e87722")
    buscar = st.text_input("🔍 Buscar producto", placeholder="Escribí parte del nombre...", key="buscar_res", label_visibility="collapsed")
    df_e_r = df_expedicion[["Producto","Unidades","KG"]].copy(); df_e_r["Fuente"] = "Expedición"
    df_p_r = df_produccion[["Producto","Unidades","KG"]].copy(); df_p_r["Fuente"] = "Producción"
    df_r = pd.concat([df_e_r, df_p_r], ignore_index=True).sort_values("Unidades", ascending=False)
    if buscar:
        df_r = df_r[df_r["Producto"].str.contains(buscar, case=False, na=False)]
    if df_r.empty:
        sac.result(label="Sin resultados", description="No hay productos que coincidan con la búsqueda.", status="info")
    else:
        df_r_show = df_r[["Fuente","Producto","Unidades","KG"]].copy()
        df_r_show["Unidades"] = df_r_show["Unidades"].apply(fmt_u)
        df_r_show["KG"]       = df_r_show["KG"].apply(fmt_kg)
        st.markdown('<div class="card"><div class="card-title">Stock total disponible por producto</div>', unsafe_allow_html=True)
        st.dataframe(df_r_show, width="stretch", hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 3 — EN PROCESO
# ═══════════════════════════════════════════════════════════════════════════════
elif tab == 3:
    sac.divider(label="Stock en Proceso", icon="hourglass-split", color="#e87722")

    total_u  = df_proceso["Unidades"].sum()
    total_kg = df_proceso["KG"].sum()
    cols_kpi = st.columns(3)
    for col_k, (label, val) in zip(cols_kpi, [
        ("Productos en proceso", fmt_u(len(df_proceso))),
        ("Total unidades",       fmt_u(total_u)),
        ("Total KG",             fmt_kg(total_kg)),
    ]):
        col_k.markdown(f'<div class="mini-metric"><div class="label">{label}</div><div class="value">{val}</div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    df_proc_show = df_proceso.copy()
    df_proc_show["Unidades"] = df_proc_show["Unidades"].apply(fmt_u)
    df_proc_show["KG"]       = df_proc_show["KG"].apply(fmt_kg)
    df_proc_show.columns = ["Producto", "Unidades en proceso", "KG en proceso"]
    st.markdown('<div class="card"><div class="card-title">Resumen por producto — stock en proceso</div>', unsafe_allow_html=True)
    st.dataframe(df_proc_show, width="stretch", hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

    sac.divider(label="Detalle de lotes", icon="list-ul", color="#555")
    col_sel, col_filt = st.columns([2, 3])
    with col_sel:
        producto_sel = st.selectbox("Producto:", options=["— Todos —"] + df_proceso["Producto"].tolist(), key="sel_proceso")
    with col_filt:
        col_check, col_fecha = st.columns([1, 2])
        with col_check:
            usar_fecha = st.checkbox("Filtrar por fecha de disponibilidad", value=False, key="usar_fecha_proc")
        with col_fecha:
            fecha_hasta = None
            if usar_fecha:
                fecha_hasta = st.date_input("Disponibles hasta:", value=hoy + pd.Timedelta(days=60), min_value=hoy.date(), key="fecha_proc")

    buscar_lote = st.text_input("Buscar lote o producto", placeholder="Ej: 262E/01, iberico...", key="buscar_lote_proc", label_visibility="collapsed")

    if producto_sel == "— Todos —":
        df_det = df_lotes.copy()
    else:
        hoja_match = encontrar_hoja(producto_sel)
        df_det = cargar_lotes(hoja_match) if hoja_match else pd.DataFrame()

    if not df_det.empty:
        if "FECHA TERMINACION" in df_det.columns and fecha_hasta:
            df_det = df_det[df_det["FECHA TERMINACION"] <= pd.Timestamp(fecha_hasta)]
        if buscar_lote:
            mask = pd.Series(False, index=df_det.index)
            for col in ["N° LOTE", "PRODUCTO", "TAMAÑO"]:
                if col in df_det.columns:
                    mask |= df_det[col].astype(str).str.contains(buscar_lote, case=False, na=False)
            df_det = df_det[mask]

    if df_det.empty:
        sac.result(label="Sin lotes", description="No hay lotes para los filtros seleccionados.", status="info")
    else:
        n_lotes = str(len(df_det))
        n_unid  = fmt_u(df_det["UNIDADES"].sum()) if "UNIDADES" in df_det.columns else "—"
        n_kg    = fmt_kg(df_det["KILOS"].sum())   if "KILOS" in df_det.columns else "—"
        cols_m  = st.columns(3)
        for col_m, (label, val) in zip(cols_m, [("Lotes", n_lotes), ("Unidades", n_unid), ("KG", n_kg)]):
            col_m.markdown(f'<div class="mini-metric"><div class="label">{label}</div><div class="value">{val}</div></div>', unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

        cols_t = [c for c in ["N° LOTE", "FECHA TERMINACION", "UNIDADES", "KILOS", "TAMAÑO"] if c in df_det.columns]
        df_tabla = df_det[cols_t].copy()
        if "FECHA TERMINACION" in df_tabla.columns:
            df_tabla = df_tabla.sort_values("FECHA TERMINACION")
            df_tabla["FECHA TERMINACION"] = df_tabla["FECHA TERMINACION"].dt.strftime("%d/%m/%Y")
        if "UNIDADES" in df_tabla.columns:
            df_tabla["UNIDADES"] = df_tabla["UNIDADES"].apply(lambda x: fmt_u(int(x)) if pd.notna(x) else "—")
        if "KILOS" in df_tabla.columns:
            df_tabla["KILOS"] = df_tabla["KILOS"].apply(lambda x: fmt_kg(x) if pd.notna(x) else "—")
        df_tabla.columns = [c.replace("FECHA TERMINACION","Disponible el").replace("UNIDADES","Unidades").replace("KILOS","KG").replace("TAMAÑO","Tamaño") for c in df_tabla.columns]

        def resaltar_disponible(row):
            if "Disponible el" in row.index:
                try:
                    fec = pd.to_datetime(row["Disponible el"], dayfirst=True, errors="coerce")
                    if pd.notna(fec) and fec <= hoy + pd.Timedelta(days=30):
                        return ["background-color:#e87722;color:#fff;font-weight:600"] * len(row)
                except Exception:
                    pass
            return [""] * len(row)

        styled = df_tabla.style.apply(resaltar_disponible, axis=1)
        st.markdown('<div class="card"><div class="card-title">Lotes · ordenados por fecha de disponibilidad</div>', unsafe_allow_html=True)
        st.dataframe(styled, width="stretch", hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)

