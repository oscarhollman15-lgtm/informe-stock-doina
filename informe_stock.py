import base64
import pandas as pd
import streamlit as st
import streamlit_antd_components as sac

st.set_page_config(
    page_title="Stock Doina — Ventas",
    page_icon="🥩",
    layout="wide",
    initial_sidebar_state="collapsed",
)

ARCHIVO = "stock doina.xlsx"
LOGO    = "DOINA Marca - negro.png"

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
@st.cache_data
def cargar_resumen_stock():
    raw = pd.read_excel(ARCHIVO, sheet_name="RESUMEN STOCK", header=None)
    df = raw.iloc[4:].copy()
    df.columns = ["Producto", "Exp_Unidades", "Exp_KG", "Prod_Unidades", "Prod_KG", "Proc_Unidades", "Proc_KG"]
    df = df.dropna(subset=["Producto"])
    df["Producto"] = df["Producto"].astype(str).str.strip()
    df = df[df["Producto"] != ""]
    for col in df.columns[1:]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    return df.reset_index(drop=True)

@st.cache_data
def cargar_lotes(hoja):
    raw = pd.read_excel(ARCHIVO, sheet_name=hoja, header=None)
    header_row = None
    for i, row in raw.iterrows():
        if row.astype(str).str.contains("FECHA SALAZON", case=False, na=False).any():
            header_row = i
            break
    if header_row is None:
        return pd.DataFrame()
    df = pd.read_excel(ARCHIVO, sheet_name=hoja, header=header_row)
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

@st.cache_data
def cargar_todos_los_lotes():
    dfs = [cargar_lotes(h) for h in HOJAS_PRODUCTO]
    return pd.concat([d for d in dfs if not d.empty], ignore_index=True)

@st.cache_data
def cargar_expedicion_detalle():
    raw = pd.read_excel(ARCHIVO, sheet_name="stock expedicion", header=None)
    # detectar fila de encabezado: primera fila con al menos 4 celdas no nulas
    header_row = 0
    for i, row in raw.iterrows():
        if row.notna().sum() >= 4:
            header_row = i
            break
    df = pd.read_excel(ARCHIVO, sheet_name="stock expedicion", header=header_row)
    df.columns = df.columns.astype(str).str.strip()
    df = df.dropna(how="all")
    # la columna H es el índice 7 dentro de las columnas del df
    return df.reset_index(drop=True)

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
df_res   = cargar_resumen_stock()
df_lotes = cargar_todos_los_lotes()
hoy      = pd.Timestamp.today().normalize()

fecha_excel = pd.read_excel(ARCHIVO, sheet_name="RESUMEN STOCK", header=None).iloc[0, 1]
fecha_str   = pd.to_datetime(fecha_excel).strftime("%d/%m/%Y") if pd.notna(fecha_excel) else "—"

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

if st.button("🔄 Recargar datos", help="Limpia el cache y recarga el Excel"):
    st.cache_data.clear()
    st.rerun()

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

    df_p = df_res[df_res["Prod_Unidades"] > 0][["Producto", "Prod_Unidades", "Prod_KG"]].copy()
    df_p = df_p.sort_values("Prod_Unidades", ascending=False)
    if buscar:
        df_p = df_p[df_p["Producto"].str.contains(buscar, case=False, na=False)]

    if df_p.empty:
        sac.result(label="Sin resultados", description="No hay productos que coincidan con la búsqueda.", status="info")
    else:
        df_p_show = df_p.copy()
        df_p_show["Prod_Unidades"] = df_p_show["Prod_Unidades"].apply(fmt_u)
        df_p_show["Prod_KG"]       = df_p_show["Prod_KG"].apply(fmt_kg)
        df_p_show.columns = ["Producto", "Unidades", "KG"]
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown('<div class="card-title">Productos disponibles en producción</div>', unsafe_allow_html=True)
        st.dataframe(df_p_show, width="stretch", hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 1 — EXPEDICIÓN
# ═══════════════════════════════════════════════════════════════════════════════
elif tab == 1:
    sac.divider(label="Stock en Expedición", icon="truck", color="#e87722")

    buscar = st.text_input("🔍 Buscar producto", placeholder="Escribí parte del nombre...", key="buscar_exp", label_visibility="collapsed")

    df_e = df_res[df_res["Exp_Unidades"] > 0][["Producto", "Exp_Unidades", "Exp_KG"]].copy()
    df_e = df_e.sort_values("Exp_Unidades", ascending=False)
    if buscar:
        df_e = df_e[df_e["Producto"].str.contains(buscar, case=False, na=False)]

    if df_e.empty:
        sac.result(label="Sin resultados", description="No hay productos que coincidan con la búsqueda.", status="info")
    else:
        df_e_show = df_e.copy()
        df_e_show["Exp_Unidades"] = df_e_show["Exp_Unidades"].apply(fmt_u)
        df_e_show["Exp_KG"]       = df_e_show["Exp_KG"].apply(fmt_kg)
        df_e_show.columns = ["Producto", "Unidades", "KG"]
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown('<div class="card-title">Productos disponibles en expedición</div>', unsafe_allow_html=True)
        st.dataframe(df_e_show, width="stretch", hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)

    # ── Detalle ────────────────────────────────────────────────────────────────
    sac.divider(label="Detalle", icon="table", color="#555")
    df_mis = cargar_expedicion_detalle()

    if df_mis.empty:
        sac.result(label="Sin datos", description="No se encontraron datos en la hoja de expedición.", status="info")
    else:
        buscar_mis = st.text_input("🔍 Buscar producto", placeholder="Escribí parte del nombre...", key="buscar_mis", label_visibility="collapsed")

        # columna H = posición 7 dentro del dataframe
        col_h_idx = 7
        col_h_name = df_mis.columns[col_h_idx] if len(df_mis.columns) > col_h_idx else None

        df_mis_filtered = df_mis.copy()
        if buscar_mis:
            # buscar en todas las columnas de texto
            mask = df_mis_filtered.apply(
                lambda col: col.astype(str).str.contains(buscar_mis, case=False, na=False)
            ).any(axis=1)
            df_mis_filtered = df_mis_filtered[mask]

        def resaltar_vida_util(row):
            if col_h_name:
                val = pd.to_numeric(row[col_h_name], errors="coerce")
                if pd.notna(val) and val < 80:
                    return ["background-color: #e87722; color: #fff; font-weight: 600"] * len(row)
            return [""] * len(row)

        styled = df_mis_filtered.style.apply(resaltar_vida_util, axis=1)
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown(f'<div class="card-title">Stock expedición · <span style="color:#e87722">naranja</span> = vida útil &lt; 80%</div>', unsafe_allow_html=True)
        st.dataframe(styled, width="stretch", hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 2 — RESUMEN
# ═══════════════════════════════════════════════════════════════════════════════
elif tab == 2:
    sac.divider(label="Resumen General — Producción + Expedición", icon="clipboard2-data", color="#e87722")

    buscar = st.text_input("🔍 Buscar producto", placeholder="Escribí parte del nombre...", key="buscar_res", label_visibility="collapsed")

    df_r = df_res.copy()
    df_r["Total_Unidades"] = df_r["Prod_Unidades"] + df_r["Exp_Unidades"]
    df_r["Total_KG"]       = df_r["Prod_KG"]       + df_r["Exp_KG"]
    df_r = df_r[df_r["Total_Unidades"] > 0][["Producto", "Total_Unidades", "Total_KG"]].sort_values("Total_Unidades", ascending=False)

    if buscar:
        df_r = df_r[df_r["Producto"].str.contains(buscar, case=False, na=False)]

    if df_r.empty:
        sac.result(label="Sin resultados", description="No hay productos que coincidan con la búsqueda.", status="info")
    else:
        df_r_show = df_r.copy()
        df_r_show["Total_Unidades"] = df_r_show["Total_Unidades"].apply(fmt_u)
        df_r_show["Total_KG"]       = df_r_show["Total_KG"].apply(fmt_kg)
        df_r_show.columns = ["Producto", "Total Unidades", "Total KG"]
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown('<div class="card-title">Stock total disponible por producto</div>', unsafe_allow_html=True)
        st.dataframe(df_r_show, width="stretch", hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 3 — STOCK EN PROCESO
# ═══════════════════════════════════════════════════════════════════════════════
elif tab == 3:
    sac.divider(label="Stock en Proceso", icon="hourglass-split", color="#e87722")

    # Resumen por producto
    df_proc = df_res[df_res["Proc_Unidades"] > 0][["Producto", "Proc_Unidades", "Proc_KG"]].copy()
    df_proc = df_proc.sort_values("Proc_Unidades", ascending=False).reset_index(drop=True)

    df_proc_show = df_proc.copy()
    df_proc_show["Proc_Unidades"] = df_proc_show["Proc_Unidades"].apply(fmt_u)
    df_proc_show["Proc_KG"]       = df_proc_show["Proc_KG"].apply(fmt_kg)
    df_proc_show.columns = ["Producto", "Unidades en proceso", "KG en proceso"]

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="card-title">Resumen total por producto</div>', unsafe_allow_html=True)
    st.dataframe(df_proc_show, width="stretch", hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

    sac.divider(label="Detalle de lotes", icon="list-ul", color="#555")

    productos_proceso = df_proc["Producto"].tolist()
    producto_sel = st.selectbox(
        "Seleccioná un producto:",
        options=["— Todos —"] + productos_proceso,
        key="sel_proceso",
    )

    col_check, col_fecha = st.columns([1, 3])
    with col_check:
        usar_fecha = st.checkbox("Filtrar por fecha de disponibilidad", value=False)
    with col_fecha:
        fecha_hasta = None
        if usar_fecha:
            fecha_hasta = st.date_input(
                "Disponibles hasta:",
                value=hoy + pd.Timedelta(days=30),
                min_value=hoy.date(),
                key="fecha_proc",
            )

    if producto_sel == "— Todos —":
        df_det = df_lotes.copy()
    else:
        hoja_match = encontrar_hoja(producto_sel)
        if hoja_match:
            df_det = cargar_lotes(hoja_match)
        else:
            df_det = pd.DataFrame()
            sac.result(
                label="Hoja no encontrada",
                description=f"No se encontró la hoja de lotes para '{producto_sel}'.",
                status="warning",
            )

    if not df_det.empty and "FECHA TERMINACION" in df_det.columns and fecha_hasta:
        df_det = df_det[df_det["FECHA TERMINACION"] <= pd.Timestamp(fecha_hasta)]

    if df_det.empty:
        sac.result(label="Sin lotes", description="No hay lotes para los filtros seleccionados.", status="info")
    else:
        n_lotes   = str(len(df_det))
        n_unid    = fmt_u(df_det["UNIDADES"].sum()) if "UNIDADES" in df_det.columns else "—"
        n_kg      = fmt_kg(df_det["KILOS"].sum())   if "KILOS"    in df_det.columns else "—"
        mostrar_mini_metricas(n_lotes, n_unid, n_kg)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown('<div class="card-title">Lotes · ordenados por fecha de disponibilidad</div>', unsafe_allow_html=True)
        st.dataframe(tabla_lotes(df_det), width="stretch", hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)
