import streamlit as st
import pandas as pd
import numpy as np
import base64
import io
from pathlib import Path
import plotly.graph_objects as go
import plotly.express as px
from utils_data import load_sheet, standardize_ptm, standardize_generic
from utils_charts import (
    chart_timeline_remessa, chart_faixa_aberto, chart_hist_aberto,
    chart_atraso, chart_valor_destino, chart_pendente_destino,
    chart_tipo_transporte, chart_transportadora,
    chart_cobertura_docs, chart_top_materiais, tabela_criticos,
    PALETTE_MAIN, PALETTE_COOL, _apply_layout, AXIS_STYLE, HOJE,
)

EXCEL  = "Controle Geral_2026_RECAP.xlsm"
SHEETS = {"PTM": "PTM", "EMERGÊNCIA": "EMERGÊNCIA", "POWER BI": "POWER BI"}

# ── Config ────
st.set_page_config(page_title="Painel PTM", page_icon="📦",
                   layout="wide", initial_sidebar_state="expanded")

# ── CSS Dark Premium ────
st.markdown("""
<style>
.stApp { background: linear-gradient(135deg, #0D1B2A 0%, #1B2A4A 50%, #0D1B2A 100%); }
[data-testid="stSidebar"] { background: linear-gradient(180deg, #0A1628 0%, #162040 100%); border-right: 1px solid rgba(33,150,243,0.2); }
[data-testid="stSidebar"] * { color: #C8D8E8 !important; }
div[data-testid="metric-container"] {
    background: linear-gradient(135deg, rgba(27,79,114,0.6) 0%, rgba(21,67,96,0.4) 100%);
    border: 1px solid rgba(33,150,243,0.3); border-radius: 14px;
    padding: 14px 18px;
    box-shadow: 0 4px 20px rgba(0,0,0,0.4), inset 0 1px 0 rgba(255,255,255,0.05);
    backdrop-filter: blur(10px); transition: transform 0.2s;
}
div[data-testid="metric-container"]:hover { transform: translateY(-2px); box-shadow: 0 8px 30px rgba(33,150,243,0.25); }
[data-testid="stMetricValue"] { font-size: 1.75rem !important; font-weight: 800 !important; color: #FFFFFF !important; }
[data-testid="stMetricLabel"] { font-size: 0.78rem !important; color: #90B4CE !important; font-weight: 500 !important; }
[data-testid="stMetricDelta"] { font-size: 0.82rem !important; }
h1 { color: #FFFFFF !important; font-weight: 800 !important; letter-spacing: -0.5px; }
h2, h3 { color: #C8D8E8 !important; font-weight: 700 !important; }
[data-testid="stPlotlyChart"] {
    background: linear-gradient(135deg, rgba(13,27,42,0.8) 0%, rgba(21,40,70,0.6) 100%);
    border: 1px solid rgba(33,150,243,0.15); border-radius: 14px;
    padding: 8px; box-shadow: 0 4px 20px rgba(0,0,0,0.3);
}
[data-testid="stDataFrame"] { border: 1px solid rgba(33,150,243,0.2); border-radius: 10px; overflow: hidden; }
.stDownloadButton > button {
    background: linear-gradient(135deg, #1B4F72, #2E86C1) !important;
    color: white !important; border: none !important; border-radius: 8px !important;
    font-weight: 600 !important; box-shadow: 0 4px 15px rgba(33,150,243,0.3) !important;
}
.stTextInput > div > div > input {
    background: rgba(255,255,255,0.05) !important;
    border: 1px solid rgba(33,150,243,0.3) !important;
    color: white !important; border-radius: 8px !important;
}
hr { border-color: rgba(33,150,243,0.2) !important; }
.block-container { padding-top: 1rem !important; padding-bottom: 2rem !important; }
[data-testid="stRadio"] label { color: #C8D8E8 !important; font-weight: 500; }
</style>
""", unsafe_allow_html=True)


# ── Logo ────
def _show_logo():
    logo_path = Path("JSLOGO.jpg")
    if logo_path.exists():
        with open(logo_path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode()
        st.sidebar.markdown(
            f'<div style="text-align:center;padding:12px 0 8px 0;">'
            f'<img src="data:image/jpeg;base64,{b64}" style="max-width:160px;border-radius:10px;'
            f'box-shadow:0 4px 15px rgba(33,150,243,0.3);" /></div>',
            unsafe_allow_html=True,
        )
    st.sidebar.markdown(
        '<div style="text-align:center;font-size:0.7rem;color:#5A7A9A;padding-bottom:8px;">'
        'Painel de Controle PTM</div>', unsafe_allow_html=True,
    )


# ── Cache ────
@st.cache_data(show_spinner="⏳ Carregando dados...")
def get_data(sheet: str) -> pd.DataFrame:
    df = load_sheet(EXCEL, sheet)
    if sheet.strip().upper() == "PTM":
        return standardize_ptm(df)
    return standardize_generic(df)


# ── Sidebar Filtros PTM ────
def sidebar_filters_ptm(df: pd.DataFrame) -> pd.DataFrame:
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🔍 Filtros — PTM")
    dff = df.copy()

    if "DT_REMESSA" in dff.columns:
        valid = dff["DT_REMESSA"].dropna()
        if not valid.empty:
            min_d, max_d = valid.min().date(), valid.max().date()
            periodo = st.sidebar.date_input("📅 Período (Dt Remessa)", value=(min_d, max_d))
            if len(periodo) == 2:
                dff = dff[
                    (dff["DT_REMESSA"].dt.date >= periodo[0]) &
                    (dff["DT_REMESSA"].dt.date <= periodo[1])
                ]

    filtros = [
        ("STATUS",          "Status"),
        ("FASE_POWER_BI",   "Fase (Power BI)"),
        ("PRIORIZACAO",     "Priorização"),
        ("LOCAL_RECEBEDOR", "Local Recebedor"),
        ("TIPO_TRANSPORTE", "Tipo Transporte"),
        ("TRANSPORTADORA",  "Transportadora"),
        ("CENTRO",          "Centro"),
    ]
    for col, label in filtros:
        if col in dff.columns and isinstance(dff[col], pd.Series):
            opts = sorted(dff[col].astype(str).str.strip().replace("", pd.NA).dropna().unique().tolist())
            if opts:
                sel = st.sidebar.multiselect(label, opts, default=opts, key=f"f_{col}")
                if sel:
                    dff = dff[dff[col].isin(sel)]

    if "PENDENTE" in dff.columns:
        if st.sidebar.toggle("⚠️ Somente Pendentes", value=False):
            dff = dff[dff["PENDENTE"] == True]

    st.sidebar.markdown(f"---\n**{len(dff):,}** registros filtrados")
    return dff


# ── Gráfico helper ────
def _chart(fig):
    if fig:
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


# ── KPI Row — linha 1: totais gerais ────
def kpi_row_ptm(df: pd.DataFrame):
    total     = len(df)
    pendentes = int(df["PENDENTE"].sum())      if "PENDENTE"      in df.columns else 0
    baixados  = int((~df["PENDENTE"]).sum())   if "PENDENTE"      in df.columns else 0
    atrasados = int(df["FLAG_ATRASADO"].sum()) if "FLAG_ATRASADO" in df.columns else 0
    emerg     = int(df["IS_EMERGENCIA"].sum()) if "IS_EMERGENCIA" in df.columns else 0
    dias_med  = round(df["DIAS_ABERTO"].mean(), 1) if "DIAS_ABERTO" in df.columns and df["DIAS_ABERTO"].notna().any() else 0
    valor_nf  = df["VALOR_NF"].sum()           if "VALOR_NF"      in df.columns else 0
    pct_pend  = round(pendentes / total * 100, 1) if total > 0 else 0

    c1,c2,c3,c4,c5,c6,c7 = st.columns(7)
    c1.metric("📋 Total Remessas",    f"{total:,}")
    c2.metric("⏳ Pendentes",          f"{pendentes:,}", f"{pct_pend}% do total")
    c3.metric("✅ Baixados/Finalizados",f"{baixados:,}", f"{100-pct_pend}% do total")
    c4.metric("🔴 Atrasados",         f"{atrasados:,}", "vs. Dt Necessidade")
    c5.metric("🚨 Emergenciais",      f"{emerg:,}")
    c6.metric("⏱️ Dias Médio Aberto", f"{dias_med} dias")
    c7.metric("💰 Valor NF Total",    f"R$ {valor_nf:,.2f}")


# ── KPI Row — linha 2: status detalhado ────
def kpi_status_row(df: pd.DataFrame):
    if "STATUS" not in df.columns:
        return
    vc = df["STATUS"].value_counts()

    status_map = {
        "COLETADO":           ("🚛", "Coletados",        "#2196F3"),
        "FINALIZADO":         ("✅", "Finalizados",       "#1ABC9C"),
        "BAIXADA":            ("📥", "Baixadas",          "#27AE60"),
        "NF EMITIDA":         ("🧾", "NF Emitida",        "#F39C12"),
        "AGUARDANDO BAIXA":   ("⏳", "Aguard. Baixa",     "#E67E22"),
        "CANCELADO":          ("❌", "Cancelados",        "#E74C3C"),
        "ENCONTRADO / COM LEP":("📄","Com LEP",           "#9B59B6"),
        "COM CHAMADO":        ("📞", "Com Chamado",       "#3498DB"),
        "SOLICITADA NF":      ("📋", "Solicitada NF",     "#1A6B9A"),
        "NF IMPRESSA":        ("🖨️", "NF Impressa",       "#7D3C98"),
    }

    present = [(k, v) for k, v in status_map.items() if k in vc.index]
    if not present:
        return

    st.markdown("#### 📊 Status das Remessas")
    chunk = 5
    for i in range(0, len(present), chunk):
        batch = present[i:i+chunk]
        cols = st.columns(len(batch))
        for col, (status_key, (icon, label, color)) in zip(cols, batch):
            qtd = int(vc.get(status_key, 0))
            pct = round(qtd / len(df) * 100, 1) if len(df) > 0 else 0
            col.markdown(
                f"""<div style="background:linear-gradient(135deg,rgba(27,79,114,0.5),rgba(13,27,42,0.7));
                border:1px solid {color}55;border-radius:12px;padding:12px 14px;
                box-shadow:0 3px 12px rgba(0,0,0,0.3);text-align:center;">
                <div style="font-size:1.5rem">{icon}</div>
                <div style="font-size:1.4rem;font-weight:800;color:{color};margin:4px 0">{qtd:,}</div>
                <div style="font-size:0.72rem;color:#90B4CE;font-weight:500">{label}</div>
                <div style="font-size:0.68rem;color:#5A7A9A">{pct}%</div>
                </div>""",
                unsafe_allow_html=True,
            )


# ── KPI Row — linha 3: fases Power BI ────
def kpi_fase_row(df: pd.DataFrame):
    if "FASE_POWER_BI" not in df.columns:
        return
    vc = df["FASE_POWER_BI"].value_counts()

    fase_map = {
        "NÃO ENCONTRADO":     ("❓", "#E74C3C"),
        "[7] Aguard. Coleta": ("🚚", "#F39C12"),
        "[5] Aguard. NFe":    ("🧾", "#E67E22"),
        "[2] Aguard. DSM":    ("📋", "#9B59B6"),
        "[8] Em Transporte":  ("🛣️", "#2196F3"),
        "[9] Aguard. MIGO":   ("📦", "#1ABC9C"),
    }

    present = [(k, v) for k, v in fase_map.items() if k in vc.index]
    if not present:
        return

    st.markdown("#### 🔄 Fases (Power BI)")
    cols = st.columns(len(present))
    for col, (fase_key, (icon, color)) in zip(cols, present):
        qtd = int(vc.get(fase_key, 0))
        pct = round(qtd / len(df) * 100, 1) if len(df) > 0 else 0
        label = fase_key.replace("[","").replace("]","").strip()
        col.markdown(
            f"""<div style="background:linear-gradient(135deg,rgba(27,79,114,0.4),rgba(13,27,42,0.6));
            border:1px solid {color}55;border-radius:12px;padding:10px 12px;
            box-shadow:0 3px 12px rgba(0,0,0,0.3);text-align:center;">
            <div style="font-size:1.3rem">{icon}</div>
            <div style="font-size:1.35rem;font-weight:800;color:{color};margin:3px 0">{qtd:,}</div>
            <div style="font-size:0.7rem;color:#90B4CE;font-weight:500">{label}</div>
            <div style="font-size:0.65rem;color:#5A7A9A">{pct}%</div>
            </div>""",
            unsafe_allow_html=True,
        )


# ── Gráfico: Status donut ────
def chart_status_donut(df):
    if "STATUS" not in df.columns:
        return None
    agg = df["STATUS"].value_counts().reset_index()
    agg.columns = ["Status", "Qtd"]
    colors = ["#2196F3","#1ABC9C","#27AE60","#F39C12","#E67E22","#E74C3C","#9B59B6","#3498DB","#1A6B9A","#7D3C98"]
    fig = go.Figure(go.Pie(
        labels=agg["Status"], values=agg["Qtd"], hole=0.52,
        marker=dict(colors=colors[:len(agg)], line=dict(color="#0D1B2A", width=2)),
        textinfo="percent+label", textfont=dict(size=10, color="white"),
        pull=[0.04 if i == 0 else 0 for i in range(len(agg))],
    ))
    _apply_layout(fig, "📊 Distribuição por Status")
    fig.update_layout(legend=dict(orientation="v", x=1.02, font=dict(size=9)))
    return fig


# ── Gráfico: Fase Power BI barras ────
def chart_fase_bar(df):
    if "FASE_POWER_BI" not in df.columns:
        return None
    agg = df[df["FASE_POWER_BI"].str.len() > 0]["FASE_POWER_BI"].value_counts().reset_index()
    agg.columns = ["Fase", "Qtd"]
    if "FASE_NUM" in df.columns:
        fn = df[["FASE_POWER_BI","FASE_NUM"]].drop_duplicates()
        agg = agg.merge(fn, left_on="Fase", right_on="FASE_POWER_BI", how="left").sort_values("FASE_NUM")
    n = len(agg)
    colors = px.colors.sample_colorscale("Blues", [i/max(n-1,1) for i in range(n)])
    fig = go.Figure(go.Bar(
        x=agg["Qtd"], y=agg["Fase"], orientation="h",
        marker=dict(color=colors, line=dict(color="rgba(255,255,255,0.15)", width=0.5)),
        text=agg["Qtd"], textposition="outside",
        textfont=dict(color="white", size=11),
    ))
    _apply_layout(fig, "🔄 Itens por Fase (Power BI)")
    fig.update_layout(yaxis=dict(categoryorder="total ascending", **AXIS_STYLE))
    return fig


# ── Gráfico: Priorização ────
def chart_priorizacao(df):
    if "PRIORIZACAO" not in df.columns:
        return None
    agg = df[df["PRIORIZACAO"].str.len() > 0]["PRIORIZACAO"].value_counts().reset_index()
    agg.columns = ["Priorização", "Qtd"]
    cmap = {"Emergência":"#E74C3C","Expresso":"#F39C12","Econômico":"#1ABC9C",
            "EMERGENCIAL":"#E74C3C","EXPRESSO":"#F39C12","ECONÔMICO":"#1ABC9C"}
    colors = [cmap.get(p, "#2E86C1") for p in agg["Priorização"]]
    fig = go.Figure(go.Bar(
        x=agg["Priorização"], y=agg["Qtd"],
        marker=dict(color=colors, line=dict(color="rgba(255,255,255,0.2)", width=0.5)),
        text=agg["Qtd"], textposition="outside",
        textfont=dict(color="white", size=12),
    ))
    _apply_layout(fig, "🚦 Priorização de Transporte")
    fig.update_layout(showlegend=False)
    return fig


# ── Gráfico: Destino top 15 ────
def chart_destino(df):
    if "LOCAL_RECEBEDOR" not in df.columns:
        return None
    agg = df[df["LOCAL_RECEBEDOR"].str.len() > 0]["LOCAL_RECEBEDOR"].value_counts().head(15).reset_index()
    agg.columns = ["Destino", "Qtd"]
    n = len(agg)
    colors = px.colors.sample_colorscale("Teal", [i/max(n-1,1) for i in range(n)])
    fig = go.Figure(go.Bar(
        x=agg["Qtd"], y=agg["Destino"], orientation="h",
        marker=dict(color=colors, line=dict(color="rgba(255,255,255,0.1)", width=0.5)),
        text=agg["Qtd"], textposition="outside",
        textfont=dict(color="white", size=11),
    ))
    _apply_layout(fig, "📍 Top 15 — Local Recebedor")
    fig.update_layout(yaxis=dict(categoryorder="total ascending", **AXIS_STYLE))
    return fig


# ── Gráfico: Cobrança PTM ────
def chart_cobranca(df):
    if "COBRANCA_PTM" not in df.columns:
        return None
    d = df[df["COBRANCA_PTM"].str.strip().str.len() > 0].copy()
    if d.empty:
        return None
    d["COBRANCA_CURTA"] = d["COBRANCA_PTM"].str[:50]
    agg = d["COBRANCA_CURTA"].value_counts().head(10).reset_index()
    agg.columns = ["Cobrança", "Qtd"]
    n = len(agg)
    colors = px.colors.sample_colorscale("Reds", [0.3 + 0.7*i/max(n-1,1) for i in range(n)])
    fig = go.Figure(go.Bar(
        x=agg["Qtd"], y=agg["Cobrança"], orientation="h",
        marker=dict(color=colors, line=dict(color="rgba(255,255,255,0.1)", width=0.5)),
        text=agg["Qtd"], textposition="outside",
        textfont=dict(color="white", size=10),
    ))
    _apply_layout(fig, "⚠️ Pendências de Cobrança PTM")
    fig.update_layout(yaxis=dict(categoryorder="total ascending", **AXIS_STYLE))
    return fig


# ── Gráfico: Pendentes vs Baixados por Destino ────
def chart_pendente_destino_local(df):
    if "LOCAL_RECEBEDOR" not in df.columns or "PENDENTE" not in df.columns:
        return None
    top15 = df[df["LOCAL_RECEBEDOR"].str.len() > 0]["LOCAL_RECEBEDOR"].value_counts().head(15).index
    d = df[df["LOCAL_RECEBEDOR"].isin(top15)].copy()
    agg = d.groupby(["LOCAL_RECEBEDOR","PENDENTE"]).size().reset_index(name="Qtd")
    agg["Situação"] = agg["PENDENTE"].map({True:"Pendente", False:"Baixado"})
    fig = go.Figure()
    for sit, color in [("Pendente","#E74C3C"),("Baixado","#1ABC9C")]:
        sub = agg[agg["Situação"]==sit]
        fig.add_trace(go.Bar(
            name=sit, x=sub["LOCAL_RECEBEDOR"], y=sub["Qtd"],
            marker=dict(color=color, line=dict(color="rgba(255,255,255,0.15)", width=0.5)),
            text=sub["Qtd"], textposition="inside",
            textfont=dict(color="white", size=10),
        ))
    _apply_layout(fig, "📦 Pendentes vs Baixados por Destino")
    fig.update_layout(barmode="stack", xaxis_tickangle=-35)
    return fig


# ── Gráfico: Status por Destino (heatmap) ────
def chart_status_destino(df):
    if "LOCAL_RECEBEDOR" not in df.columns or "STATUS" not in df.columns:
        return None
    top10_dest = df[df["LOCAL_RECEBEDOR"].str.len() > 0]["LOCAL_RECEBEDOR"].value_counts().head(10).index
    d = df[df["LOCAL_RECEBEDOR"].isin(top10_dest)].copy()
    pivot = d.groupby(["LOCAL_RECEBEDOR","STATUS"]).size().unstack(fill_value=0)
    fig = go.Figure(go.Heatmap(
        z=pivot.values,
        x=pivot.columns.tolist(),
        y=pivot.index.tolist(),
        colorscale="Blues",
        text=pivot.values,
        texttemplate="%{text}",
        textfont=dict(size=10, color="white"),
        hoverongaps=False,
    ))
    _apply_layout(fig, "🗺️ Status por Destino (Top 10)", height=400)
    fig.update_layout(
        xaxis=dict(tickangle=-30, **AXIS_STYLE),
        yaxis=dict(**AXIS_STYLE),
    )
    return fig


# ── Página PTM ────
def page_ptm(df: pd.DataFrame):
    st.markdown(
        f'<h1 style="margin-bottom:0">📦 PTM — Painel de Controle</h1>'
        f'<p style="color:#6A8FAF;font-size:0.85rem;margin-top:4px;">'
        f'Atualizado: <b>{pd.Timestamp.now().strftime("%d/%m/%Y %H:%M")}</b>'
        f' &nbsp;|&nbsp; Base: <b>{len(df):,} remessas</b></p>',
        unsafe_allow_html=True,
    )
    st.divider()

    kpi_row_ptm(df)
    st.markdown("<br>", unsafe_allow_html=True)

    kpi_status_row(df)
    st.markdown("<br>", unsafe_allow_html=True)

    kpi_fase_row(df)
    st.divider()

    c1, c2 = st.columns(2)
    with c1: _chart(chart_status_donut(df))
    with c2: _chart(chart_fase_bar(df))

    c1, c2 = st.columns(2)
    with c1: _chart(chart_atraso(df))
    with c2: _chart(chart_pendente_destino_local(df))

    c1, c2 = st.columns(2)
    with c1: _chart(chart_priorizacao(df))
    with c2: _chart(chart_tipo_transporte(df))

    c1, c2 = st.columns(2)
    with c1: _chart(chart_timeline_remessa(df, "W"))
    with c2: _chart(chart_timeline_remessa(df, "ME"))

    c1, c2 = st.columns(2)
    with c1: _chart(chart_destino(df))
    with c2: _chart(chart_valor_destino(df))

    c1, c2 = st.columns(2)
    with c1: _chart(chart_cobertura_docs(df))
    with c2: _chart(chart_transportadora(df))

    c1, c2 = st.columns(2)
    with c1: _chart(chart_faixa_aberto(df))
    with c2: _chart(chart_hist_aberto(df))

    c1, c2 = st.columns(2)
    with c1: _chart(chart_cobranca(df))
    with c2: _chart(chart_status_destino(df))

    _chart(chart_top_materiais(df))

    st.divider()
    st.markdown("### 🔴 Top 30 — Pendentes Mais Antigos")
    top = tabela_criticos(df, n=30)
    if top is not None and not top.empty:
        st.dataframe(top, use_container_width=True, height=480)

    st.divider()
    st.markdown("### 📋 Detalhamento Completo")
    busca = st.text_input("🔎 Buscar por Remessa, Pedido, NM, Denominação, LEP, DSM, NF, Destino...")
    df_show = df.copy()
    if busca:
        mask = df_show.astype(str).apply(
            lambda col: col.str.contains(busca, case=False, na=False)
        ).any(axis=1)
        df_show = df_show[mask]

    st.caption(f"{len(df_show):,} registros exibidos")
    st.dataframe(df_show, use_container_width=True, height=550)

    try:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df_show.to_excel(writer, index=False, sheet_name="PTM")
        buf.seek(0)
        st.download_button("⬇️ Exportar .xlsx", data=buf,
            file_name="ptm_filtrado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception:
        csv = df_show.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig")
        st.download_button("⬇️ Exportar CSV", data=csv, file_name="ptm_filtrado.csv", mime="text/csv")

    # ── Rodapé animado ────
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown("""
<style>
@keyframes blink-random {
    0%   { opacity: 1; }
    15%  { opacity: 0.15; }
    30%  { opacity: 0.9; }
    47%  { opacity: 0.05; }
    55%  { opacity: 1; }
    68%  { opacity: 0.3; }
    80%  { opacity: 0.85; }
    91%  { opacity: 0.1; }
    100% { opacity: 1; }
}
.footer-blink {
    animation: blink-random 4.7s ease-in-out infinite;
    text-align: center;
    font-size: 0.68rem;
    color: #3A5A7A;
    letter-spacing: 0.08em;
    font-weight: 400;
    padding: 6px 0 10px 0;
    user-select: none;
}
.footer-sep {
    color: #1E3A5A;
    margin: 0 8px;
}
</style>
<div class="footer-blink">
    WebDesign&nbsp;<span style="color:#2E5A7A">Djalma A Barbosa</span>
    <span class="footer-sep">·</span>
    B.Dados&nbsp;<span style="color:#2E5A7A">Thiago Aniz</span>
    <span class="footer-sep">·</span>
    <span style="color:#1E4A6A;letter-spacing:0.12em">RECAP — PETROBRAS — 2026</span>
</div>
""", unsafe_allow_html=True)


# ── Página EMERGÊNCIA ────
def page_emergencia(df: pd.DataFrame):
    st.markdown('<h1>🚨 EMERGÊNCIA — Monitoramento</h1>', unsafe_allow_html=True)
    st.divider()

    total = len(df)
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("📋 Total Registros", f"{total:,}")

    date_col = None
    for c in df.columns:
        if df[c].dtype == "datetime64[ns]" and df[c].notna().any():
            date_col = c
            break

    if date_col:
        recentes = int((df[date_col].dt.year >= 2025).sum())
        c2.metric("📅 Registros 2025+", f"{recentes:,}")

    status_col = None
    for c in df.columns:
        if "status" in c.lower() or "situac" in c.lower():
            status_col = c
            break

    if status_col:
        pendentes = int(df[status_col].astype(str).str.upper().str.contains("PEND|ABERTO|OPEN", na=False).sum())
        c3.metric("⏳ Pendentes", f"{pendentes:,}")

    st.divider()

    col_plots = [c for c in df.columns if df[c].dtype == object and df[c].notna().any()]

    if len(col_plots) >= 2:
        c1, c2 = st.columns(2)
        for i, col in enumerate(col_plots[:4]):
            agg = df[col].astype(str).str.strip().value_counts().head(15).reset_index()
            agg.columns = [col, "Qtd"]
            fig = go.Figure(go.Bar(
                x=agg["Qtd"], y=agg[col], orientation="h",
                marker=dict(color=PALETTE_MAIN[i % len(PALETTE_MAIN)],
                            line=dict(color="rgba(255,255,255,0.1)", width=0.5)),
                text=agg["Qtd"], textposition="outside",
                textfont=dict(color="white", size=10),
            ))
            _apply_layout(fig, f"📊 {col}")
            fig.update_layout(yaxis=dict(categoryorder="total ascending"))
            target = c1 if i % 2 == 0 else c2
            with target:
                st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    if date_col:
        d = df.dropna(subset=[date_col]).copy()
        d = d[d[date_col].dt.year >= 2025]
        if not d.empty:
            ts = d.set_index(date_col).resample("W").size().reset_index(name="Qtd")
            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=ts[date_col], y=ts["Qtd"],
                mode="lines+markers",
                line=dict(color="#E74C3C", width=2.5, shape="spline"),
                marker=dict(size=6, color="#F1948A"),
                fill="tozeroy", fillcolor="rgba(231,76,60,0.15)",
            ))
            _apply_layout(fig, "📅 Timeline Semanal — EMERGÊNCIA")
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    st.divider()
    st.markdown("### 📋 Dados Completos")
    busca = st.text_input("🔎 Buscar...", key="busca_emerg")
    df_show = df.copy()
    if busca:
        mask = df_show.astype(str).apply(lambda col: col.str.contains(busca, case=False, na=False)).any(axis=1)
        df_show = df_show[mask]
    st.caption(f"{len(df_show):,} registros")
    st.dataframe(df_show, use_container_width=True, height=550)

    try:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df_show.to_excel(writer, index=False, sheet_name="EMERGÊNCIA")
        buf.seek(0)
        st.download_button("⬇️ Exportar .xlsx", data=buf,
            file_name="emergencia_filtrado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception:
        pass


# ── Página POWER BI ────
def page_powerbi(df: pd.DataFrame):
    st.markdown('<h1>📊 POWER BI — Visão Complementar</h1>', unsafe_allow_html=True)
    st.divider()

    total = len(df)
    date_cols = [c for c in df.columns if df[c].dtype == "datetime64[ns]" and df[c].notna().any()]

    for dc in date_cols:
        df = df[~((df[dc].notna()) & (df[dc].dt.year < 2020))]

    c1,c2,c3,c4 = st.columns(4)
    c1.metric("📋 Total Registros", f"{total:,}")

    if date_cols:
        dc = date_cols[0]
        ano_atual = int(df[dc].dt.year.mode()[0]) if df[dc].notna().any() else 2026
        recentes = int((df[dc].dt.year == ano_atual).sum())
        c2.metric(f"📅 Ano {ano_atual}", f"{recentes:,}")

    st.divider()

    cat_cols = [c for c in df.columns if df[c].dtype == object and df[c].notna().any() and df[c].nunique() <= 50]

    if cat_cols:
        pairs = [(cat_cols[i], cat_cols[i+1] if i+1 < len(cat_cols) else None)
                 for i in range(0, min(len(cat_cols), 8), 2)]
        for col_a, col_b in pairs:
            c1, c2 = st.columns(2)
            for col, target in [(col_a, c1), (col_b, c2)]:
                if col is None:
                    continue
                agg = df[col].astype(str).str.strip().value_counts().head(15).reset_index()
                agg.columns = [col, "Qtd"]
                n = len(agg)
                colors = PALETTE_MAIN[:n]
                fig = go.Figure(go.Bar(
                    x=agg["Qtd"], y=agg[col], orientation="h",
                    marker=dict(color=colors, line=dict(color="rgba(255,255,255,0.1)", width=0.5)),
                    text=agg["Qtd"], textposition="outside",
                    textfont=dict(color="white", size=10),
                ))
                _apply_layout(fig, f"📊 {col}")
                fig.update_layout(yaxis=dict(categoryorder="total ascending"))
                with target:
                    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    for dc in date_cols[:2]:
        d = df.dropna(subset=[dc]).copy()
        d = d[d[dc].dt.year >= 2025]
        if d.empty:
            continue
        c1, c2 = st.columns(2)
        for freq, label, target in [("W","Semanal",c1),("ME","Mensal",c2)]:
            ts = d.set_index(dc).resample(freq).size().reset_index(name="Qtd")
            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=ts[dc], y=ts["Qtd"],
                mode="lines+markers",
                line=dict(color="#2196F3", width=2.5, shape="spline"),
                marker=dict(size=6, color="#64B5F6"),
                fill="tozeroy", fillcolor="rgba(33,150,243,0.15)",
            ))
            _apply_layout(fig, f"📅 {dc} — {label}")
            with target:
                st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    st.divider()
    st.markdown("### 📋 Dados Completos")
    busca = st.text_input("🔎 Buscar...", key="busca_pbi")
    df_show = df.copy()
    if busca:
        mask = df_show.astype(str).apply(lambda col: col.str.contains(busca, case=False, na=False)).any(axis=1)
        df_show = df_show[mask]
    st.caption(f"{len(df_show):,} registros")
    st.dataframe(df_show, use_container_width=True, height=550)

    try:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df_show.to_excel(writer, index=False, sheet_name="POWER BI")
        buf.seek(0)
        st.download_button("⬇️ Exportar .xlsx", data=buf,
            file_name="powerbi_filtrado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception:
        pass


# ── Main ────
def main():
    _show_logo()
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🗂️ Navegação")
    page = st.sidebar.radio("", list(SHEETS.keys()), index=0, label_visibility="collapsed")

    df_raw = get_data(SHEETS[page])

    if page == "PTM":
        df_f = sidebar_filters_ptm(df_raw)
        if df_f.empty:
            st.warning("⚠️ Nenhum dado com os filtros aplicados.")
            return
        page_ptm(df_f)
    elif page == "EMERGÊNCIA":
        page_emergencia(df_raw)
    else:
        page_powerbi(df_raw)


if __name__ == "__main__":
    main()
