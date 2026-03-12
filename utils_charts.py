import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

HOJE = pd.Timestamp.now().normalize()

# ── Paleta sofisticada ────
PALETTE_MAIN   = ["#1B4F72","#2E86C1","#1ABC9C","#F39C12","#E74C3C","#8E44AD","#2ECC71","#E67E22","#3498DB","#C0392B"]
PALETTE_COOL   = ["#0D3B66","#1A6B9A","#2196F3","#64B5F6","#90CAF9","#BBDEFB"]
PALETTE_WARM   = ["#7B1818","#C0392B","#E74C3C","#F1948A","#FADBD8"]
PALETTE_GREEN  = ["#0B5345","#117A65","#1ABC9C","#76D7C4","#A2D9CE"]
PALETTE_PURPLE = ["#4A235A","#7D3C98","#9B59B6","#C39BD3","#E8DAEF"]

LAYOUT_BASE = dict(
    paper_bgcolor="rgba(15,20,40,0.0)",
    plot_bgcolor="rgba(15,20,40,0.0)",
    font=dict(family="Segoe UI, Arial", size=12, color="#E0E6F0"),
    title_font=dict(size=15, color="#FFFFFF", family="Segoe UI Semibold"),
    legend=dict(bgcolor="rgba(255,255,255,0.05)", bordercolor="rgba(255,255,255,0.1)", borderwidth=1),
    margin=dict(l=10, r=10, t=50, b=10),
    hoverlabel=dict(bgcolor="#1B2A4A", font_size=12, font_color="white"),
)

AXIS_STYLE = dict(
    gridcolor="rgba(255,255,255,0.07)",
    zerolinecolor="rgba(255,255,255,0.1)",
    tickfont=dict(color="#A0B0C8"),
    title_font=dict(color="#C0D0E0"),
)


def _apply_layout(fig, title="", height=380):
    fig.update_layout(**LAYOUT_BASE, title=title, height=height)
    fig.update_xaxes(**AXIS_STYLE)
    fig.update_yaxes(**AXIS_STYLE)
    return fig


def _safe_col(df, col):
    return col in df.columns and df[col].notna().any()


# ── 1. Status (donut 3D-style) ────
def chart_status(df):
    if not _safe_col(df, "STATUS"):
        return None
    agg = df["STATUS"].value_counts().reset_index()
    agg.columns = ["Status", "Qtd"]
    fig = go.Figure(go.Pie(
        labels=agg["Status"], values=agg["Qtd"],
        hole=0.5,
        marker=dict(colors=PALETTE_MAIN, line=dict(color="#0D1B2A", width=2)),
        textinfo="percent+label",
        textfont=dict(size=11, color="white"),
        pull=[0.04]*len(agg),
    ))
    _apply_layout(fig, "📊 Distribuição por Status")
    fig.update_layout(legend=dict(orientation="v", x=1.02, font=dict(size=10)))
    return fig


# ── 2. Fase Power BI (barras horizontais 3D) ────
def chart_fase_powerbi(df):
    if not _safe_col(df, "FASE_POWER_BI"):
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


# ── 3. Priorização (barras verticais) ────
def chart_priorizacao(df):
    if not _safe_col(df, "PRIORIZACAO"):
        return None
    agg = df[df["PRIORIZACAO"].str.len() > 0]["PRIORIZACAO"].value_counts().reset_index()
    agg.columns = ["Priorização", "Qtd"]
    cmap = {"Emergência":"#E74C3C","Expresso":"#F39C12","Econômico":"#1ABC9C","EMERGENCIAL":"#E74C3C","EXPRESSO":"#F39C12","ECONÔMICO":"#1ABC9C"}
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


# ── 4. Destino top 15 ────
def chart_destino(df):
    if not _safe_col(df, "LOCAL_RECEBEDOR"):
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


# ── 5. Timeline remessas (area com gradiente) ────
def chart_timeline_remessa(df, freq="W"):
    if not _safe_col(df, "DT_REMESSA"):
        return None
    d = df.dropna(subset=["DT_REMESSA"]).copy()
    # Filtrar apenas 2026
    d = d[d["DT_REMESSA"].dt.year >= 2025]
    if d.empty:
        return None
    ts = d.set_index("DT_REMESSA").resample(freq).size().reset_index(name="Remessas")
    label = "Semanal" if freq == "W" else "Mensal"
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=ts["DT_REMESSA"], y=ts["Remessas"],
        mode="lines+markers",
        line=dict(color="#2196F3", width=2.5, shape="spline"),
        marker=dict(size=6, color="#64B5F6", line=dict(color="#0D3B66", width=1)),
        fill="tozeroy",
        fillcolor="rgba(33,150,243,0.15)",
        name="Remessas",
    ))
    _apply_layout(fig, f"📅 Volume de Remessas — {label}")
    fig.update_layout(xaxis_title="Período", yaxis_title="Qtd")
    return fig


# ── 6. Faixa de dias em aberto ────
def chart_faixa_aberto(df):
    if not _safe_col(df, "FAIXA_ABERTO"):
        return None
    ordem = ["Futuro","0–7 dias","8–15 dias","16–30 dias","31–60 dias","61–90 dias","> 90 dias","Sem data"]
    agg = df["FAIXA_ABERTO"].value_counts().reindex(ordem).dropna().reset_index()
    agg.columns = ["Faixa", "Qtd"]
    colors = ["#1ABC9C","#2ECC71","#F39C12","#E67E22","#E74C3C","#C0392B","#7B1818","#7F8C8D"][:len(agg)]
    fig = go.Figure(go.Bar(
        x=agg["Faixa"], y=agg["Qtd"],
        marker=dict(color=colors, line=dict(color="rgba(255,255,255,0.15)", width=0.5)),
        text=agg["Qtd"], textposition="outside",
        textfont=dict(color="white", size=11),
    ))
    _apply_layout(fig, "⏱️ Itens por Faixa de Dias em Aberto")
    return fig


# ── 7. Histograma dias em aberto ────
def chart_hist_aberto(df):
    if not _safe_col(df, "DIAS_ABERTO"):
        return None
    fig = go.Figure(go.Histogram(
        x=df["DIAS_ABERTO"].dropna(), nbinsx=40,
        marker=dict(color="#2196F3", line=dict(color="rgba(255,255,255,0.2)", width=0.5)),
        opacity=0.85,
    ))
    _apply_layout(fig, "📊 Histograma — Dias em Aberto")
    fig.update_layout(xaxis_title="Dias", yaxis_title="Qtd")
    return fig


# ── 8. Atraso vs Necessidade ────
def chart_atraso(df):
    if not _safe_col(df, "ATRASO_DIAS"):
        return None
    d = df.dropna(subset=["ATRASO_DIAS"]).copy()
    d["Situação"] = d["ATRASO_DIAS"].apply(
        lambda x: "🔴 Atrasado" if x > 0 else ("🟡 No Prazo" if x == 0 else "🟢 Adiantado")
    )
    agg = d["Situação"].value_counts().reset_index()
    agg.columns = ["Situação", "Qtd"]
    cmap = {"🔴 Atrasado":"#E74C3C","🟡 No Prazo":"#F39C12","🟢 Adiantado":"#1ABC9C"}
    colors = [cmap.get(s, "#2E86C1") for s in agg["Situação"]]
    fig = go.Figure(go.Bar(
        x=agg["Situação"], y=agg["Qtd"],
        marker=dict(color=colors, line=dict(color="rgba(255,255,255,0.2)", width=0.5)),
        text=agg["Qtd"], textposition="outside",
        textfont=dict(color="white", size=13),
    ))
    _apply_layout(fig, "⚠️ Situação vs. Data de Necessidade")
    fig.update_layout(showlegend=False)
    return fig


# ── 9. Valor NF por Destino ────
def chart_valor_destino(df):
    if not _safe_col(df, "LOCAL_RECEBEDOR") or not _safe_col(df, "VALOR_NF"):
        return None
    agg = (
        df[df["LOCAL_RECEBEDOR"].str.len() > 0]
        .groupby("LOCAL_RECEBEDOR")["VALOR_NF"].sum()
        .reset_index().sort_values("VALOR_NF", ascending=False).head(15)
    )
    n = len(agg)
    colors = px.colors.sample_colorscale("Oranges", [i/max(n-1,1) for i in range(n)])
    fig = go.Figure(go.Bar(
        x=agg["VALOR_NF"], y=agg["LOCAL_RECEBEDOR"], orientation="h",
        marker=dict(color=colors, line=dict(color="rgba(255,255,255,0.1)", width=0.5)),
        text=agg["VALOR_NF"].apply(lambda v: f"R$ {v:,.0f}"),
        textposition="outside", textfont=dict(color="white", size=10),
    ))
    _apply_layout(fig, "💰 Top 15 Destinos — Valor NF (R$)")
    fig.update_layout(yaxis=dict(categoryorder="total ascending", **AXIS_STYLE))
    return fig


# ── 10. Pendentes vs Baixados por Destino ────
def chart_pendente_destino(df):
    if not _safe_col(df, "LOCAL_RECEBEDOR") or "PENDENTE" not in df.columns:
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


# ── 11. Tipo de Transporte (donut) ────
def chart_tipo_transporte(df):
    if not _safe_col(df, "TIPO_TRANSPORTE"):
        return None
    agg = df[df["TIPO_TRANSPORTE"].str.len() > 0]["TIPO_TRANSPORTE"].value_counts().reset_index()
    agg.columns = ["Tipo", "Qtd"]
    cmap = {"EMERGENCIAL":"#E74C3C","EXPRESSO":"#F39C12","ECONÔMICO":"#1ABC9C","Emergência":"#E74C3C","Expresso":"#F39C12","Econômico":"#1ABC9C"}
    colors = [cmap.get(t, "#2E86C1") for t in agg["Tipo"]]
    fig = go.Figure(go.Pie(
        labels=agg["Tipo"], values=agg["Qtd"], hole=0.5,
        marker=dict(colors=colors, line=dict(color="#0D1B2A", width=2)),
        textinfo="percent+label", textfont=dict(size=11, color="white"),
        pull=[0.04]*len(agg),
    ))
    _apply_layout(fig, "🚚 Tipo de Transporte Utilizado")
    return fig


# ── 12. Transportadora ────
def chart_transportadora(df):
    if not _safe_col(df, "TRANSPORTADORA"):
        return None
    agg = df[df["TRANSPORTADORA"].str.len() > 0]["TRANSPORTADORA"].value_counts().reset_index()
    agg.columns = ["Transportadora", "Qtd"]
    n = len(agg)
    colors = px.colors.sample_colorscale("Viridis", [i/max(n-1,1) for i in range(n)])
    fig = go.Figure(go.Bar(
        x=agg["Transportadora"], y=agg["Qtd"],
        marker=dict(color=colors, line=dict(color="rgba(255,255,255,0.15)", width=0.5)),
        text=agg["Qtd"], textposition="outside",
        textfont=dict(color="white", size=11),
    ))
    _apply_layout(fig, "🏢 Remessas por Transportadora")
    fig.update_layout(showlegend=False, xaxis_tickangle=-30)
    return fig


# ── 13. Cobertura documental ────
def chart_cobertura_docs(df):
    flags = {"TEM_LEP":"LEP","TEM_NF":"NF","TEM_DSM":"DSM"}
    rows = []
    for flag, label in flags.items():
        if flag in df.columns:
            sim = int(df[flag].sum())
            nao = len(df) - sim
            rows.append({"Documento":label,"Com":sim,"Sem":nao})
    if not rows:
        return None
    agg = pd.DataFrame(rows)
    fig = go.Figure()
    fig.add_trace(go.Bar(name="Com Documento", x=agg["Documento"], y=agg["Com"],
        marker=dict(color="#1ABC9C", line=dict(color="rgba(255,255,255,0.15)", width=0.5)),
        text=agg["Com"], textposition="inside", textfont=dict(color="white")))
    fig.add_trace(go.Bar(name="Sem Documento", x=agg["Documento"], y=agg["Sem"],
        marker=dict(color="#E74C3C", line=dict(color="rgba(255,255,255,0.15)", width=0.5)),
        text=agg["Sem"], textposition="inside", textfont=dict(color="white")))
    _apply_layout(fig, "📋 Cobertura Documental (LEP / NF / DSM)")
    fig.update_layout(barmode="stack", legend=dict(orientation="h", y=-0.15))
    return fig


# ── 14. Top materiais ────
def chart_top_materiais(df):
    if not _safe_col(df, "DENOMINACAO"):
        return None
    agg = df[df["DENOMINACAO"].str.len() > 0]["DENOMINACAO"].value_counts().head(15).reset_index()
    agg.columns = ["Material", "Qtd"]
    n = len(agg)
    colors = px.colors.sample_colorscale("Purples", [i/max(n-1,1) for i in range(n)])
    fig = go.Figure(go.Bar(
        x=agg["Qtd"], y=agg["Material"], orientation="h",
        marker=dict(color=colors, line=dict(color="rgba(255,255,255,0.1)", width=0.5)),
        text=agg["Qtd"], textposition="outside",
        textfont=dict(color="white", size=11),
    ))
    _apply_layout(fig, "🔩 Top 15 — Materiais Mais Frequentes", height=420)
    fig.update_layout(yaxis=dict(categoryorder="total ascending", **AXIS_STYLE))
    return fig


# ── 15. Tabela críticos ────
def tabela_criticos(df, n=30):
    sort_col = "DIAS_ABERTO" if "DIAS_ABERTO" in df.columns else None
    cols = [c for c in [
        "REMESSA","PEDIDO","ITEM","NM","DENOMINACAO",
        "STATUS","FASE_POWER_BI","PRIORIZACAO",
        "DT_REMESSA","DT_NECESSIDADE","DT_BAIXA",
        "DIAS_ABERTO","ATRASO_DIAS","LEAD_TIME_DIAS",
        "LOCAL_RECEBEDOR","CENTRO",
        "DSM","SONOF","NF","LEP","RETORNO_LOEP",
        "TIPO_TRANSPORTE","TRANSPORTADORA",
        "COBRANCA_PTM","PENDENTE",
    ] if c in df.columns]
    if sort_col:
        pend = df[df["PENDENTE"]==True] if "PENDENTE" in df.columns else df
        return pend.sort_values(sort_col, ascending=False).head(n)[cols]
    return df.head(n)[cols]
