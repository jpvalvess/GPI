import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

# =========================
# Configuração
# =========================
st.set_page_config(page_title="Relatório de Metal GPI", layout="wide")
st.title("Relatório de Metal GPI")

path = r"relatorio_metal.xlsx"
df = pd.read_excel(path, header=0)

# =========================
# Limpeza básica
# =========================
df = df.dropna(how="all")
df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
df = df.fillna(0)
df.columns = df.columns.astype(str).str.strip()  # remove espaços invisíveis

# =========================
# Descobrir coluna de parafusos (robusto; mantém compatibilidade)
# =========================
possiveis_paraf = [
    "QUANT. PARAF.", "QUANT. PARAF", "QTDE. PARAF.", "QTDE. PARAF",
    "PARAFUSOS", "QTD PARAFUSOS", "QTD. PARAF.", "QTD PARAF."
]

def achar_coluna_parafuso(cols):
    for nome in possiveis_paraf:
        if nome in cols:
            return nome
    norm = {c: c.replace(".", "").replace(" ", "").upper() for c in cols}
    alvo_norms = [n.replace(".", "").replace(" ", "").upper() for n in possiveis_paraf]
    for c, cn in norm.items():
        if cn in alvo_norms:
            return c
    return None

col_paraf = achar_coluna_parafuso(df.columns)
if col_paraf is None:
    col_paraf = "QUANT. PARAF."
    df[col_paraf] = 0.0  # não será usado agora, mas evita erros futuros

# =========================
# Filtros
# =========================
n_bm_opts = sorted(pd.Series(df["Nº BM"]).astype(str).unique().tolist())
desc_opts = sorted(pd.Series(df["DESCRIÇÃO"]).astype(str).unique().tolist())

st.sidebar.header("Filtros")
bm_selecionados = st.sidebar.multiselect("Nº BM", options=n_bm_opts)
desc_selecionadas = st.sidebar.multiselect("Descrição", options=desc_opts)

df_filtrado = df.copy()
df_filtrado["Nº BM"] = df_filtrado["Nº BM"].astype(str)
if bm_selecionados:
    df_filtrado = df_filtrado[df_filtrado["Nº BM"].isin(bm_selecionados)]
if desc_selecionadas:
    df_filtrado = df_filtrado[df_filtrado["DESCRIÇÃO"].astype(str).isin(desc_selecionadas)]

# Tipos numéricos seguros
num_cols = [
    "QUANT. PEÇAS","QUANT. KG", col_paraf,
    "FAB.","DESM.","MONT.",
    "QTDE. FAB.","QTDE. DESM.","QTDE. MONT."
]
for c in num_cols:
    if c in df_filtrado.columns:
        df_filtrado[c] = pd.to_numeric(df_filtrado[c], errors="coerce").fillna(0.0)

# =========================
# Utilidades
# =========================
cores = {
    'Desmontada': '#EF553B',
    'Fabricada':  '#636EFA',
    'Implantada': '#00CC96'
}
def num_br(v, casas=2):
    try:
        fmt = f"{float(v):,.{casas}f}"
        return fmt.replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "-"

# =========================
# KPIs e Taxa de Implantação
# =========================
col1, col2, col3, col4 = st.columns(4)

kg_fab   = float(df_filtrado["FAB."].sum())   if "FAB." in df_filtrado.columns else 0.0
kg_des   = float(df_filtrado["DESM."].sum())  if "DESM." in df_filtrado.columns else 0.0
kg_mont  = float(df_filtrado["MONT."].sum())  if "MONT." in df_filtrado.columns else 0.0

qtd_fab  = float(df_filtrado["QTDE. FAB."].sum())   if "QTDE. FAB." in df_filtrado.columns else 0.0
qtd_des  = float(df_filtrado["QTDE. DESM."].sum())  if "QTDE. DESM." in df_filtrado.columns else 0.0
qtd_mont = float(df_filtrado["QTDE. MONT."].sum())  if "QTDE. MONT." in df_filtrado.columns else 0.0

taxa_implant_kg = (kg_mont / kg_des) if kg_des > 0 else 0.0

col1.metric("Desmontada (KG)", num_br(kg_des))
col2.metric("Fabricada (KG)",  num_br(kg_fab))
col3.metric("Implantada (KG)", num_br(kg_mont))
col4.metric("Taxa de Implantação (KG)", f"{taxa_implant_kg*100:,.1f}%".replace(",", "."))

col5, col6, col7 = st.columns(3)
col5.metric("Desmontada (UN)", num_br(qtd_des, 0))
col6.metric("Fabricada (UN)",  num_br(qtd_fab, 0))
col7.metric("Implantada (UN)", num_br(qtd_mont, 0))

# =========================
# Visão Geral
# =========================
st.markdown("### Visão Geral")

col_g1, col_g2 = st.columns(2)

df_mensal = df_filtrado.groupby("Nº BM", as_index=False).agg(
    Desmontada=("DESM.", "sum"),
    Fabricada=("FAB.", "sum"),
    Implantada=("MONT.", "sum")
).sort_values("Nº BM")

fig1 = px.line(
    df_mensal,
    x="Nº BM",
    y=["Desmontada", "Fabricada", "Implantada"],
    markers=True,
    title="Realizado por BM (KG)",
    color_discrete_map=cores,
    labels={"value": "Peso (KG)", "variable": "Categoria"}
)
fig1.update_layout(
    hovermode='x unified',
    legend=dict(title="", orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    xaxis_title="Nº BM",
    yaxis_title="Peso (KG)"
)
col_g1.plotly_chart(fig1, use_container_width=True)

df_acum = df_mensal.copy()
df_acum["Desmontada"] = df_acum["Desmontada"].cumsum()
df_acum["Fabricada"]  = df_acum["Fabricada"].cumsum()
df_acum["Implantada"] = df_acum["Implantada"].cumsum()

fig2 = px.line(
    df_acum,
    x="Nº BM",
    y=["Desmontada", "Fabricada", "Implantada"],
    markers=True,
    title="Realizado Acumulado (KG)",
    color_discrete_map=cores,
    labels={"value": "Peso (KG)", "variable": "Categoria"}
)
fig2.update_layout(
    hovermode='x unified',
    legend=dict(title="", orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    xaxis_title="Nº BM",
    yaxis_title="Peso Acumulado (KG)"
)
col_g2.plotly_chart(fig2, use_container_width=True)

# Taxa de implantação por BM
st.markdown("#### Taxa de Implantação por BM (Implantado / Desmontado)")
taxa_bm = df_mensal.copy()
taxa_bm["Taxa Implantação"] = np.where(
    taxa_bm["Desmontada"] > 0, taxa_bm["Implantada"] / taxa_bm["Desmontada"], np.nan
)
fig_taxa_bm = px.bar(
    taxa_bm, x="Nº BM", y="Taxa Implantação",
    color_discrete_sequence=[cores["Implantada"]],
    labels={"Taxa Implantação":"Taxa"}
)
fig_taxa_bm.update_layout(title="Taxa de Implantação por BM")
fig_taxa_bm.update_yaxes(tickformat=".0%")
st.plotly_chart(fig_taxa_bm, use_container_width=True)

# =========================
# Produtividade e Mix (apenas Top-10 por Implantação)
# =========================
st.markdown("### Produtividade e Mix")

top_desc = (df_filtrado.groupby("DESCRIÇÃO", as_index=False)["MONT."].sum()
            .sort_values("MONT.", ascending=False).head(10))

fig_top = px.bar(
    top_desc, x="MONT.", y="DESCRIÇÃO",
    orientation="h",
    color_discrete_sequence=[cores["Implantada"]],
    labels={"MONT.":"KG Implantado","DESCRIÇÃO":"Descrição"},
    title="Top-10 Descrições por Implantação (KG)"
)
fig_top.update_layout(yaxis_categoryorder="total ascending")
st.plotly_chart(fig_top, use_container_width=True)

# =========================
# Detalhamento
# =========================
st.markdown("### Detalhamento")
st.dataframe(df_filtrado, use_container_width=True)
