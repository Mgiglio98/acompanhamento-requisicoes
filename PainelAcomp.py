import pandas as pd
import streamlit as st
from pathlib import Path

# --- Configuração inicial ---
st.set_page_config(page_title="Acompanhamento de Requisições", page_icon="📋", layout="wide")

# --- Carregar base ---
DATA_PATH = "AcompReq.xlsx"

try:
    df = pd.read_excel(DATA_PATH)
except FileNotFoundError:
    st.error("⚠️ O arquivo 'AcompReq.xlsx' não foi encontrado na raiz do repositório.")
    st.stop()

# --- Filtrar semana atual ---
df['REQ_DATA'] = pd.to_datetime(df['REQ_DATA'])
semana_atual = df[df['REQ_DATA'].dt.isocalendar().week == pd.Timestamp.now().isocalendar().week]

# --- Tabela Principal agrupada por Requisição ---
agrupado = (
    semana_atual
    .groupby('REQ_CDG')
    .agg(
        EMPRD_DESC=('EMPRD_DESC', 'first'),
        EMPRD_UF=('EMPRD_UF', 'first'),
        REQ_DATA=('REQ_DATA', 'first'),
        QTD_INSUMOS=('INSUMO_DESC', 'count'),
        QTD_COMPRADOS=('OF_CDG', lambda x: x.notna().sum()))
    .reset_index())

agrupado['QTD_PENDENTE'] = agrupado['QTD_INSUMOS'] - agrupado['QTD_COMPRADOS']

agrupado['STATUS'] = agrupado['QTD_PENDENTE'].apply(
    lambda x: "✅ Todos comprados" if x == 0 else f"⏳ {x} pendente(s)")

agrupado = agrupado.sort_values(['REQ_DATA', 'QTD_PENDENTE'], ascending=[True, False])

# --- Painel Visual---
st.title("📋 Acompanhamento de Requisições — Semana Atual")
st.metric("Total Requisições", len(agrupado))
st.metric("Requisições com OF", (agrupado['tem_of'] == "✅ Já tem OF").sum())

st.subheader("📊 Resumo por Requisição")
st.dataframe(agrupado)

st.subheader("🔎 Requisições sem OF")
st.dataframe(semana_atual[semana_atual['OF_CDG'].isna()][
    ['REQ_CDG', 'INSUMO_DESC', 'QTD_PED']])
