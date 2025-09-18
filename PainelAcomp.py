import pandas as pd
import streamlit as st
from pathlib import Path

# --- Configuração inicial ---
st.set_page_config(page_title="Acompanhamento de Requisições", page_icon="📋", layout="wide")

# --- Carregar base ---
DATA_PATH = Path(__file__).resolve().parent.parent / "AcompReq.xlsx"

if not DATA_PATH.exists():
    st.error("⚠️ O arquivo de dados não foi encontrado em 'data/AcompReq.xlsx'")
    st.stop()

df = pd.read_excel(DATA_PATH)

# --- Filtrar semana atual ---
df['REQ_DATA'] = pd.to_datetime(df['REQ_DATA'])
semana_atual = df[df['REQ_DATA'].dt.isocalendar().week == pd.Timestamp.now().isocalendar().week]

# --- Agrupar por requisição e marcar se já tem OF ---
agrupado = (
    semana_atual
    .groupby('REQ_CDG')
    .agg({
        'EMPRD_DESC': 'first',
        'EMPRD_UF': 'first',
        'REQ_DATA': 'first',
        'OF_CDG': lambda x: x.notna().any(),
        'INSUMO_DESC': 'count'})
    .reset_index()
    .rename(columns={'OF_CDG': 'tem_of', 'INSUMO_DESC': 'qtd_insumos'}))
agrupado['tem_of'] = agrupado['tem_of'].map({True: "✅ Já tem OF", False: "❌ Sem OF"})

# --- Painel ---
st.title("📋 Acompanhamento de Requisições — Semana Atual")
st.metric("Total Requisições", len(agrupado))
st.metric("Requisições com OF", (agrupado['tem_of'] == "✅ Já tem OF").sum())

st.subheader("📊 Resumo por Requisição")
st.dataframe(agrupado)

st.subheader("🔎 Requisições sem OF")
st.dataframe(semana_atual[semana_atual['OF_CDG'].isna()][
    ['REQ_CDG', 'INSUMO_DESC', 'INSUMO_CATEGORIA', 'QTD_PED']])
