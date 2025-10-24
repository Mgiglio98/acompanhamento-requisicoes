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

df = df.drop_duplicates(subset=["REQ_CDG", "INSUMO_CDG", "EMPRD"])

# --- Filtrar semana atual e passada ---
df['REQ_DATA'] = pd.to_datetime(df['REQ_DATA'])
semana_atual_num = pd.Timestamp.now().isocalendar().week
semanas_desejadas = [semana_atual_num, semana_atual_num - 1]
df_duas_semanas = df[df['REQ_DATA'].dt.isocalendar().week.isin(semanas_desejadas)]

# --- Tabela Principal agrupada por Requisição ---
agrupado = (
    df_duas_semanas
    .groupby(['EMPRD', 'REQ_CDG'], as_index=False)
    .agg(
        EMPRD_DESC=('EMPRD_DESC', 'first'),
        EMPRD_UF=('EMPRD_UF', 'first'),
        REQ_DATA=('REQ_DATA', 'first'),
        QTD_INSUMOS=('INSUMO_DESC', 'count'),
        QTD_COMPRADOS=('OF_CDG', lambda x: x.notna().sum())
    )
)

agrupado['QTD_PENDENTE'] = agrupado['QTD_INSUMOS'] - agrupado['QTD_COMPRADOS']

agrupado['STATUS'] = agrupado['QTD_PENDENTE'].apply(
    lambda x: "✅ Todos Comprados" if x == 0 else f"⏳ Não Finalizada")

agrupado = agrupado.sort_values(['REQ_DATA', 'QTD_PENDENTE'], ascending=[True, False])

agrupado = agrupado.set_index('REQ_CDG')

# --- Painel Visual---
st.title("📋 Acompanhamento de Requisições — Semana Atual")

col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("📦 Total Requisições", len(agrupado))
with col2:
    st.metric("✅ Totalmente Compradas", (agrupado['QTD_PENDENTE'] == 0).sum())
with col3:
    st.metric("⏳ Com Pendências", (agrupado['QTD_PENDENTE'] > 0).sum())
with col4:
    total_ofs = df_duas_semanas['OF_CDG'].dropna().nunique()
    st.metric("🧾 Total de OFs Criadas", total_ofs)

st.subheader("📊 Resumo por Requisição")
st.dataframe(agrupado)

st.subheader("🔎 Insumos sem OF")
colunas_exibir = ['EMPRD', 'EMPRD_DESC', 'REQ_CDG', 'INSUMO_CDG', 'INSUMO_DESC']
base_sem_of = df_duas_semanas[df_duas_semanas['OF_CDG'].isna()][colunas_exibir].reset_index(drop=True)
st.dataframe(base_sem_of)


