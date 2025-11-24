import pandas as pd
import streamlit as st
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

# --- Configuração inicial ---
st.set_page_config(page_title="Acompanhamento de Requisições", page_icon="📋", layout="wide")

def enviar_email_smtp(destinatario, assunto, corpo):
    smtp_server = "smtp.office365.com"
    smtp_port = 587
    remetente = "matheus.almeida@osborne.com.br"
    senha = st.secrets["SMTP_PASSWORD"]

    msg = MIMEMultipart()
    msg["From"] = remetente
    msg["To"] = destinatario
    msg["Subject"] = assunto

    msg.attach(MIMEText(corpo, "plain"))

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(remetente, senha)
            server.send_message(msg)
        return True, None

    except Exception as e:
        return False, str(e)

# --- Carregar base ---
DATA_PATH = "AcompReq.xlsx"
DATA_ADM_PATH = "AdmxEmprd.xlsx"

try:
    df = pd.read_excel(DATA_PATH)
except FileNotFoundError:
    st.error("⚠️ O arquivo 'AcompReq.xlsx' não foi encontrado na raiz do repositório.")
    st.stop()

df = df.drop_duplicates(subset=["REQ_CDG", "INSUMO_CDG", "EMPRD"])
df['REQ_DATA'] = pd.to_datetime(df['REQ_DATA'])

# Garante EMPRD como string para facilitar o merge
df["EMPRD"] = df["EMPRD"].astype(str)

# --- Carregar base de administrativos por obra ---
try:
    df_adm = pd.read_excel(DATA_ADM_PATH)
except FileNotFoundError:
    st.error("⚠️ O arquivo 'AdmxEmprd.xlsx' não foi encontrado na raiz do repositório.")
    st.stop()

# Garante EMPRD como string também
df_adm["EMPRD"] = df_adm["EMPRD"].astype(str)

ADM_EMAILS = {
    "MARIA EDUARDA": "maria.eduarda@osborne.com.br",
}

df = df.merge(df_adm, on="EMPRD", how="left")

df["ADM"] = df["ADM"].astype(str).str.strip().str.upper()

# --- Painel Visual---
st.title("📋 Acompanhamento de Requisições — Semana Atual")

# --- Filtro de Obras (EMPRD) logo abaixo do título ---
emprds_disponiveis = sorted(df["EMPRD"].unique())

emprds_escolhidos = st.multiselect(
    "Selecione a(s) Obras (EMPRD):",
    options=emprds_disponiveis,
    default=emprds_disponiveis,
)

# aplica o filtro
if len(emprds_escolhidos) > 0:
    df = df[df["EMPRD"].isin(emprds_escolhidos)]

# --- Filtrar semana atual e passada ---
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
        QTD_COMPRADOS=('OF_CDG', lambda x: x.notna().sum()),
        ADM = ('ADM', 'first')
    )
)

agrupado['QTD_PENDENTE'] = agrupado['QTD_INSUMOS'] - agrupado['QTD_COMPRADOS']
agrupado['STATUS'] = agrupado['QTD_PENDENTE'].apply(
    lambda x: "✅ Todos Comprados" if x == 0 else f"⏳ Não Finalizada")

agrupado = agrupado.sort_values(['REQ_DATA', 'QTD_PENDENTE'], ascending=[True, False])
agrupado = agrupado.set_index('REQ_CDG')

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

st.subheader("📨 Envio de E-mails para Administrativos")

if st.button("Enviar e-mails (teste)"):
    pendentes = agrupado[agrupado["QTD_PENDENTE"] > 0].reset_index()

    if pendentes.empty:
        st.info("Nenhuma requisição pendente para enviar.")
    else:
        grupos = pendentes.groupby("ADM_NOME")

        for adm, grupo in grupos:
            email = ADM_EMAILS.get(adm)

            if email is None:
                st.warning(f"⚠️ ADM '{adm}' não possui e-mail configurado no código.")
                continue

            # Monta corpo do email
            corpo = f"""
Olá {adm},

Segue abaixo o resumo das requisições que ainda possuem itens pendentes:\n
{grupo[['REQ_CDG', 'EMPRD', 'QTD_PENDENTE']].to_string(index=False)}\n
Atenciosamente,
Equipe Suprimentos
"""

            assunto = f"Pendências de Requisições - Obras ({adm})"

            enviado, erro = enviar_email_smtp(destinatario, assunto, corpo)

            if enviado:
                st.success(f"📧 E-mail enviado com sucesso para {adm} — ({email})")
            else:
                st.error(f"❌ Erro ao enviar e-mail para {adm}: {erro}")

st.subheader("📊 Resumo por Requisição")
st.dataframe(agrupado)

st.subheader("🔎 Insumos sem OF")
colunas_exibir = ['EMPRD', 'EMPRD_DESC', 'REQ_CDG', 'INSUMO_CDG', 'INSUMO_DESC']
base_sem_of = df_duas_semanas[df_duas_semanas['OF_CDG'].isna()][colunas_exibir].reset_index(drop=True)
st.dataframe(base_sem_of)

