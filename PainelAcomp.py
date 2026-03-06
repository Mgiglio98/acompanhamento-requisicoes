import pandas as pd
import streamlit as st
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

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
df['REQ_DATA'] = pd.to_datetime(df['REQ_DATA'], errors = "coerce")

# Garante EMPRD como string para facilitar o merge
df["EMPRD"] = df["EMPRD"].astype(str)

# --- Remove requisições do empreendimento 500 ---
df = df[df["EMPRD"] != "500"]

df["OF_CDG"] = df["OF_CDG"].apply(
    lambda x: int(x) if isinstance(x, float) and not pd.isna(x) else x
)

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
    "JOICE": "joice.oliveira@osborne.com.br",
    "GRAZIELE": "graziele.horacio@osborne.com.br",
    "MICAELE": "micaele.ferreira@osborne.com.br",
    "ROBERTO": "roberto.santos@osborne.com.br",
}

# --- Merge com administrativos ---
df = df.merge(df_adm, on="EMPRD", how="left")

df["ADM"] = (
    df["ADM"]
    .astype(str)
    .str.normalize("NFKD")
    .str.encode("ascii", "ignore")
    .str.decode("utf-8")
    .str.strip()
    .str.upper()
    .replace("NAN", None)
)

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

# --- Filtrar semana atual e passada (por data, não só pelo número da semana) ---
hoje = pd.Timestamp.now().normalize() + pd.Timedelta(days=1)
limite = hoje - pd.Timedelta(days=14)

df_duas_semanas = df[
    (df['REQ_DATA'] >= limite) &
    (df['REQ_DATA'] <= hoje)
].copy()

# Define pendência real: insumo apto e sem OF
df_duas_semanas["PENDENTE_REAL"] = (
    df_duas_semanas["OF_CDG"].isna()
    & (df_duas_semanas["INSUMO_STATUS"] == "Apto")
)

if not df_duas_semanas.empty:
    periodo_min = df_duas_semanas["REQ_DATA"].min().strftime("%d/%m/%Y")
    periodo_max = df_duas_semanas["REQ_DATA"].max().strftime("%d/%m/%Y")
    st.markdown(
        f"**Período filtrado:** {periodo_min} → {periodo_max}"
    )
else:
    st.markdown("**Período filtrado:** sem requisições no intervalo selecionado.")

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
        QTD_PENDENTE=('PENDENTE_REAL', 'sum'),
        ADM = ('ADM', 'first')
    )
)

agrupado['STATUS'] = agrupado['QTD_PENDENTE'].apply(
    lambda x: "✅ Todos Comprados" if x == 0 else f"⏳ Não Finalizada")

agrupado = agrupado.sort_values(['REQ_DATA', 'EMPRD', 'REQ_CDG'])
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

if st.button("Enviar e-mails"):

    # Agora usamos TODAS as requisições do agrupado (não só pendentes)
    todos = agrupado.reset_index()

    # lista de ADMs reais na base filtrada
    adms_lista = [a for a in todos["ADM"].dropna().unique() if a in ADM_EMAILS]

    for adm in adms_lista:
        email = ADM_EMAILS[adm]

        # Requisições deste ADM
        reqs_adm = todos[todos["ADM"] == adm]["REQ_CDG"].unique()

        # Base detalhada (com insumos + OFs)
        detalhado_adm = df_duas_semanas[
            (df_duas_semanas["ADM"] == adm)
            & (df_duas_semanas["REQ_CDG"].isin(reqs_adm))
        ].copy()

        if detalhado_adm.empty:
            st.warning(f"⚠️ Não há requisições para o ADM {adm}.")
            continue

        # Monta corpo do email
        linhas_email = []
        linhas_email.append(f"Olá {adm},\n")
        linhas_email.append("Segue abaixo as requisições das suas obras realizadas recentemente:\n")

        # Agrupa por obra e requisição
        for (emprd, req), df_req in detalhado_adm.groupby(["EMPRD", "REQ_CDG"]):

            linhas_email.append(f"OC {emprd}")
            linhas_email.append(f"Requisição {req}")

            # OFs geradas para a requisição
            ofs = df_req["OF_CDG"].dropna().unique()
            ofs = [str(int(float(of))) for of in ofs]
            if len(ofs) > 0:
                linhas_email.append(" - OFs geradas:")
                for of in ofs:
                    linhas_email.append(f"     • {of}")
            else:
                linhas_email.append(" - Nenhuma OF gerada")

            # Insumos pendentes de OF
            insumos_pend = df_req[df_req["PENDENTE_REAL"]]["INSUMO_DESC"].dropna().unique()
            if len(insumos_pend) > 0:
                linhas_email.append(" - Insumos pendentes de OF:")
                for insumo in insumos_pend:
                    linhas_email.append(f"     • {insumo}")
            else:
                linhas_email.append(" - Todos os insumos da REQ possuem OF")

            linhas_email.append("")  # linha em branco

        linhas_email.append("Qualquer dúvida, estou à disposição.")

        corpo = "\n".join(linhas_email)
        assunto = f"Resumo das Requisições — Obras Adm {adm}"

        enviado, erro = enviar_email_smtp(email, assunto, corpo)

        if enviado:
            st.success(f"📧 E-mail enviado com sucesso para {adm} — ({email})")
        else:
            st.error(f"❌ Erro ao enviar e-mail para {adm}: {erro}")

st.subheader("📊 Resumo por Requisição")
st.dataframe(agrupado)

# --- Área em 2 colunas para a segunda tabela + nova visualização ---
col_esq, col_dir = st.columns(2)

with col_esq:
    st.subheader("🔎 Insumos sem OF")
    colunas_exibir = ["REQ_CDG", "EMPRD", "EMPRD_DESC", "INSUMO_CDG", "INSUMO_DESC"]
    base_sem_of = df_duas_semanas[df_duas_semanas["PENDENTE_REAL"]][colunas_exibir].copy()
    st.dataframe(
        base_sem_of,
        use_container_width=True,
        hide_index=True
    )

with col_dir:
    st.subheader("📈 Nova visualização")
    st.info("Espaço reservado para a próxima visualização.")
