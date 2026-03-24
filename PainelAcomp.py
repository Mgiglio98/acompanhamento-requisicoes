import pandas as pd
import streamlit as st
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

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

DATA_PATH = "AcompReq.xlsx"
DATA_ADM_PATH = "AdmxEmprd.xlsx"
DATA_OF = "Relat_OF.xlsx"

@st.cache_data
def carregar_bases():
    df = pd.read_excel(DATA_PATH)
    df_adm = pd.read_excel(DATA_ADM_PATH)
    df_of = pd.read_excel(DATA_OF)
    return df, df_adm, df_of

df, df_adm, df_of = carregar_bases()

df = df.drop_duplicates(subset=["REQ_CDG", "INSUMO_CDG", "EMPRD"])
df['REQ_DATA'] = pd.to_datetime(df['REQ_DATA'], errors = "coerce")

df["EMPRD"] = df["EMPRD"].astype("string").str.strip()
df_adm["EMPRD"] = df_adm["EMPRD"].astype("string").str.strip()

df = df[df["EMPRD"] != "500"]

df["OF_CDG"] = pd.to_numeric(df["OF_CDG"], errors="coerce").astype("Int64")

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
    .astype("string")
    .str.normalize("NFKD")
    .str.encode("ascii", "ignore")
    .str.decode("utf-8")
    .str.strip()
    .str.upper()
)
df["ADM"] = df["ADM"].replace({"<NA>": None, "NAN": None})

df_of["OF_CDG"] = pd.to_numeric(df_of["OF_CDG"], errors="coerce").astype("Int64")

df_of = df_of.rename(columns={"OF_DATA": "OF_DATA_RELAT"})

# --- Mantém da base de OF apenas o que será usado no painel ---
colunas_of = ["OF_CDG", "OF_DATA_RELAT", "STATUS_DESC"]
df_of = df_of[colunas_of].drop_duplicates()

# --- Merge com base de OFs ---
df = df.merge(df_of, on="OF_CDG", how="left")

st.title("📋 Acompanhamento de Requisições — Últimos 14 dias")

# --- Filtro inicial de obra ---
emprds_disponiveis = sorted(df["EMPRD"].dropna().unique().tolist())

# valores padrão
default_emprds = emprds_disponiveis

# período base
fim_periodo = pd.Timestamp.now().normalize() + pd.Timedelta(days=1)
inicio_periodo = fim_periodo - pd.Timedelta(days=14)

# base temporária inicial
df_base_temp = df.copy()

col1, col2 = st.columns(2)

with col1:
    emprds_escolhidos = st.multiselect(
        "Selecione a(s) Obras (EMPRD):",
        options=emprds_disponiveis,
        default=default_emprds,
    )

if len(emprds_escolhidos) > 0:
    df_base_temp = df_base_temp[df_base_temp["EMPRD"].isin(emprds_escolhidos)]

# base temporária no período para descobrir os status disponíveis
df_temp_periodo = df_base_temp[
    (df_base_temp["REQ_DATA"] >= inicio_periodo) &
    (df_base_temp["REQ_DATA"] < fim_periodo)
].copy()

df_temp_periodo["PENDENTE_REAL"] = (
    df_temp_periodo["OF_CDG"].isna()
    & (df_temp_periodo["INSUMO_STATUS"] == "Apto")
)

status_por_req_temp = (
    df_temp_periodo
    .groupby(["EMPRD", "REQ_CDG"])["PENDENTE_REAL"]
    .sum()
    .apply(lambda x: "✅ Todos Comprados" if x == 0 else "⏳ Não Finalizada")
    .rename("STATUS_REQ")
    .reset_index()
)

df_temp_periodo = df_temp_periodo.merge(
    status_por_req_temp,
    on=["EMPRD", "REQ_CDG"],
    how="left"
)

status_req_opcoes = sorted(df_temp_periodo["STATUS_REQ"].dropna().unique().tolist())

with col2:
    status_req_escolhidos = st.multiselect(
        "Selecione o(s) Status da Requisição:",
        options=status_req_opcoes,
        default=status_req_opcoes,
    )

# base final filtrada por obra
df_base = df.copy()

if len(emprds_escolhidos) > 0:
    df_base = df_base[df_base["EMPRD"].isin(emprds_escolhidos)]

# base final no período
df_duas_semanas = df_base[
    (df_base["REQ_DATA"] >= inicio_periodo) &
    (df_base["REQ_DATA"] < fim_periodo)
].copy()

# pendência real
df_duas_semanas["PENDENTE_REAL"] = (
    df_duas_semanas["OF_CDG"].isna()
    & (df_duas_semanas["INSUMO_STATUS"] == "Apto")
)

# status calculado
status_por_req = (
    df_duas_semanas
    .groupby(["EMPRD", "REQ_CDG"])["PENDENTE_REAL"]
    .sum()
    .apply(lambda x: "✅ Todos Comprados" if x == 0 else "⏳ Não Finalizada")
    .rename("STATUS_REQ")
    .reset_index()
)

df_duas_semanas = df_duas_semanas.merge(
    status_por_req,
    on=["EMPRD", "REQ_CDG"],
    how="left"
)

# aplica filtro de status
if len(status_req_escolhidos) > 0:
    df_duas_semanas = df_duas_semanas[
        df_duas_semanas["STATUS_REQ"].isin(status_req_escolhidos)
    ]
if not df_duas_semanas.empty:
    periodo_min = df_duas_semanas["REQ_DATA"].min().strftime("%d/%m/%Y")
    periodo_max = df_duas_semanas["REQ_DATA"].max().strftime("%d/%m/%Y")
    st.markdown(
        f"**Período filtrado:** {periodo_min} → {periodo_max}")
else:
    st.markdown("**Período filtrado:** sem requisições no intervalo selecionado.")

# --- Tabela Principal agrupada por Requisição ---
agrupado = (
    df_duas_semanas
    .groupby(['EMPRD', 'REQ_CDG'], as_index=False)
    .agg(
        EMPRD_DESC=('EMPRD_DESC', 'first'),
        EMPRD_UF=('EMPRD_UF', 'first'),
        REQ_DATA=('REQ_DATA', 'min'),
        QTD_INSUMOS=('INSUMO_DESC', 'count'),
        QTD_PENDENTE=('PENDENTE_REAL', 'sum'),
        ADM=('ADM', 'first'),
        STATUS=('STATUS_REQ', 'first'),
    )
)

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

    todos = agrupado.reset_index()

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

            linhas_email.append(f"Obra: {emprd}")
            linhas_email.append(f"Requisição: {req}")

            # OFs geradas para a requisição
            ofs = df_req["OF_CDG"].dropna().astype(str).unique().tolist()
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
agrupado_view = agrupado.reset_index().copy()
agrupado_view["REQ_DATA"] = pd.to_datetime(
    agrupado_view["REQ_DATA"], errors="coerce"
).dt.strftime("%d/%m/%Y")

st.dataframe(agrupado_view, use_container_width=True, hide_index=True)

col_esq, col_dir = st.columns(2)

with col_esq:
    st.subheader("📈 OF's Geradas")
    colunas_exibir = ["REQ_CDG", "EMPRD", "EMPRD_DESC", "OF_CDG", "OF_DATA_RELAT", "STATUS_DESC"]
    base_of_status = (
        df_duas_semanas[df_duas_semanas["OF_CDG"].notna()][colunas_exibir]
        .drop_duplicates()
        .sort_values("OF_CDG", ascending=True)
        .copy())
    base_of_status["OF_DATA_RELAT"] = pd.to_datetime(base_of_status["OF_DATA_RELAT"], errors="coerce").dt.strftime("%d/%m/%Y")
    base_of_status = base_of_status.rename(columns={"OF_DATA_RELAT": "OF_DATA"})
    st.dataframe(
        base_of_status,
        use_container_width=True,
        hide_index=True)

with col_dir:
    st.subheader("🔎 Insumos Pendentes")
    colunas_exibir = ["REQ_CDG", "REQ_DATA", "EMPRD", "EMPRD_DESC", "INSUMO_CDG", "INSUMO_DESC"]
    base_sem_of = df_duas_semanas[df_duas_semanas["PENDENTE_REAL"]][colunas_exibir].copy()
    base_sem_of["REQ_DATA"] = base_sem_of["REQ_DATA"].dt.strftime("%d/%m/%Y")
    st.dataframe(
        base_sem_of,
        use_container_width=True,
        hide_index=True)

