import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Acompanhamento de Requisições",
    page_icon="📋",
    layout="wide"
)

# =========================
# BASES
# =========================

DATA_PATH = "AcompReq_Base.xlsx"
DATA_ADM_PATH = "AdmxEmprd.xlsx"

def carregar_bases():
    df = pd.read_excel(DATA_PATH)
    df_adm = pd.read_excel(DATA_ADM_PATH)
    return df, df_adm

df, df_adm = carregar_bases()

# Padroniza nomes das colunas
df.columns = df.columns.str.upper()
df_adm.columns = df_adm.columns.str.upper()

# Tratamentos principais
df = df.drop_duplicates(subset=["REQ_CDG", "INSUMO_CDG", "EMPRD"])
df["REQ_DATA"] = pd.to_datetime(df["REQ_DATA"], errors="coerce")
df["OF_DATA"] = pd.to_datetime(df["OF_DATA"], errors="coerce")

df["EMPRD"] = df["EMPRD"].astype("string").str.strip()
df_adm["EMPRD"] = df_adm["EMPRD"].astype("string").str.strip()

df = df[df["EMPRD"] != "500"]

df["OF_CDG"] = pd.to_numeric(df["OF_CDG"], errors="coerce").astype("Int64")

# Merge com administrativos
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


# =========================
# PAINEL
# =========================

st.title("📋 Acompanhamento de Requisições — 2026")

emprds_disponiveis = sorted(df["EMPRD"].dropna().unique().tolist())
default_emprds = emprds_disponiveis

ano_atual = pd.Timestamp.now().year
inicio_periodo = pd.Timestamp(year=ano_atual, month=1, day=1)
fim_periodo = pd.Timestamp(year=ano_atual + 1, month=1, day=1)

col1, col2 = st.columns(2)

with col1:
    emprds_escolhidos = st.multiselect(
        "Selecione a(s) Obras (EMPRD):",
        options=emprds_disponiveis,
        default=default_emprds,
    )

df_base_temp = df.copy()

if len(emprds_escolhidos) > 0:
    df_base_temp = df_base_temp[df_base_temp["EMPRD"].isin(emprds_escolhidos)]

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
    .apply(lambda x: "✅ Finalizada" if x == 0 else "⏳ Com Pendências")
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

df_base = df.copy()

if len(emprds_escolhidos) > 0:
    df_base = df_base[df_base["EMPRD"].isin(emprds_escolhidos)]

df_duas_semanas = df_base[
    (df_base["REQ_DATA"] >= inicio_periodo) &
    (df_base["REQ_DATA"] < fim_periodo)
].copy()

df_duas_semanas["PENDENTE_REAL"] = (
    df_duas_semanas["OF_CDG"].isna()
    & (df_duas_semanas["INSUMO_STATUS"] == "Apto")
)

status_por_req = (
    df_duas_semanas
    .groupby(["EMPRD", "REQ_CDG"])["PENDENTE_REAL"]
    .sum()
    .apply(lambda x: "✅ Finalizada" if x == 0 else "⏳ Com Pendências")
    .rename("STATUS_REQ")
    .reset_index()
)

df_duas_semanas = df_duas_semanas.merge(
    status_por_req,
    on=["EMPRD", "REQ_CDG"],
    how="left"
)

if len(status_req_escolhidos) > 0:
    df_duas_semanas = df_duas_semanas[
        df_duas_semanas["STATUS_REQ"].isin(status_req_escolhidos)
    ]

if not df_duas_semanas.empty:
    periodo_min = df_duas_semanas["REQ_DATA"].min().strftime("%d/%m/%Y")
    periodo_max = df_duas_semanas["REQ_DATA"].max().strftime("%d/%m/%Y")
    st.markdown(f"**Período filtrado:** {periodo_min} → {periodo_max}")
else:
    st.markdown("**Período filtrado:** sem requisições no intervalo selecionado.")

# =========================
# RESUMO AGRUPADO
# =========================

agrupado = (
    df_duas_semanas
    .groupby(["EMPRD", "REQ_CDG"], as_index=False)
    .agg(
        EMPRD_DESC=("EMPRD_DESC", "first"),
        EMPRD_UF=("EMPRD_UF", "first"),
        REQ_DATA=("REQ_DATA", "min"),
        QTD_INSUMOS=("INSUMO_DESC", "count"),
        QTD_PENDENTE=("PENDENTE_REAL", "sum"),
        ADM=("ADM", "first"),
        STATUS=("STATUS_REQ", "first"),
    )
)

agrupado = agrupado.sort_values(["REQ_DATA", "EMPRD", "REQ_CDG"])
agrupado = agrupado.set_index("REQ_CDG")

col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric("📦 Total Requisições", len(agrupado))

with col2:
    st.metric("✅ Total Finalizadas", (agrupado["QTD_PENDENTE"] == 0).sum())

with col3:
    st.metric("⏳ Com Pendências", (agrupado["QTD_PENDENTE"] > 0).sum())

with col4:
    total_ofs = df_duas_semanas["OF_CDG"].dropna().nunique()
    st.metric("🧾 Total de OFs Criadas", total_ofs)

# =========================
# TABELAS
# =========================

st.subheader("📊 Resumo por Requisição")

agrupado_view = agrupado.reset_index().copy()

agrupado_view["REQ_DATA"] = pd.to_datetime(
    agrupado_view["REQ_DATA"],
    errors="coerce"
).dt.strftime("%d/%m/%Y")

agrupado_view = agrupado_view.rename(columns={
    "REQ_CDG": "Requisição",
    "EMPRD": "Nº da Obra",
    "EMPRD_DESC": "Empreendimento",
    "EMPRD_UF": "Estado",
    "REQ_DATA": "Data da Requisição",
    "QTD_INSUMOS": "Insumos Solicitados",
    "QTD_PENDENTE": "Insumos Pendentes",
    "ADM": "ADM da Obra",
    "STATUS": "Status de Compra",
})

st.dataframe(agrupado_view, use_container_width=True, hide_index=True)

col_esq, col_dir = st.columns(2)

with col_esq:
    st.subheader("📈 OF's Geradas")

    colunas_exibir = [
        "REQ_CDG",
        "EMPRD",
        "EMPRD_DESC",
        "OF_CDG",
        "OF_DATA",
        "STATUS_DESC",
    ]

    base_of_status = (
        df_duas_semanas[df_duas_semanas["OF_CDG"].notna()][colunas_exibir]
        .drop_duplicates()
        .sort_values("OF_CDG", ascending=True)
        .copy()
    )

    base_of_status["OF_DATA"] = pd.to_datetime(
        base_of_status["OF_DATA"],
        errors="coerce"
    ).dt.strftime("%d/%m/%Y")

    base_of_status = base_of_status.rename(columns={
        "REQ_CDG": "Requisição",
        "EMPRD": "Nº da Obra",
        "EMPRD_DESC": "Empreendimento",
        "OF_CDG": "OF",
        "OF_DATA": "Data da OF",
        "STATUS_DESC": "Status da OF",
    })

    st.dataframe(base_of_status, use_container_width=True, hide_index=True)

with col_dir:
    st.subheader("🔎 Insumos Pendentes")

    colunas_exibir = [
        "REQ_CDG",
        "REQ_DATA",
        "EMPRD",
        "EMPRD_DESC",
        "INSUMO_DESC",
    ]

    base_sem_of = df_duas_semanas[
        df_duas_semanas["PENDENTE_REAL"]
    ][colunas_exibir].copy()

    base_sem_of["REQ_DATA"] = base_sem_of["REQ_DATA"].dt.strftime("%d/%m/%Y")

    base_sem_of = base_sem_of.rename(columns={
        "REQ_DATA": "Data da Requisição",
        "REQ_CDG": "Requisição",
        "EMPRD": "Nº da Obra",
        "EMPRD_DESC": "Empreendimento",
        "INSUMO_DESC": "Insumo",
    })

    st.dataframe(base_sem_of, use_container_width=True, hide_index=True)
