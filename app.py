import pandas as pd
import streamlit as st
from datetime import datetime, timedelta

st.set_page_config(layout="wide")
st.title("📅 Sistema de Escala")

# ========================
# CONSTANTES
# ========================

TIPOS_ESPECIAIS = ["Recesso", "Final de Semana", "Noturno", "Apoio"]
INTERVALO_MINIMO = 6

# ========================
# LOAD DATA
# ========================

@st.cache_data
def carregar_dados():
    df = pd.read_csv("dados.csv")
    df.columns = df.columns.str.strip().str.upper()

    df["NOME"] = df["NOME"].astype(str).str.strip()

    df = df.dropna(
        how="all",
        subset=["FINAL DE SEMANA", "NOTURNO", "RECESSO", "APOIO"]
    )

    df_long = df.melt(
        id_vars=["NOME"],
        var_name="tipo_plantao",
        value_name="ultima_data"
    )

    df_long = df_long.dropna(subset=["ultima_data"])

    df_long["ultima_data"] = pd.to_datetime(
        df_long["ultima_data"],
        format="%d/%m/%Y",
        errors="coerce"
    )

    df_long = df_long.dropna(subset=["ultima_data"])

    return df_long


@st.cache_data
def carregar_ferias():
    df = pd.read_csv("ferias.csv")
    df.columns = df.columns.str.strip().str.upper()

    df["NOME"] = df["NOME"].astype(str).str.strip()

    df_long = df.melt(
        id_vars=["NOME"],
        var_name="data",
        value_name="ferias"
    )

    df_long["data"] = pd.to_datetime(
        df_long["data"],
        format="%d/%m/%y",
        errors="coerce"
    )

    df_long = df_long.dropna(subset=["data"])
    df_long = df_long[df_long["ferias"] == True]

    return df_long


@st.cache_data
def carregar_regioes():
    df = pd.read_csv("regioes.csv")
    df.columns = df.columns.str.strip().str.upper()

    df["NOME"] = df["NOME"].astype(str).str.strip()
    df["REGIAO"] = df["REGIAO"].astype(str).str.strip()

    return df.set_index("NOME")["REGIAO"].to_dict()


@st.cache_data
def carregar_calendario(ANO):
    df_raw = pd.read_csv("calendario_raw.csv")

    eventos = []

    for _, row in df_raw.iterrows():
        if pd.notna(row.get("Feriados")):
            data = datetime.strptime(row["Feriados"] + f"/{ANO}", "%d/%m/%Y")
            eventos.append({"data": data, "tipo": "feriado"})

        if pd.notna(row.get("Recesso")):
            data = datetime.strptime(row["Recesso"] + f"/{ANO}", "%d/%m/%Y")
            eventos.append({"data": data, "tipo": "recesso"})

    return pd.DataFrame(eventos)


df_long = carregar_dados()
df_ferias = carregar_ferias()
mapeamento_regioes = carregar_regioes()

# ========================
# INPUTS
# ========================

st.sidebar.header("Configurações Gerais")

ANO = st.sidebar.number_input("Ano", value=2026)
MES = st.sidebar.number_input("Mês", value=4)

qtd_recesso = st.sidebar.number_input("Recesso por dia", 0, 50, 2)
qtd_final = st.sidebar.number_input("Final de Semana por dia", 0, 50, 2)
qtd_noturno = st.sidebar.number_input("Noturno por dia", 0, 50, 2)
qtd_apoio = st.sidebar.number_input("Apoio por dia", 0, 50, 0)

# ========================
# CONFIG REGIÕES
# ========================

regioes = sorted(set(mapeamento_regioes.values()))
st.sidebar.header("⚙️ Configuração por Região")

config_regioes = {}

for reg in regioes:
    with st.sidebar.expander(f"Região: {reg}"):

        qtd_diarios = st.number_input(
            f"Qtd plantões diurnos por dia - {reg}",
            min_value=0,
            max_value=50,
            value=2,
            key=f"diarios_{reg}"
        )

        max_mes = st.number_input(
            f"Máx plantões diurnos por pessoa no mês - {reg}",
            min_value=0,
            max_value=50,
            value=2,
            key=f"max_mes_{reg}"
        )

        config_regioes[reg] = {
            "diurnos_por_dia": qtd_diarios,
            "max_mes": max_mes
        }

# ========================
# FUNÇÕES
# ========================

def esta_de_ferias(nome, data):
    return not df_ferias[
        (df_ferias["NOME"] == nome) &
        (df_ferias["data"] == data)
    ].empty


def tipo_do_dia(data, df_calendario):
    linha = df_calendario[df_calendario["data"] == data]

    if not linha.empty:
        tipo = linha.iloc[0]["tipo"]
        if tipo == "recesso":
            return "recesso"
        elif tipo == "feriado":
            return "final_semana"

    if data.weekday() >= 5:
        return "final_semana"

    return "noturno"

# ========================
# GERADOR
# ========================

def gerar_escala(ANO, MES, qtd_recesso, qtd_final, qtd_noturno, qtd_apoio):

    df_calendario = carregar_calendario(ANO)

    datas = pd.date_range(datetime(ANO, MES, 1),
                          datetime(ANO, MES + 1, 1) - timedelta(days=1))

    agenda = {
        data: {
            "tipo": tipo_do_dia(data, df_calendario),
            "Recesso": [],
            "Final de Semana": [],
            "Noturno": [],
            "Apoio": [],
            "Diurno": {}
        } for data in datas
    }

    nomes_por_regiao = {}
    for nome, reg in mapeamento_regioes.items():
        nomes_por_regiao.setdefault(reg, []).append(nome)

    contador = {
        nome: {
            "Diurno": 0,
            "Noturno": 0,
            "Final de Semana": 0,
            "Recesso": 0,
            "Apoio": 0
        }
        for nome in mapeamento_regioes
    }

    historico_diurno = {nome: [] for nome in mapeamento_regioes}
    historico_especiais = {nome: [] for nome in mapeamento_regioes}
    historico_noturno_semana = {nome: set() for nome in mapeamento_regioes}
    historico_plantao = {nome: [] for nome in mapeamento_regioes}
    uso_mes = {nome: 0 for nome in mapeamento_regioes}

    usados_global = set()
    usados_dia = {data: set() for data in datas}

    # ========================
    # ESPECIAIS PRIMEIRO
    # ========================

    def map_tipo(tipo):
        return {
            "Recesso": "recesso",
            "Final de Semana": "final_semana",
            "Noturno": "noturno",
            "Apoio": "apoio"
        }.get(tipo)

    def preencher_tipo(tipo, quantidade):

        if quantidade == 0:
            return

        base = (
            df_long[df_long["tipo_plantao"] == tipo.upper()]
            .sort_values("ultima_data")
            .drop_duplicates(subset=["NOME"])
        )

        for data, info in agenda.items():

            if info["tipo"] != map_tipo(tipo):
                continue

            while len(info[tipo]) < quantidade:

                for _, row in base.iterrows():
                    nome = row["NOME"]

                    if (
                        nome not in usados_global
                        and not esta_de_ferias(nome, data)
                        and nome not in usados_dia[data]
                    ):

                        # BLOQUEIO: diurno próximo
                        if any(abs((data - d).days) < INTERVALO_MINIMO for d in historico_diurno[nome]):
                            continue

                        info[tipo].append(nome)
                        usados_global.add(nome)
                        usados_dia[data].add(nome)

                        contador[nome][tipo] += 1

                        historico_especiais[nome].append(data)
                        historico_plantao[nome].append(data)

                        if tipo == "Noturno":
                            semana = data.isocalendar()[1]
                            historico_noturno_semana[nome].add(semana)

                        break
                else:
                    break

    preencher_tipo("Recesso", qtd_recesso)
    preencher_tipo("Final de Semana", qtd_final)
    preencher_tipo("Noturno", qtd_noturno)
    preencher_tipo("Apoio", qtd_apoio)

    # ========================
    # DIURNO DEPOIS
    # ========================

    for data in datas:

        if agenda[data]["tipo"] != "noturno":
            continue

        semana = data.isocalendar()[1]

        for reg, nomes in nomes_por_regiao.items():

            qtd = config_regioes[reg]["diurnos_por_dia"]
            max_mes = config_regioes[reg]["max_mes"]

            escolhidos = []

            candidatos = (
                df_long[df_long["NOME"].isin(nomes)]
                .sort_values("ultima_data")
                .drop_duplicates("NOME")
            )

            for _, row in candidatos.iterrows():

                nome = row["NOME"]

                if len(escolhidos) >= qtd:
                    break

                if esta_de_ferias(nome, data):
                    continue

                if uso_mes[nome] >= max_mes:
                    continue

                # BLOQUEIO: especial próximo
                if any(abs((data - d).days) < INTERVALO_MINIMO for d in historico_especiais[nome]):
                    continue

                if any((data - d).days < 6 for d in historico_diurno[nome]):
                    continue

                if semana in historico_noturno_semana[nome]:
                    continue

                ultimos_14 = [d for d in historico_plantao[nome] if (data - d).days <= 14]
                if len(ultimos_14) >= 3:
                    continue

                escolhidos.append(nome)

                historico_diurno[nome].append(data)
                historico_plantao[nome].append(data)
                uso_mes[nome] += 1
                contador[nome]["Diurno"] += 1

            agenda[data]["Diurno"][reg] = escolhidos

    df_resultado = pd.DataFrame([
        {
            "Data": d.strftime("%d/%m/%Y"),
            "Tipo": agenda[d]["tipo"],
            "Recesso": ", ".join(agenda[d]["Recesso"]),
            "Final de Semana": ", ".join(agenda[d]["Final de Semana"]),
            "Noturno": ", ".join(agenda[d]["Noturno"]),
            "Apoio": ", ".join(agenda[d]["Apoio"]),
            "Diurno": str(agenda[d]["Diurno"])
        } for d in datas
    ])

    df_contador = pd.DataFrame.from_dict(contador, orient="index")
    df_contador["Total"] = df_contador.sum(axis=1)

    return df_resultado, df_contador.reset_index().rename(columns={"index": "Nome"})

# ========================
# UI
# ========================

if st.button("🚀 Gerar Escala"):

    df, df_contador = gerar_escala(
        ANO, MES, qtd_recesso, qtd_final, qtd_noturno, qtd_apoio
    )

    st.subheader("📅 Escala")
    st.dataframe(df, use_container_width=True)

    st.subheader("📊 Contador de Plantões")
    st.dataframe(df_contador, use_container_width=True)