import io
import random
import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(layout="wide")
st.title("📅 Sistema de Escala")

# ========================
# CONSTANTES
# ========================

TIPOS_ESPECIAIS = ["Recesso", "Final de Semana", "Noturno", "Apoio"]
INTERVALO_MINIMO = 6

# Prioridade de dias da semana para plantões diurnos
# weekday(): 0=seg, 1=ter, 2=qua, 3=qui, 4=sex, 5=sáb, 6=dom
# Ordem de preenchimento: segunda > sexta > quarta > quinta > terça
PRIORIDADE_DIA_SEMANA = {
    0: 1,  # segunda-feira  → prioridade 1 (maior)
    4: 2,  # sexta-feira    → prioridade 2
    2: 3,  # quarta-feira   → prioridade 3
    3: 4,  # quinta-feira   → prioridade 4
    1: 5,  # terça-feira    → prioridade 5 (menor)
}

# ========================
# LOAD DATA
# ========================

@st.cache_data
def carregar_dados():
    df = pd.read_csv("dados.csv")
    df.columns = df.columns.str.strip().str.upper()
    df["NOME"] = df["NOME"].astype(str).str.strip()
    df = df.dropna(how="all", subset=["FINAL DE SEMANA", "NOTURNO", "RECESSO", "APOIO"])
    df_long = df.melt(id_vars=["NOME"], var_name="tipo_plantao", value_name="ultima_data")
    df_long = df_long.dropna(subset=["ultima_data"])
    df_long["ultima_data"] = pd.to_datetime(df_long["ultima_data"], format="%d/%m/%Y", errors="coerce")
    df_long = df_long.dropna(subset=["ultima_data"])
    return df_long


@st.cache_data
def carregar_ferias():
    df = pd.read_csv("ferias.csv")
    df.columns = df.columns.str.strip().str.upper()
    df["NOME"] = df["NOME"].astype(str).str.strip()
    df_long = df.melt(id_vars=["NOME"], var_name="data", value_name="ferias")
    df_long["data"] = pd.to_datetime(df_long["data"], format="%d/%m/%y", errors="coerce")
    df_long = df_long.dropna(subset=["data"])
    df_long = df_long[df_long["ferias"] == True]
    return df_long


@st.cache_data
def calcular_criticidade():
    """
    Calcula o índice de criticidade de cada pessoa com base no número de
    dias de férias marcados como True no ferias.csv.
    Quanto mais dias de férias, mais crítica é a pessoa (menos dias disponíveis),
    e maior deve ser sua prioridade de escalamento.
    Retorna um dict {nome: dias_de_ferias} — valores maiores = mais crítico.
    """
    df = pd.read_csv("ferias.csv")
    df.columns = df.columns.str.strip().str.upper()
    df["NOME"] = df["NOME"].astype(str).str.strip()
    colunas_datas = [c for c in df.columns if c != "NOME"]
    criticidade = {}
    for _, row in df.iterrows():
        nome = row["NOME"]
        total_ferias = sum(1 for col in colunas_datas if row[col] == True)
        criticidade[nome] = total_ferias
    return criticidade


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
criticidade_ferias = calcular_criticidade()

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

n_iteracoes = st.sidebar.number_input(
    "Iterações para otimização",
    min_value=1, max_value=50, value=10,
    help=(
        "Quantas vezes o algoritmo roda com variações para encontrar "
        "o melhor resultado. Mais iterações = melhor resultado, porém mais lento."
    )
)

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
            min_value=0, max_value=50, value=2, key=f"diarios_{reg}"
        )
        max_mes = st.number_input(
            f"Máx plantões diurnos por pessoa no mês - {reg}",
            min_value=0, max_value=50, value=2, key=f"max_mes_{reg}"
        )
        config_regioes[reg] = {"diurnos_por_dia": qtd_diarios, "max_mes": max_mes}

# ========================
# FUNÇÕES
# ========================

def esta_de_ferias(nome, data):
    return not df_ferias[
        (df_ferias["NOME"] == nome) & (df_ferias["data"] == data)
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


def prioridade_diurno(data):
    return PRIORIDADE_DIA_SEMANA.get(data.weekday(), 99)


def get_criticidade(nome):
    return criticidade_ferias.get(nome, 0)


# ========================
# PONTUAÇÃO (análise combinatória)
# Critérios em ordem de prioridade (comparação lexicográfica):
#   1. Especiais preenchidos no mês
#   2. Diurnos em datas críticas preenchidos no mês
#   3. Aparições de nomes críticos (com repetição) em especiais + diurnos
#   4. Total de diurnos preenchidos no mês
# ========================

def calcular_score(agenda):
    especiais_preenchidos = sum(
        len(info[tipo])
        for info in agenda.values()
        for tipo in TIPOS_ESPECIAIS
    )
    diurnos_criticos = sum(
        sum(len(nomes) for nomes in info["Diurno"].values())
        for data, info in agenda.items()
        if data.weekday() in PRIORIDADE_DIA_SEMANA
    )
    nomes_criticos_count = 0
    for info in agenda.values():
        for tipo in TIPOS_ESPECIAIS:
            for nome in info[tipo]:
                if get_criticidade(nome) > 0:
                    nomes_criticos_count += 1
        for nomes in info["Diurno"].values():
            for nome in nomes:
                if get_criticidade(nome) > 0:
                    nomes_criticos_count += 1
    diurnos_total = sum(
        sum(len(nomes) for nomes in info["Diurno"].values())
        for info in agenda.values()
    )
    return (especiais_preenchidos, diurnos_criticos, nomes_criticos_count, diurnos_total)


# ========================
# GERADOR INTERNO
# Aceita uma semente para controlar o ruído aleatório de desempate.
# O ruído é aplicado APÓS os critérios principais (ultima_data, criticidade),
# portanto só afeta empates — nunca quebra a ordem de prioridade.
# ========================

def _gerar_escala_interna(ANO, MES, qtd_recesso, qtd_final, qtd_noturno, qtd_apoio, semente):

    rng = random.Random(semente)

    df_calendario = carregar_calendario(ANO)

    datas = pd.date_range(
        datetime(ANO, MES, 1),
        datetime(ANO, MES + 1, 1) - timedelta(days=1)
    )

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
        nome: {"Diurno": 0, "Noturno": 0, "Final de Semana": 0, "Recesso": 0, "Apoio": 0}
        for nome in mapeamento_regioes
    }

    historico_diurno     = {nome: [] for nome in mapeamento_regioes}
    historico_especiais  = {nome: [] for nome in mapeamento_regioes}
    historico_noturno_semana = {nome: set() for nome in mapeamento_regioes}
    historico_plantao    = {nome: [] for nome in mapeamento_regioes}
    uso_mes              = {nome: 0  for nome in mapeamento_regioes}

    # Controla o limite de 1 especial por pessoa por mês.
    # Quem está neste set NÃO pode receber um 2º especial,
    # mas PODE e DEVE ser considerado para plantões diurnos.
    ja_fez_especial = set()

    usados_dia = {data: set() for data in datas}

    def _ruido():
        return rng.uniform(0, 1e-9)

    # ========================
    # ESPECIAIS PRIMEIRO
    # Ordem obrigatória: Recesso → Final de Semana → Noturno → Apoio
    # Quem faz um tipo NÃO pode fazer outro (max 1 especial/mês por pessoa).
    # Após o especial, a pessoa fica elegível para plantões diurnos.
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
        base = base.copy()
        base["criticidade"] = base["NOME"].map(get_criticidade)
        base["ruido"] = [_ruido() for _ in range(len(base))]
        base = base.sort_values(
            ["ultima_data", "criticidade", "ruido"],
            ascending=[True, False, True]
        )

        for data, info in agenda.items():
            if info["tipo"] != map_tipo(tipo):
                continue

            while len(info[tipo]) < quantidade:
                for _, row in base.iterrows():
                    nome = row["NOME"]
                    if (
                        nome not in ja_fez_especial
                        and not esta_de_ferias(nome, data)
                        and nome not in usados_dia[data]
                    ):
                        if any(abs((data - d).days) < INTERVALO_MINIMO for d in historico_diurno[nome]):
                            continue

                        info[tipo].append(nome)
                        ja_fez_especial.add(nome)
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
    # REGRA: quem fez especial PODE e DEVE fazer diurno também.
    # ja_fez_especial NÃO é verificado aqui — apenas bloqueia 2º especial.
    #
    # Processamento por prioridade de dia da semana:
    # segunda > sexta > quarta > quinta > terça
    #
    # Dentro de cada dia, candidatos ordenados por:
    #   1. Já fez especial este mês (prioridade máxima — precisa de diurnos)
    #   2. Menos diurnos no mês (distribuição equitativa)
    #   3. Último plantão mais antigo (respeita quem faz mais tempo)
    #   4. Mais crítico (desempate)
    #   5. Ruído controlado (variação entre iterações, apenas em empates finais)
    # ========================

    datas_diurnas = sorted(
        [data for data in datas if agenda[data]["tipo"] == "noturno"],
        key=lambda d: (prioridade_diurno(d), d)
    )

    for data in datas_diurnas:
        semana = data.isocalendar()[1]

        for reg, nomes in nomes_por_regiao.items():
            qtd     = config_regioes[reg]["diurnos_por_dia"]
            max_mes = config_regioes[reg]["max_mes"]
            escolhidos = []

            candidatos_base = (
                df_long[df_long["NOME"].isin(nomes)]
                .sort_values("ultima_data")
                .drop_duplicates("NOME")
            )
            candidatos_base = candidatos_base.copy()
            candidatos_base["tem_especial"]  = candidatos_base["NOME"].map(lambda n: 1 if n in ja_fez_especial else 0)
            candidatos_base["criticidade"]   = candidatos_base["NOME"].map(get_criticidade)
            candidatos_base["uso_mes_atual"] = candidatos_base["NOME"].map(uso_mes)
            candidatos_base["ruido"]         = [_ruido() for _ in range(len(candidatos_base))]
            candidatos_base = candidatos_base.sort_values(
                ["tem_especial", "uso_mes_atual", "ultima_data", "criticidade", "ruido"],
                ascending=[False, True, True, False, True]
            )

            for _, row in candidatos_base.iterrows():
                nome = row["NOME"]
                if len(escolhidos) >= qtd:
                    break
                if esta_de_ferias(nome, data):
                    continue
                if uso_mes[nome] >= max_mes:
                    continue
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

    return agenda, datas, contador


def _formatar_resultado(agenda, datas, contador):
    df_resultado = pd.DataFrame([
        {
            "Data": d.strftime("%d/%m/%Y"),
            "Dia": d.strftime("%A"),
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
    df_contador["Dias de Férias (Criticidade)"] = df_contador.index.map(get_criticidade)
    return df_resultado, df_contador.reset_index().rename(columns={"index": "Nome"})


# ========================
# OTIMIZADOR
# Roda N iterações com sementes diferentes e retorna a melhor escala
# segundo a comparação lexicográfica dos 4 critérios de pontuação.
# ========================

def gerar_melhor_escala(ANO, MES, qtd_recesso, qtd_final, qtd_noturno, qtd_apoio, n_iteracoes):
    melhor_agenda   = None
    melhor_datas    = None
    melhor_contador = None
    melhor_score    = (-1, -1, -1, -1)
    historico_scores = []

    barra = st.progress(0, text="Calculando iterações...")

    for i in range(n_iteracoes):
        semente = i
        agenda, datas, contador = _gerar_escala_interna(
            ANO, MES, qtd_recesso, qtd_final, qtd_noturno, qtd_apoio, semente
        )
        score = calcular_score(agenda)
        historico_scores.append({
            "Iteração": i + 1,
            "Semente": semente,
            "Especiais preenchidos": score[0],
            "Diurnos em datas críticas": score[1],
            "Aparições de nomes críticos": score[2],
            "Total de diurnos": score[3],
        })
        if score > melhor_score:
            melhor_score    = score
            melhor_agenda   = agenda
            melhor_datas    = datas
            melhor_contador = contador

        barra.progress((i + 1) / n_iteracoes, text=f"Iteração {i + 1}/{n_iteracoes}...")

    barra.empty()
    df_resultado, df_contador = _formatar_resultado(melhor_agenda, melhor_datas, melhor_contador)
    df_scores = pd.DataFrame(historico_scores)
    return df_resultado, df_contador, melhor_score, df_scores, melhor_agenda, melhor_datas


# ========================
# EXPORTAÇÃO EXCEL
# Layout inspirado na imagem de referência:
#   - Linha 1: datas (dd/mm/aa) — fundo azul escuro, texto branco
#   - Linha 2: dia da semana   — fundo azul claro
#   - Blocos R1…Rn: uma linha por slot de diurno, alternando amarelo/branco
#     fins de semana com fundo rosado; separador cinza entre regiões
#   - Linha PS: plantões especiais, coloridos por tipo
#   - Linha AP: apoio
# ========================

DIAS_PT = {
    "Monday": "segunda-feira", "Tuesday": "terça-feira",
    "Wednesday": "quarta-feira", "Thursday": "quinta-feira",
    "Friday": "sexta-feira", "Saturday": "sábado", "Sunday": "domingo",
}

# Paleta de cores (hex sem #)
COR = {
    "cab_data":    "4472C4",  # azul — cabeçalho datas
    "cab_dia":     "D9E1F2",  # azul claro — cabeçalho dia da semana
    "reg_label":   "2F4F4F",  # verde-escuro — label R1, R2…
    "diurno_a":    "FFF2CC",  # amarelo claro — linhas pares de diurno
    "diurno_b":    "FFFFFF",  # branco — linhas ímpares de diurno
    "fds_bg":      "FFF0F0",  # rosado — coluna de fim de semana
    "sep":         "F2F2F2",  # cinza — separador entre regiões
    "ps_label":    "1F3864",  # azul muito escuro — label PS
    "ps_noturno":  "BDD7EE",  # azul claro — noturno
    "ps_fds":      "FCE4D6",  # salmão — final de semana
    "ps_recesso":  "E2EFDA",  # verde claro — recesso
    "ap_label":    "595959",  # cinza escuro — label AP
    "ap_bg":       "F2F2F2",  # cinza claro — célula de apoio
}

def _f(hex_color):
    """Cria um PatternFill sólido."""
    return PatternFill("solid", start_color=hex_color, fgColor=hex_color)

def _b(style="thin", color="CCCCCC"):
    s = Side(style=style, color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def gerar_excel(agenda, datas, mapeamento_regioes, config_regioes):
    wb = Workbook()
    ws = wb.active
    ws.title = "Escala"

    regioes_ord = sorted(set(mapeamento_regioes.values()))
    datas_list  = list(datas)
    n_datas     = len(datas_list)

    slots_por_regiao = {
        reg: max(config_regioes[reg]["diurnos_por_dia"], 1)
        for reg in regioes_ord
    }

    # ── Cabeçalho linha 1: datas ──────────────────────────────────────────────
    ws.cell(row=1, column=1, value="").fill = _f(COR["reg_label"])
    ws.cell(row=1, column=1).border = _b()

    for ci, data in enumerate(datas_list, start=2):
        c = ws.cell(row=1, column=ci, value=data.strftime("%d/%m/%y"))
        c.font      = Font(bold=True, color="FFFFFF", size=10, name="Arial")
        c.fill      = _f(COR["cab_data"])
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = _b()

    # ── Cabeçalho linha 2: dia da semana ──────────────────────────────────────
    ws.cell(row=2, column=1, value="").fill   = _f(COR["reg_label"])
    ws.cell(row=2, column=1).border = _b()

    for ci, data in enumerate(datas_list, start=2):
        dia = DIAS_PT.get(data.strftime("%A"), data.strftime("%A"))
        c = ws.cell(row=2, column=ci, value=dia)
        c.font      = Font(size=9, color="444444", name="Arial")
        c.fill      = _f(COR["cab_dia"])
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = _b()

    cur = 3  # linha atual

    # ── Blocos de diurno por região ───────────────────────────────────────────
    for ri, reg in enumerate(regioes_ord):
        n_slots = slots_por_regiao[reg]
        cor_diurno = COR["diurno_a"] if ri % 2 == 0 else COR["diurno_b"]

        for slot in range(n_slots):
            row = cur + slot

            # Coluna A: label da região (só na primeira linha do bloco)
            lc = ws.cell(row=row, column=1)
            lc.value     = reg if slot == 0 else ""
            lc.font      = Font(bold=True, color="FFFFFF", size=10, name="Arial")
            lc.fill      = _f(COR["reg_label"])
            lc.alignment = Alignment(horizontal="center", vertical="center")
            lc.border    = _b()

            for ci, data in enumerate(datas_list, start=2):
                nomes    = agenda[data]["Diurno"].get(reg, [])
                valor    = nomes[slot] if slot < len(nomes) else ""
                tipo_dia = agenda[data]["tipo"]
                bg = COR["fds_bg"] if tipo_dia in ("final_semana", "recesso") else cor_diurno

                c = ws.cell(row=row, column=ci, value=valor)
                c.fill      = _f(bg)
                c.font      = Font(size=9, bold=bool(valor), name="Arial")
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border    = _b()

        cur += n_slots

        # Separador cinza entre regiões
        for ci in range(1, n_datas + 2):
            c = ws.cell(row=cur, column=ci, value="")
            c.fill   = _f(COR["sep"])
            c.border = _b()
        cur += 1

    # ── Linha PS (especiais, exceto Apoio) ────────────────────────────────────
    max_ps = max(
        max(
            len(agenda[d]["Recesso"]) +
            len(agenda[d]["Final de Semana"]) +
            len(agenda[d]["Noturno"])
            for d in datas_list
        ),
        1
    )

    for slot in range(max_ps):
        row = cur + slot

        lc = ws.cell(row=row, column=1)
        lc.value     = "PS" if slot == 0 else ""
        lc.font      = Font(bold=True, color="FFFFFF", size=10, name="Arial")
        lc.fill      = _f(COR["ps_label"])
        lc.alignment = Alignment(horizontal="center", vertical="center")
        lc.border    = _b()

        for ci, data in enumerate(datas_list, start=2):
            # Monta lista ordenada: Recesso → FdS → Noturno
            itens = (
                [(n, "Recesso")      for n in agenda[data]["Recesso"]] +
                [(n, "Final de Semana") for n in agenda[data]["Final de Semana"]] +
                [(n, "Noturno")      for n in agenda[data]["Noturno"]]
            )
            tipo_dia = agenda[data]["tipo"]

            if slot < len(itens):
                nome, tipo_esp = itens[slot]
                valor = f"{nome} {tipo_esp}"
                bg = (
                    COR["ps_recesso"] if tipo_esp == "Recesso"
                    else COR["ps_fds"] if tipo_esp == "Final de Semana"
                    else COR["ps_noturno"]
                )
            else:
                valor = ""
                bg = COR["ps_fds"] if tipo_dia in ("final_semana", "recesso") else COR["ps_noturno"]

            c = ws.cell(row=row, column=ci, value=valor)
            c.fill      = _f(bg)
            c.font      = Font(size=9, bold=bool(valor), name="Arial")
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border    = _b()

    cur += max_ps

    # ── Linha AP (Apoio) ──────────────────────────────────────────────────────
    max_ap = max(
        max((len(agenda[d]["Apoio"]) for d in datas_list), default=0),
        1
    )

    for slot in range(max_ap):
        row = cur + slot

        lc = ws.cell(row=row, column=1)
        lc.value     = "AP" if slot == 0 else ""
        lc.font      = Font(bold=True, color="FFFFFF", size=10, name="Arial")
        lc.fill      = _f(COR["ap_label"])
        lc.alignment = Alignment(horizontal="center", vertical="center")
        lc.border    = _b()

        for ci, data in enumerate(datas_list, start=2):
            apoios = agenda[data]["Apoio"]
            valor  = apoios[slot] if slot < len(apoios) else "Apoio"
            c = ws.cell(row=row, column=ci, value=valor)
            c.fill      = _f(COR["ap_bg"])
            c.font      = Font(size=9, color="666666", name="Arial")
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border    = _b()

    # ── Larguras e alturas ────────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 7
    for ci in range(2, n_datas + 2):
        ws.column_dimensions[get_column_letter(ci)].width = 18

    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 14

    # Congela cabeçalhos e coluna de label
    ws.freeze_panes = "B3"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ========================
# UI
# ========================

if st.button("🚀 Gerar Escala"):

    df, df_contador, melhor_score, df_scores, melhor_agenda, melhor_datas = gerar_melhor_escala(
        ANO, MES, qtd_recesso, qtd_final, qtd_noturno, qtd_apoio, int(n_iteracoes)
    )

    # --- Painel de pontuação da melhor escala ---
    st.subheader("🏆 Melhor Escala Encontrada")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Especiais preenchidos",    melhor_score[0])
    col2.metric("Diurnos em datas críticas", melhor_score[1])
    col3.metric("Aparições de críticos",    melhor_score[2])
    col4.metric("Total de diurnos",         melhor_score[3])

    # --- Comparativo de todas as iterações ---
    with st.expander("📊 Comparativo de todas as iterações"):
        def destacar_melhor(df):
            score_cols = [
                "Especiais preenchidos",
                "Diurnos em datas críticas",
                "Aparições de nomes críticos",
                "Total de diurnos",
            ]
            scores    = list(df[score_cols].itertuples(index=False, name=None))
            idx_melhor = scores.index(max(scores))
            estilos = [
                ["background-color: #d4edda; font-weight: bold"
                 if i == idx_melhor else "" for _ in df.columns]
                for i in range(len(df))
            ]
            return pd.DataFrame(estilos, index=df.index, columns=df.columns)

        st.dataframe(
            df_scores.style.apply(destacar_melhor, axis=None),
            use_container_width=True
        )

    # --- Escala e contadores ---
    st.subheader("📅 Escala")
    st.dataframe(df, use_container_width=True)

    st.subheader("📊 Contador de Plantões")
    st.dataframe(df_contador, use_container_width=True)

    # --- Painel de criticidade ---
    st.subheader("🚨 Índice de Criticidade por Pessoa")
    df_crit = pd.DataFrame([
        {"Nome": nome, "Dias de Férias": dias}
        for nome, dias in sorted(
            criticidade_ferias.items(), key=lambda x: x[1], reverse=True
        )
        if dias > 0
    ])
    if not df_crit.empty:
        st.dataframe(df_crit, use_container_width=True)
    else:
        st.info("Nenhuma pessoa com férias marcadas no mês.")

    # --- Exportar Excel formatado ---
    st.subheader("📥 Exportar Planilha")
    excel_buf = gerar_excel(melhor_agenda, melhor_datas, mapeamento_regioes, config_regioes)
    st.download_button(
        label="⬇️ Baixar planilha Excel (.xlsx)",
        data=excel_buf,
        file_name=f"escala_{ANO}_{MES:02d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
