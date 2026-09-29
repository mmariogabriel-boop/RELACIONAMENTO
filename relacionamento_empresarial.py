
from io import BytesIO

from pathlib import Path

import pandas as pd

import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from PIL import Image as PILImage

import plotly.express as px

import plotly.graph_objects as go

import streamlit as st

BASE_DIR = Path(__file__).resolve().parent

LOGO_PATH = BASE_DIR / "logo_mediatorie.png"

st.set_page_config(

    page_title="Relacionamento Empresarial | Mediatorie",

    page_icon="📊",

    layout="wide",

)

VERDE = "#86BC25"

GRAFITE = "#404640"

FUNDO = "#F5F7F2"

st.markdown(

    f"""

    <style>

        .stApp {{background-color:{FUNDO};}}

        .block-container {{

            padding-top:3.6rem !important;

            max-width:1500px;

        }}

        [data-testid="stSidebar"] {{

            background:linear-gradient(180deg,#FFF 0%,#F0F4EA 100%);

        }}

        [data-testid="stMetric"] {{

            background:#FFF;

            border:1px solid #E2E8DC;

            border-left:5px solid {VERDE};

            border-radius:12px;

            padding:14px;

        }}

    </style>

    """,

    unsafe_allow_html=True,

)

EMPRESA_ALVO = 1000

TIPO_MOVIMENTACAO = "MOVIMENTAÇÃO"

DATA_CORTE = pd.Timestamp("2025-12-01")

CANCELAMENTOS_VALIDOS = [

    "SEM JUSTA CAUSA",

    "JUSTA CAUSA",

    "A PEDIDO DO BENEFICIÁRIO (RN561)",

    "BENEFÍCIO DEMITIDO/APOSENTADO (RN279)",

    "DESLIGAMENTO DA EMPRESA (RN279)",

    "INADIMPLENTE",

]

COLUNAS = [

    "Empresa",

    "Tipo Movimentação",

    "Data Início Beneficiário",

    "Data Inclusão Vida",

    "Descrição Cancelamento",

    "Descrição Status Atual Beneficiário",

    "Descrição Beneficiário",

    "Id Acomodação",

    "Qt. Vidas Empresa",

    "Grupo Econômico",

]

@st.cache_data(show_spinner=False)

def ler_excel(nome, conteudo):

    del nome

    return pd.read_excel(BytesIO(conteudo))

@st.cache_data(show_spinner=False)
def processar_visitas(nome, conteudo):
    bruto = pd.read_excel(BytesIO(conteudo), header=None)

    linha_cabecalho = None
    limite = min(len(bruto), 20)

    for i in range(limite):
        valores = (
            bruto.iloc[i]
            .astype(str)
            .str.strip()
            .str.upper()
            .tolist()
        )

        if "CÓDIGO TAREFA AUVO" in valores and "DATA" in valores:
            linha_cabecalho = i
            break

    if linha_cabecalho is None:
        raise ValueError(
            f'Não foi possível localizar o cabeçalho no arquivo "{nome}".'
        )

    visitas = pd.read_excel(
        BytesIO(conteudo),
        header=linha_cabecalho,
    )

    visitas.columns = (
        visitas.columns
        .astype(str)
        .str.strip()
    )

    obrigatorias = [
        "Data",
        "Tipo da tarefa",
        "Responsável",
        "Empresa",
        "Carteira",
    ]

    faltantes = [
        coluna
        for coluna in obrigatorias
        if coluna not in visitas.columns
    ]

    if faltantes:
        raise KeyError(
            f'Arquivo "{nome}" sem as colunas: '
            + ", ".join(faltantes)
        )

    visitas["Data"] = pd.to_datetime(
        visitas["Data"],
        errors="coerce",
        dayfirst=True,
    )

    # Remove linhas de totais/rodapé do relatório.
    visitas = (
        visitas.loc[visitas["Data"].notna()]
        .copy()
        .reset_index(drop=True)
    )

    for coluna in [
        "Tipo da tarefa",
        "Responsável",
        "Empresa",
        "Carteira",
        "Orientações",
    ]:
        if coluna in visitas.columns:
            visitas[coluna] = (
                visitas[coluna]
                .astype("string")
                .str.strip()
            )

    visitas["Responsável"] = (
        visitas["Responsável"]
        .fillna("Não informado")
        .replace("", "Não informado")
    )

    visitas["Empresa"] = (
        visitas["Empresa"]
        .fillna("Não informado")
        .replace("", "Não informado")
    )

    visitas["Carteira"] = (
        visitas["Carteira"]
        .fillna("Não informado")
        .replace("", "Não informado")
    )

    tipo_upper = (
        visitas["Tipo da tarefa"]
        .fillna("Não informado")
        .str.upper()
    )

    visitas["Categoria"] = visitas["Tipo da tarefa"].fillna(
        "Não informado"
    )

    visitas.loc[
        tipo_upper.str.contains("PALESTRA", na=False),
        "Categoria",
    ] = "PALESTRA"

    visitas.loc[
        tipo_upper.str.contains("VISITA", na=False),
        "Categoria",
    ] = "VISITA"

    visitas.loc[
        tipo_upper.str.contains("DEMANDA INTERNA", na=False),
        "Categoria",
    ] = "DEMANDA INTERNA"

    visitas["Competencia_Data"] = (
        visitas["Data"]
        .dt.to_period("M")
        .dt.to_timestamp()
    )

    visitas["Competencia"] = (
        visitas["Competencia_Data"]
        .dt.strftime("%m/%Y")
    )

    visitas["Arquivo Origem"] = nome

    def horas_decimal(valor):
        if pd.isna(valor):
            return 0.0

        if hasattr(valor, "hour"):
            return (
                float(valor.hour)
                + float(valor.minute) / 60
                + float(valor.second) / 3600
            )

        if isinstance(valor, pd.Timedelta):
            return valor.total_seconds() / 3600

        texto = str(valor).strip()

        try:
            partes = texto.split(":")
            if len(partes) >= 2:
                horas = float(partes[0])
                minutos = float(partes[1])
                segundos = float(partes[2]) if len(partes) > 2 else 0
                return horas + minutos / 60 + segundos / 3600
        except Exception:
            pass

        return 0.0

    if "Horas" in visitas.columns:
        visitas["Horas Decimal"] = (
            visitas["Horas"]
            .apply(horas_decimal)
        )
    else:
        visitas["Horas Decimal"] = 0.0

    return visitas

def consolidar_visitas(arquivos):
    bases = []
    erros = []

    for arquivo_visita in arquivos:
        try:
            base = processar_visitas(
                arquivo_visita.name,
                arquivo_visita.getvalue(),
            )
            bases.append(base)
        except Exception as erro:
            erros.append(
                f"{arquivo_visita.name}: {erro}"
            )

    if not bases:
        return pd.DataFrame(), erros

    consolidado = (
        pd.concat(bases, ignore_index=True)
        .reset_index(drop=True)
    )

    return consolidado, erros

def grafico_visitas_responsavel(base):
    dados = (
        base.groupby("Responsável", dropna=False)
        .size()
        .reset_index(name="Visitas")
        .sort_values("Visitas")
    )

    fig = px.bar(
        dados,
        x="Visitas",
        y="Responsável",
        orientation="h",
        text="Visitas",
    )

    fig.update_traces(
        marker_color=VERDE,
        textposition="outside",
    )

    fig.update_layout(
        title="Quantidade de visitas por responsável",
        height=max(420, 55 * len(dados) + 180),
        paper_bgcolor="white",
        plot_bgcolor="white",
        font=dict(family="Arial", color=GRAFITE),
        xaxis_title="Quantidade de visitas",
        yaxis_title=None,
        showlegend=False,
        margin=dict(l=30, r=40, t=80, b=50),
    )

    return fig

def grafico_visitas_competencia(base):
    dados = (
        base.groupby(
            ["Competencia_Data", "Responsável"],
            dropna=False,
        )
        .size()
        .reset_index(name="Visitas")
        .sort_values("Competencia_Data")
    )

    dados["Competência"] = (
        dados["Competencia_Data"]
        .dt.strftime("%m/%Y")
    )

    fig = px.bar(
        dados,
        x="Competência",
        y="Visitas",
        color="Responsável",
        barmode="group",
        text="Visitas",
    )

    fig.update_traces(textposition="outside")

    fig.update_layout(
        title="Visitas por competência e responsável",
        height=500,
        paper_bgcolor="white",
        plot_bgcolor="white",
        font=dict(family="Arial", color=GRAFITE),
        xaxis_title="Competência",
        yaxis_title="Quantidade de visitas",
        legend_title=None,
        margin=dict(l=30, r=30, t=80, b=50),
    )

    return fig

def grafico_visitas_categoria(base):
    dados = (
        base.groupby("Categoria", dropna=False)
        .size()
        .reset_index(name="Visitas")
        .sort_values("Visitas", ascending=False)
    )

    fig = px.bar(
        dados,
        x="Categoria",
        y="Visitas",
        text="Visitas",
    )

    fig.update_traces(
        marker_color=VERDE,
        textposition="outside",
    )

    fig.update_layout(
        title="Visitas por tipo de atividade",
        height=450,
        paper_bgcolor="white",
        plot_bgcolor="white",
        font=dict(family="Arial", color=GRAFITE),
        xaxis_title=None,
        yaxis_title="Quantidade",
        showlegend=False,
        margin=dict(l=30, r=30, t=80, b=90),
    )

    return fig

def grafico_visitas_carteira(base):
    dados = (
        base.groupby("Carteira", dropna=False)
        .size()
        .reset_index(name="Visitas")
        .sort_values("Visitas", ascending=False)
    )

    fig = px.bar(
        dados,
        x="Carteira",
        y="Visitas",
        text="Visitas",
    )

    fig.update_traces(
        marker_color=VERDE,
        textposition="outside",
    )

    fig.update_layout(
        title="Visitas por carteira",
        height=450,
        paper_bgcolor="white",
        plot_bgcolor="white",
        font=dict(family="Arial", color=GRAFITE),
        xaxis_title="Carteira",
        yaxis_title="Quantidade",
        showlegend=False,
        margin=dict(l=30, r=30, t=80, b=60),
    )

    return fig

def grafico_visitas_dia(base):
    dados = (
        base.groupby(
            ["Data", "Responsável"],
            dropna=False,
        )
        .size()
        .reset_index(name="Visitas")
        .sort_values("Data")
    )

    fig = px.line(
        dados,
        x="Data",
        y="Visitas",
        color="Responsável",
        markers=True,
    )

    fig.update_layout(
        title="Evolução diária das visitas",
        height=450,
        paper_bgcolor="white",
        plot_bgcolor="white",
        font=dict(family="Arial", color=GRAFITE),
        xaxis_title="Data",
        yaxis_title="Quantidade de visitas",
        legend_title=None,
        margin=dict(l=30, r=30, t=80, b=50),
    )

    return fig

@st.cache_data(show_spinner=False)

def processar(nome, conteudo):

    df = ler_excel(nome, conteudo).copy()

    df.columns = df.columns.astype(str).str.strip()

    faltantes = [c for c in COLUNAS if c not in df.columns]

    if faltantes:

        raise KeyError(

            "Faltam as colunas: " + ", ".join(faltantes)

        )

    # Regras iniciais

    df = df[

        df["Empresa"].eq(EMPRESA_ALVO)

        & df["Tipo Movimentação"].eq(TIPO_MOVIMENTACAO)

    ].copy()

    df["Data Início Beneficiário"] = pd.to_datetime(

        df["Data Início Beneficiário"],

        errors="coerce",

        dayfirst=True,

    )

    df["Data Inclusão Vida"] = pd.to_datetime(

        df["Data Inclusão Vida"],

        errors="coerce",

        dayfirst=True,

    )

    df["Diferenca"] = (

        df["Data Inclusão Vida"]

        - df["Data Início Beneficiário"]

    ).dt.days

    # Descrição permitida OU campo vazio

    df = df[

        df["Descrição Cancelamento"].isin(CANCELAMENTOS_VALIDOS)

        | df["Descrição Cancelamento"].isna()

    ].copy()

    # Conferência

    nao_corretos = df[

        df["Diferenca"].gt(0)

        & df["Descrição Status Atual Beneficiário"].eq("Ativado")

    ][

        [

            "Descrição Beneficiário",

            "Data Início Beneficiário",

            "Data Inclusão Vida",

            "Diferenca",

        ]

    ].copy()

    # Base válida

    certo = (
        df.loc[df["Diferenca"].le(0)]
        .copy()
        .reset_index(drop=True)
    )

    certo["Competencia"] = (

        certo["Data Inclusão Vida"].dt.strftime("%m/%Y")

    )

    certo["Competencia_Data"] = pd.to_datetime(

        certo["Competencia"],

        format="%m/%Y",

        errors="coerce",

    )

    certo = certo[

        certo["Competencia_Data"].gt(DATA_CORTE)

    ].copy()

    return df, certo, nao_corretos

def separar(base):

    saude = base[

        ~base["Id Acomodação"].isin(["ODO", "SEM"])

    ].copy()

    odonto = base[

        base["Id Acomodação"].eq("ODO")

    ].copy()

    return saude, odonto

def agrupar_carteira(base):

    resultado = (

        base.groupby(

            ["Competencia", "Qt. Vidas Empresa"],

            dropna=False,

        )

        .agg({"Descrição Beneficiário": "count"})

        .reset_index()

    )

    resultado["Competencia"] = pd.to_datetime(

        resultado["Competencia"],

        format="%m/%Y",

        errors="coerce",

    )

    return resultado.sort_values(

        ["Competencia", "Qt. Vidas Empresa"]

    ).reset_index(drop=True)

def agrupar_grupo_economico(base):
    dados = base.copy()

    grupo = dados["Grupo Econômico"].astype("string").str.strip()

    dados = dados[
        grupo.notna()
        & grupo.ne("")
        & ~grupo.str.casefold().eq("não informado")
    ].copy()

    if dados.empty:
        return pd.DataFrame(
            columns=[
                "Competencia",
                "Grupo Econômico",
                "Descrição Beneficiário",
            ]
        )

    resultado = (
        dados.groupby(
            ["Competencia", "Grupo Econômico"],
            dropna=False,
        )
        .agg({"Descrição Beneficiário": "count"})
        .reset_index()
    )

    resultado["Competencia"] = pd.to_datetime(
        resultado["Competencia"],
        format="%m/%Y",
        errors="coerce",
    )

    resultado["Grupo Econômico"] = (
        resultado["Grupo Econômico"]
        .astype("string")
        .str.strip()
    )

    return resultado.sort_values(
        ["Competencia", "Descrição Beneficiário"],
        ascending=[True, False],
    ).reset_index(drop=True)


def montar_tabela_grupo_economico(agrupado):
    """Monta uma matriz Grupo Econômico x Competência para exibição no dashboard/PDF."""
    if agrupado is None or agrupado.empty:
        return pd.DataFrame(columns=["Grupo Econômico", "Total"])

    dados = agrupado.copy()

    dados = dados[
        dados["Grupo Econômico"].notna()
        & dados["Grupo Econômico"].astype("string").str.strip().ne("")
        & ~dados["Grupo Econômico"]
        .astype("string")
        .str.strip()
        .str.casefold()
        .eq("não informado")
    ].copy()

    if dados.empty:
        return pd.DataFrame(columns=["Grupo Econômico", "Total"])

    dados["Competência"] = dados["Competencia"].dt.strftime("%m/%Y")

    tabela = dados.pivot_table(
        index="Grupo Econômico",
        columns="Competência",
        values="Descrição Beneficiário",
        aggfunc="sum",
        fill_value=0,
    )

    competencias = sorted(
        tabela.columns.tolist(),
        key=lambda valor: pd.to_datetime(
            valor,
            format="%m/%Y",
            errors="coerce",
        ),
    )

    tabela = tabela[competencias]
    tabela["Total"] = tabela.sum(axis=1)

    tabela = (
        tabela.sort_values("Total", ascending=False)
        .reset_index()
    )

    for coluna in tabela.columns:
        if coluna != "Grupo Econômico":
            tabela[coluna] = (
                pd.to_numeric(tabela[coluna], errors="coerce")
                .fillna(0)
                .astype(int)
            )

    return tabela


def agrupar_competencia(base):

    return (

        base.groupby("Competencia")

        .agg({"Descrição Beneficiário": "count"})

        .reset_index()

        .rename(

            columns={"Descrição Beneficiário": "Quantidade"}

        )

    )

def criar_metas_2026():

    competencias = pd.date_range(

        start="2026-01-01",

        end="2026-12-01",

        freq="MS",

    )

    metas = pd.DataFrame(

        {

            "Competencia": competencias,

            "Meta Saúde": [

                870, 805, 833, 870, 944, 913,

                967, 953, 994, 996, 998, 993,

            ],

            "Meta Odonto": [

                227, 230, 235, 259, 257, 270,

                287, 308, 302, 315, 325, 337,

            ],

        }

    )

    return metas

def montar_comparativo_meta(

    base: pd.DataFrame,

    metas: pd.DataFrame,

    coluna_meta: str,

    competencias_selecionadas: list[str],

) -> pd.DataFrame:

    competencias_datas = pd.to_datetime(

        competencias_selecionadas,

        format="%m/%Y",

        errors="coerce",

    )

    base_meta = metas[

        metas["Competencia"].isin(

            competencias_datas

        )

    ].copy()

    realizado = (

        base

        .groupby("Competencia")

        .agg(

            {

                "Descrição Beneficiário": "count",

            }

        )

        .reset_index()

        .rename(

            columns={

                "Descrição Beneficiário": "Realizado",

            }

        )

    )

    realizado["Competencia"] = pd.to_datetime(

        realizado["Competencia"],

        format="%m/%Y",

        errors="coerce",

    )

    comparativo = base_meta.merge(

        realizado,

        how="left",

        on="Competencia",

    )

    comparativo["Realizado"] = (

        comparativo["Realizado"]

        .fillna(0)

        .astype(int)

    )

    comparativo["Diferença"] = (

        comparativo["Realizado"]

        - comparativo[coluna_meta]

    )

    comparativo["Atingimento (%)"] = (

        comparativo["Realizado"]

        .div(comparativo[coluna_meta])

        .mul(100)

    )

    return (

        comparativo

        .sort_values("Competencia")

        .reset_index(drop=True)

    )

def grafico_realizado_meta(

    comparativo: pd.DataFrame,

    coluna_meta: str,

    titulo: str,

) -> go.Figure:

    dados = comparativo.copy()

    dados["Competencia_Label"] = (

        dados["Competencia"]

        .dt.strftime("%m/%Y")

    )

    fig = go.Figure()

    fig.add_bar(

        x=dados["Competencia_Label"],

        y=dados[coluna_meta],

        name="Meta",

        marker_color="#8B908D",

        text=dados[coluna_meta],

        textposition="outside",

    )

    fig.add_bar(

        x=dados["Competencia_Label"],

        y=dados["Realizado"],

        name="Realizado",

        marker_color=VERDE,

        text=dados["Realizado"],

        textposition="outside",

    )

    fig.update_layout(

        title=titulo,

        barmode="group",

        height=470,

        paper_bgcolor="white",

        plot_bgcolor="white",

        margin=dict(

            l=35,

            r=25,

            t=80,

            b=55,

        ),

        legend=dict(

            orientation="h",

            y=1.12,

            x=1,

            xanchor="right",

        ),

        font=dict(

            family="Arial",

            color=GRAFITE,

        ),

    )

    fig.update_xaxes(

        title="Competência",

        showgrid=False,

    )

    fig.update_yaxes(

        title="Quantidade",

        gridcolor="#E8EDE4",

        zeroline=False,

    )

    return fig

def percentual_br(valor: float) -> str:

    return f"{valor:.1f}%".replace(".", ",")

def grafico_geral(saude, odonto):

    s = agrupar_competencia(saude)

    s["Segmento"] = "Saúde"

    o = agrupar_competencia(odonto)

    o["Segmento"] = "Odonto"

    dados = pd.concat([s, o], ignore_index=True)

    ordem = (

        pd.to_datetime(

            dados["Competencia"],

            format="%m/%Y",

            errors="coerce",

        )

    )

    dados = dados.assign(_ordem=ordem).sort_values("_ordem")

    fig = px.bar(

        dados,

        x="Competencia",

        y="Quantidade",

        color="Segmento",

        barmode="group",

        text="Quantidade",

        color_discrete_map={

            "Saúde": VERDE,

            "Odonto": "#6F7772",

        },

    )

    fig.update_traces(textposition="outside")

    fig.update_layout(

        title="Movimentações por competência — Saúde x Odonto",

        height=450,

        paper_bgcolor="white",

        plot_bgcolor="white",

        font=dict(family="Arial", color=GRAFITE),

        legend_title=None,

    )

    return fig

def grafico_carteira(agrupado, titulo):

    dados = agrupado.copy()

    dados["Competência"] = dados["Competencia"].dt.strftime("%m/%Y")

    dados["Qt. Vidas Empresa"] = (

        dados["Qt. Vidas Empresa"]

        .astype("string")

        .fillna("Não informado")

    )

    fig = px.bar(

        dados,

        x="Competência",

        y="Descrição Beneficiário",

        color="Qt. Vidas Empresa",

        barmode="stack",

        text="Descrição Beneficiário",

    )

    fig.update_layout(

        title=titulo,

        height=530,

        paper_bgcolor="white",

        plot_bgcolor="white",

        font=dict(family="Arial", color=GRAFITE),

        yaxis_title="Quantidade de beneficiários",

        legend_title="Qt. Vidas Empresa",

    )

    return fig

def ranking_carteira(agrupado, titulo):

    ranking = (

        agrupado.groupby("Qt. Vidas Empresa", dropna=False)

        ["Descrição Beneficiário"]

        .sum()

        .reset_index()

        .rename(columns={"Descrição Beneficiário": "Quantidade"})

        .sort_values("Quantidade", ascending=False)

        .head(15)

        .sort_values("Quantidade")

    )

    ranking["Qt. Vidas Empresa"] = (

        ranking["Qt. Vidas Empresa"]

        .astype("string")

        .fillna("Não informado")

    )

    fig = px.bar(

        ranking,

        x="Quantidade",

        y="Qt. Vidas Empresa",

        orientation="h",

        text="Quantidade",

    )

    fig.update_traces(

        marker_color=VERDE,

        textposition="outside",

    )

    fig.update_layout(

        title=titulo,

        height=470,

        paper_bgcolor="white",

        plot_bgcolor="white",

        font=dict(family="Arial", color=GRAFITE),

        xaxis_title="Quantidade de beneficiários",

    )

    return fig



def obter_logo_bytes():
    """Retorna a logo fixa da Mediatorie disponível na pasta do projeto."""
    if LOGO_PATH.exists():
        try:
            return LOGO_PATH.read_bytes()
        except Exception:
            return None

    return None


def _formatar_inteiro_br(valor):
    try:
        return f"{int(round(float(valor))):,}".replace(",", ".")
    except Exception:
        return "0"


def _formatar_decimal_br(valor, casas=1):
    try:
        return (
            f"{float(valor):,.{casas}f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )
    except Exception:
        return "0,0"


def _formatar_percentual_br(valor):
    try:
        return f"{float(valor):.1f}%".replace(".", ",")
    except Exception:
        return "0,0%"


def _resumo_filtro(valores, max_itens=5):
    valores = [str(v) for v in valores if pd.notna(v)]
    if not valores:
        return "Nenhum"
    if len(valores) <= max_itens:
        return ", ".join(valores)
    return f"{valores[0]} a {valores[-1]} ({len(valores)} selecionados)"


def _texto_curto(valor, limite=88):
    texto = str(valor)
    if len(texto) <= limite:
        return texto
    return texto[: limite - 3] + "..."


def _quebrar_texto_pdf(valor, limite=88):
    palavras = str(valor).split()
    if not palavras:
        return ""

    linhas = []
    atual = palavras[0]
    for palavra in palavras[1:]:
        candidato = f"{atual} {palavra}"
        if len(candidato) <= limite:
            atual = candidato
        else:
            linhas.append(atual)
            atual = palavra
    linhas.append(atual)
    return "\n".join(linhas)


# =========================================================
# PDF - PADRÃO VISUAL DO RELATÓRIO COMERCIAL
# =========================================================

PDF_VERDE_TEXTO = "#558A12"
PDF_VERDE_CLARO = "#EEF5E3"
PDF_CINZA = "#747B77"
PDF_BORDA = "#D5DDD0"
PDF_LINHA = "#80BD16"
PDF_LINHA_ALTERNADA = "#F7F9F5"


def _nova_pagina_pdf(pagina):
    """Cria uma página A4 retrato com cabeçalho e rodapé padrão Mediatorie."""
    fig = plt.figure(figsize=(8.27, 11.69), facecolor="white")

    fig.text(
        0.075,
        0.972,
        "MEDIATORIE | DASHBOARD RELACIONAMENTO",
        fontsize=7.8,
        fontweight="bold",
        color=PDF_VERDE_TEXTO,
        va="top",
    )

    fig.add_artist(
        plt.Line2D(
            [0.075, 0.925],
            [0.957, 0.957],
            transform=fig.transFigure,
            color=PDF_LINHA,
            linewidth=1.0,
        )
    )

    fig.text(
        0.075,
        0.028,
        "Mediatorie Administradora de Benefícios",
        fontsize=6.8,
        color=PDF_CINZA,
        va="bottom",
    )
    fig.text(
        0.925,
        0.028,
        f"Página {pagina}",
        fontsize=6.8,
        color=PDF_CINZA,
        ha="right",
        va="bottom",
    )

    return fig


def _titulo_pagina_pdf(fig, titulo, subtitulo=None):
    fig.text(
        0.085,
        0.925,
        titulo,
        fontsize=15.5,
        fontweight="bold",
        color=PDF_VERDE_TEXTO,
        va="top",
    )

    if subtitulo:
        fig.text(
            0.085,
            0.897,
            subtitulo,
            fontsize=8.2,
            color=PDF_CINZA,
            va="top",
        )


def _logo_pdf(fig, logo_bytes, bbox=(0.33, 0.725, 0.34, 0.16)):
    if not logo_bytes:
        return

    try:
        imagem = PILImage.open(BytesIO(logo_bytes)).convert("RGBA")
        ax = fig.add_axes(list(bbox))
        ax.imshow(imagem)
        ax.axis("off")
    except Exception:
        pass


def _tabela_kpis_pdf(fig, rotulos, valores, bbox=(0.085, 0.805, 0.83, 0.074)):
    ax = fig.add_axes(list(bbox))
    ax.axis("off")

    tabela = ax.table(
        cellText=[[str(v) for v in valores]],
        colLabels=[str(r) for r in rotulos],
        cellLoc="center",
        colLoc="center",
        bbox=[0, 0, 1, 1],
    )
    tabela.auto_set_font_size(False)
    tabela.set_fontsize(7.7)

    for (linha, coluna), celula in tabela.get_celld().items():
        celula.set_edgecolor(PDF_BORDA)
        celula.set_linewidth(0.55)
        celula.PAD = 0.06

        if linha == 0:
            celula.set_facecolor(PDF_VERDE_CLARO)
            celula.get_text().set_color(PDF_VERDE_TEXTO)
            celula.get_text().set_weight("bold")
        else:
            celula.set_facecolor("white")
            celula.get_text().set_color("black")
            celula.get_text().set_weight("bold")

    return tabela


def _tabela_chave_valor_pdf(fig, itens, bbox=(0.11, 0.49, 0.78, 0.22)):
    ax = fig.add_axes(list(bbox))
    ax.axis("off")

    dados = [[str(k), _texto_curto(v)] for k, v in itens]
    tabela = ax.table(
        cellText=dados,
        cellLoc="left",
        colWidths=[0.29, 0.71],
        bbox=[0, 0, 1, 1],
    )
    tabela.auto_set_font_size(False)
    tabela.set_fontsize(7.5)

    for (linha, coluna), celula in tabela.get_celld().items():
        celula.set_edgecolor(PDF_BORDA)
        celula.set_linewidth(0.55)
        celula.PAD = 0.045

        if coluna == 0:
            celula.set_facecolor(PDF_VERDE_CLARO)
            celula.get_text().set_color(PDF_VERDE_TEXTO)
            celula.get_text().set_weight("bold")
        else:
            celula.set_facecolor("white")
            celula.get_text().set_color(GRAFITE)

    return tabela


def _tabela_dataframe_pdf(
    fig,
    dados,
    bbox,
    font_size=6.8,
    col_widths=None,
    alinhar_primeira_esquerda=True,
    linhas_maximas=None,
):
    """Desenha tabela no padrão visual do relatório comercial."""
    ax = fig.add_axes(list(bbox))
    ax.axis("off")

    if dados is None or dados.empty:
        ax.text(
            0.5,
            0.5,
            "Sem dados para o período selecionado.",
            ha="center",
            va="center",
            fontsize=9,
            color=PDF_CINZA,
        )
        return None

    tabela_df = dados.copy()
    if linhas_maximas is not None:
        tabela_df = tabela_df.head(linhas_maximas)

    tabela_df = tabela_df.fillna("")

    tabela = ax.table(
        cellText=tabela_df.astype(str).values.tolist(),
        colLabels=[str(c) for c in tabela_df.columns],
        cellLoc="left",
        colLoc="left",
        colWidths=col_widths,
        bbox=[0, 0, 1, 1],
    )
    tabela.auto_set_font_size(False)
    tabela.set_fontsize(font_size)

    for (linha, coluna), celula in tabela.get_celld().items():
        celula.set_edgecolor(PDF_BORDA)
        celula.set_linewidth(0.5)
        celula.PAD = 0.045

        if linha == 0:
            celula.set_facecolor(PDF_LINHA)
            celula.get_text().set_color("white")
            celula.get_text().set_weight("bold")
        else:
            celula.set_facecolor(
                "white" if linha % 2 else PDF_LINHA_ALTERNADA
            )
            celula.get_text().set_color(GRAFITE)

        if alinhar_primeira_esquerda and coluna == 0:
            celula.get_text().set_ha("left")
        elif coluna > 0:
            celula.get_text().set_ha("left")

    return tabela


def _estilo_eixo_pdf(ax):
    ax.set_facecolor("white")
    ax.grid(axis="y", color="#E9ECE8", linewidth=0.6)
    ax.set_axisbelow(True)
    ax.tick_params(axis="both", labelsize=6.8, colors="#222222")
    ax.spines["top"].set_color("#444444")
    ax.spines["right"].set_color("#444444")
    ax.spines["left"].set_color("#444444")
    ax.spines["bottom"].set_color("#444444")
    ax.spines["top"].set_linewidth(0.55)
    ax.spines["right"].set_linewidth(0.55)
    ax.spines["left"].set_linewidth(0.55)
    ax.spines["bottom"].set_linewidth(0.55)


def _rotular_barras_pdf(ax, fontsize=5.8):
    for container in ax.containers:
        try:
            ax.bar_label(
                container,
                fmt=lambda v: _formatar_inteiro_br(v),
                padding=2,
                fontsize=fontsize,
                color="#222222",
            )
        except Exception:
            pass


def _grafico_geral_pdf(ax, saude, odonto):
    s = agrupar_competencia(saude).rename(columns={"Quantidade": "Saúde"})
    o = agrupar_competencia(odonto).rename(columns={"Quantidade": "Odonto"})
    dados = s.merge(o, on="Competencia", how="outer").fillna(0)

    if dados.empty:
        ax.text(0.5, 0.5, "Sem dados", ha="center", va="center")
        ax.set_axis_off()
        return

    dados["_ordem"] = pd.to_datetime(
        dados["Competencia"],
        format="%m/%Y",
        errors="coerce",
    )
    dados = dados.sort_values("_ordem")

    labels = dados["Competencia"].tolist()
    x = list(range(len(labels)))
    largura = 0.36

    ax.bar(
        [i - largura / 2 for i in x],
        dados["Saúde"].tolist(),
        largura,
        label="Saúde",
        color=VERDE,
    )
    ax.bar(
        [i + largura / 2 for i in x],
        dados["Odonto"].tolist(),
        largura,
        label="Odonto",
        color="#777E7A",
    )

    ax.set_title(
        "Produção por competência",
        loc="left",
        fontsize=8.5,
        fontweight="bold",
        pad=4,
    )
    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=35, ha="right", fontsize=6.5)
    ax.set_ylabel("Beneficiários", fontsize=7)
    ax.legend(
        frameon=False,
        ncol=2,
        fontsize=6.5,
        loc="upper center",
        bbox_to_anchor=(0.52, 1.13),
    )
    _estilo_eixo_pdf(ax)
    _rotular_barras_pdf(ax)


def _grafico_barras_meta_pdf(ax, comparativo, coluna_meta, titulo):
    if comparativo is None or comparativo.empty:
        ax.text(0.5, 0.5, "Sem dados", ha="center", va="center")
        ax.set_axis_off()
        return

    dados = comparativo.copy().sort_values("Competencia")
    labels = dados["Competencia"].dt.strftime("%m/%Y").tolist()
    x = list(range(len(labels)))
    largura = 0.36

    ax.bar(
        [i - largura / 2 for i in x],
        dados[coluna_meta].tolist(),
        width=largura,
        label="Meta",
        color="#969B98",
    )
    ax.bar(
        [i + largura / 2 for i in x],
        dados["Realizado"].tolist(),
        width=largura,
        label="Realizado",
        color=VERDE,
    )

    ax.set_title(titulo, loc="left", fontsize=8.5, fontweight="bold", pad=4)
    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=35, ha="right", fontsize=6.5)
    ax.set_ylabel("Quantidade", fontsize=7)
    ax.legend(
        frameon=False,
        ncol=2,
        fontsize=6.5,
        loc="upper center",
        bbox_to_anchor=(0.52, 1.13),
    )
    _estilo_eixo_pdf(ax)
    _rotular_barras_pdf(ax, fontsize=5.4)


def _grafico_stack_competencia_pdf(ax, agrupado, categoria, titulo, top_n=8):
    if agrupado is None or agrupado.empty:
        ax.text(0.5, 0.5, "Sem dados", ha="center", va="center")
        ax.set_axis_off()
        return

    dados = agrupado.copy()
    dados[categoria] = dados[categoria].astype("string").fillna("Não informado")

    principais = (
        dados.groupby(categoria, dropna=False)["Descrição Beneficiário"]
        .sum()
        .sort_values(ascending=False)
        .head(top_n)
        .index
        .astype(str)
        .tolist()
    )

    dados = dados[dados[categoria].astype(str).isin(principais)].copy()
    pivot = dados.pivot_table(
        index="Competencia",
        columns=categoria,
        values="Descrição Beneficiário",
        aggfunc="sum",
        fill_value=0,
    ).sort_index()

    if pivot.empty:
        ax.text(0.5, 0.5, "Sem dados", ha="center", va="center")
        ax.set_axis_off()
        return

    x = list(range(len(pivot.index)))
    base = None
    for coluna in pivot.columns:
        valores = pivot[coluna].to_numpy()
        ax.bar(x, valores, bottom=base, label=str(coluna))
        base = valores if base is None else base + valores

    ax.set_title(titulo, loc="left", fontsize=8.5, fontweight="bold", pad=4)
    ax.set_xticks(x)
    ax.set_xticklabels(
        [d.strftime("%m/%Y") for d in pivot.index],
        rotation=35,
        ha="right",
        fontsize=6.5,
    )
    ax.set_ylabel("Beneficiários", fontsize=7)
    ax.legend(
        frameon=False,
        fontsize=5.3,
        ncol=4,
        loc="upper center",
        bbox_to_anchor=(0.5, 1.17),
    )
    _estilo_eixo_pdf(ax)


def _grafico_visitas_comp_pdf(ax, visitas):
    if visitas is None or visitas.empty:
        ax.text(0.5, 0.5, "Sem dados", ha="center", va="center")
        ax.set_axis_off()
        return

    dados = (
        visitas.groupby(["Competencia_Data", "Responsável"], dropna=False)
        .size()
        .reset_index(name="Visitas")
    )
    pivot = dados.pivot_table(
        index="Competencia_Data",
        columns="Responsável",
        values="Visitas",
        aggfunc="sum",
        fill_value=0,
    ).sort_index()

    x = list(range(len(pivot.index)))
    n = max(len(pivot.columns), 1)
    largura = min(0.8 / n, 0.25)
    deslocamentos = [(i - (n - 1) / 2) * largura for i in range(n)]

    for deslocamento, coluna in zip(deslocamentos, pivot.columns):
        ax.bar(
            [i + deslocamento for i in x],
            pivot[coluna].to_numpy(),
            width=largura,
            label=str(coluna),
        )

    ax.set_title(
        "Visitas por competência e responsável",
        loc="left",
        fontsize=8.5,
        fontweight="bold",
        pad=4,
    )
    ax.set_xticks(x)
    ax.set_xticklabels(
        [d.strftime("%m/%Y") for d in pivot.index],
        rotation=35,
        ha="right",
        fontsize=6.5,
    )
    ax.set_ylabel("Visitas", fontsize=7)
    ax.legend(
        frameon=False,
        fontsize=5.8,
        ncol=3,
        loc="upper center",
        bbox_to_anchor=(0.5, 1.15),
    )
    _estilo_eixo_pdf(ax)
    _rotular_barras_pdf(ax, fontsize=5.2)


def _grafico_barra_horizontal_pdf(ax, serie, titulo, xlabel):
    if serie is None or serie.empty:
        ax.text(0.5, 0.5, "Sem dados", ha="center", va="center")
        ax.set_axis_off()
        return

    serie = serie.sort_values().tail(12)
    barras = ax.barh(serie.index.astype(str), serie.values, color=VERDE)
    ax.set_title(titulo, loc="left", fontsize=8.5, fontweight="bold", pad=4)
    ax.set_xlabel(xlabel, fontsize=7)
    ax.grid(axis="x", color="#E9ECE8", linewidth=0.6)
    ax.set_axisbelow(True)
    ax.tick_params(axis="y", labelsize=5.7)
    ax.tick_params(axis="x", labelsize=6.2)
    for barra, valor in zip(barras, serie.values):
        ax.text(
            barra.get_width(),
            barra.get_y() + barra.get_height() / 2,
            f" {_formatar_inteiro_br(valor)}",
            va="center",
            fontsize=5.7,
        )


def _preparar_tabela_meta(comparativo, coluna_meta):
    if comparativo is None or comparativo.empty:
        return pd.DataFrame(
            columns=["Competência", "Meta", "Realizado", "Diferença", "Atingimento"]
        )

    tabela = comparativo.copy().sort_values("Competencia")
    tabela["Competência"] = tabela["Competencia"].dt.strftime("%m/%Y")
    tabela["Meta"] = tabela[coluna_meta].map(_formatar_inteiro_br)
    tabela["Realizado"] = tabela["Realizado"].map(_formatar_inteiro_br)
    tabela["Diferença"] = tabela["Diferença"].map(
        lambda v: f"{int(v):+,}".replace(",", ".")
    )
    tabela["Atingimento"] = tabela["Atingimento (%)"].map(_formatar_percentual_br)

    return tabela[
        ["Competência", "Meta", "Realizado", "Diferença", "Atingimento"]
    ]


def _preparar_ranking_porte(agrupado):
    if agrupado is None or agrupado.empty:
        return pd.DataFrame(columns=["Qt. Vidas Empresa", "Quantidade"])

    tabela = (
        agrupado.groupby("Qt. Vidas Empresa", dropna=False)["Descrição Beneficiário"]
        .sum()
        .reset_index()
        .rename(columns={"Descrição Beneficiário": "Quantidade"})
        .sort_values("Quantidade", ascending=False)
        .head(12)
    )
    tabela["Qt. Vidas Empresa"] = (
        tabela["Qt. Vidas Empresa"].astype("string").fillna("Não informado")
    )
    tabela["Quantidade"] = tabela["Quantidade"].map(_formatar_inteiro_br)
    return tabela


def _adicionar_paginas_grupo_economico_pdf(
    pdf,
    tabela,
    titulo,
    pagina,
    linhas_por_pagina=25,
):
    """Grupo Econômico somente em tabela, no padrão retrato do relatório gerencial."""
    if tabela is None or tabela.empty:
        fig = _nova_pagina_pdf(pagina)
        _titulo_pagina_pdf(fig, titulo)
        fig.text(
            0.085,
            0.84,
            "Sem grupos econômicos informados para os filtros selecionados.",
            fontsize=9,
            color=PDF_CINZA,
        )
        pdf.savefig(fig, facecolor="white")
        plt.close(fig)
        return pagina + 1

    dados = tabela.copy()
    total_partes = max(1, (len(dados) + linhas_por_pagina - 1) // linhas_por_pagina)

    for parte, inicio in enumerate(range(0, len(dados), linhas_por_pagina), start=1):
        trecho = dados.iloc[inicio : inicio + linhas_por_pagina].copy()

        for coluna in trecho.columns[1:]:
            trecho[coluna] = trecho[coluna].map(_formatar_inteiro_br)

        fig = _nova_pagina_pdf(pagina)
        subtitulo = None
        if total_partes > 1:
            subtitulo = f"Grupo Econômico x Competência - parte {parte}/{total_partes}"
        _titulo_pagina_pdf(fig, titulo, subtitulo)

        n_colunas = max(len(trecho.columns), 1)
        if n_colunas > 1:
            primeira = 0.28
            demais = (1.0 - primeira) / (n_colunas - 1)
            larguras = [primeira] + [demais] * (n_colunas - 1)
        else:
            larguras = [1.0]

        altura_tabela = min(0.62, max(0.11, 0.030 * (len(trecho) + 1)))
        y_tabela = 0.85 - altura_tabela

        _tabela_dataframe_pdf(
            fig,
            trecho,
            bbox=(0.085, y_tabela, 0.83, altura_tabela),
            font_size=5.35 if n_colunas >= 9 else 6.0,
            col_widths=larguras,
        )

        fig.text(
            0.085,
            max(0.17, y_tabela - 0.025),
            "Registros sem Grupo Econômico informado não são apresentados.",
            fontsize=7.0,
            color=PDF_CINZA,
        )

        pdf.savefig(fig, facecolor="white")
        plt.close(fig)
        pagina += 1

    return pagina


def gerar_pdf_dashboard(
    saude,
    odonto,
    saude_agru,
    odonto_agru,
    saude_grupo_eco,
    odonto_grupo_eco,
    comparativo_saude,
    comparativo_odonto,
    bruto,
    valido,
    nao_corretos,
    competencias,
    qt_vidas,
    visitas_filtradas=None,
    logo_bytes=None,
    nome_base=None,
    nomes_visitas=None,
):
    """Gera o Relatório Gerencial de Relacionamento no padrão visual do Comercial."""
    buffer = BytesIO()
    pagina = 1
    nomes_visitas = nomes_visitas or []

    with PdfPages(buffer) as pdf:
        metadata = pdf.infodict()
        metadata["Title"] = "Relatório Gerencial - Relacionamento Empresarial"
        metadata["Author"] = "Mediatorie Administradora de Benefícios"
        metadata["Subject"] = "Dashboard de Relacionamento Empresarial"

        # -------------------------------------------------
        # CAPA
        # -------------------------------------------------
        fig = _nova_pagina_pdf(pagina)
        _logo_pdf(fig, logo_bytes, bbox=(0.31, 0.745, 0.38, 0.155))

        fig.text(
            0.5,
            0.700,
            "Relatório Gerencial - Relacionamento Empresarial",
            ha="center",
            va="top",
            fontsize=18.5,
            fontweight="bold",
            color=GRAFITE,
        )
        fig.text(
            0.5,
            0.658,
            "Movimentações, metas, grupos econômicos e visitas do período selecionado.",
            ha="center",
            va="top",
            fontsize=9.5,
            color=PDF_CINZA,
        )

        if nomes_visitas:
            if len(nomes_visitas) <= 2:
                visitas_txt = ", ".join(nomes_visitas)
            else:
                visitas_txt = f"{len(nomes_visitas)} arquivos consolidados"
        else:
            visitas_txt = "Nenhum arquivo de visitas carregado"

        itens_capa = [
            ("Gerado em", pd.Timestamp.now().strftime("%d/%m/%Y %H:%M")),
            ("Base Total", nome_base or "Arquivo carregado no dashboard"),
            ("Relatórios de Visitas", visitas_txt),
            ("Competência", _resumo_filtro(competencias)),
            ("Qt. Vidas Empresa", _resumo_filtro(qt_vidas)),
        ]
        _tabela_chave_valor_pdf(
            fig,
            itens_capa,
            bbox=(0.11, 0.435, 0.78, 0.19),
        )

        fig.text(
            0.085,
            0.400,
            "O relatório utiliza exatamente o recorte de filtros aplicado no dashboard no momento da geração.",
            fontsize=8.0,
            color=GRAFITE,
        )

        pdf.savefig(fig, facecolor="white")
        plt.close(fig)
        pagina += 1

        # -------------------------------------------------
        # RESUMO EXECUTIVO
        # -------------------------------------------------
        fig = _nova_pagina_pdf(pagina)
        _titulo_pagina_pdf(fig, "Resumo executivo")

        total_saude = len(saude)
        total_odonto = len(odonto)
        total = total_saude + total_odonto
        total_incons = len(nao_corretos)

        _tabela_kpis_pdf(
            fig,
            ["Saúde", "Odonto", "Total", "Diferença > 0"],
            [
                _formatar_inteiro_br(total_saude),
                _formatar_inteiro_br(total_odonto),
                _formatar_inteiro_br(total),
                _formatar_inteiro_br(total_incons),
            ],
            bbox=(0.085, 0.805, 0.83, 0.075),
        )

        ax = fig.add_axes([0.14, 0.535, 0.73, 0.225])
        _grafico_geral_pdf(ax, saude, odonto)

        pdf.savefig(fig, facecolor="white")
        plt.close(fig)
        pagina += 1

        # -------------------------------------------------
        # META - SAÚDE
        # -------------------------------------------------
        for nome_segmento, comparativo, coluna_meta in [
            ("Saúde", comparativo_saude, "Meta Saúde"),
            ("Odonto", comparativo_odonto, "Meta Odonto"),
        ]:
            fig = _nova_pagina_pdf(pagina)
            _titulo_pagina_pdf(fig, f"Realizado x Meta - {nome_segmento}")

            realizado = int(comparativo["Realizado"].sum()) if comparativo is not None and not comparativo.empty else 0
            meta = int(comparativo[coluna_meta].sum()) if comparativo is not None and not comparativo.empty else 0
            diferenca = realizado - meta
            atingimento = (realizado / meta * 100) if meta else 0

            _tabela_kpis_pdf(
                fig,
                ["Realizado", "Meta", "Atingimento", "Diferença"],
                [
                    _formatar_inteiro_br(realizado),
                    _formatar_inteiro_br(meta),
                    _formatar_percentual_br(atingimento),
                    f"{diferenca:+,}".replace(",", "."),
                ],
                bbox=(0.085, 0.805, 0.83, 0.075),
            )

            ax = fig.add_axes([0.14, 0.535, 0.73, 0.225])
            _grafico_barras_meta_pdf(
                ax,
                comparativo,
                coluna_meta,
                f"{nome_segmento} - Realizado x Meta",
            )

            tabela_meta = _preparar_tabela_meta(comparativo, coluna_meta)
            _tabela_dataframe_pdf(
                fig,
                tabela_meta,
                bbox=(0.085, 0.255, 0.83, 0.225),
                font_size=6.4,
                col_widths=[0.20, 0.20, 0.20, 0.20, 0.20],
            )

            pdf.savefig(fig, facecolor="white")
            plt.close(fig)
            pagina += 1

        # -------------------------------------------------
        # PORTE DA CARTEIRA - SAÚDE E ODONTO
        # -------------------------------------------------
        for nome_segmento, agrupado in [
            ("Saúde", saude_agru),
            ("Odonto", odonto_agru),
        ]:
            fig = _nova_pagina_pdf(pagina)
            _titulo_pagina_pdf(fig, f"{nome_segmento} - Porte da carteira")

            ax = fig.add_axes([0.14, 0.57, 0.73, 0.245])
            _grafico_stack_competencia_pdf(
                ax,
                agrupado,
                "Qt. Vidas Empresa",
                f"{nome_segmento} por competência e Qt. Vidas Empresa",
                top_n=8,
            )

            ranking = _preparar_ranking_porte(agrupado)
            _tabela_dataframe_pdf(
                fig,
                ranking,
                bbox=(0.20, 0.29, 0.60, 0.20),
                font_size=6.5,
                col_widths=[0.64, 0.36],
            )

            pdf.savefig(fig, facecolor="white")
            plt.close(fig)
            pagina += 1

        # -------------------------------------------------
        # GRUPOS ECONÔMICOS - SOMENTE TABELAS
        # -------------------------------------------------
        pagina = _adicionar_paginas_grupo_economico_pdf(
            pdf,
            montar_tabela_grupo_economico(saude_grupo_eco),
            "Saúde - Grupos Econômicos",
            pagina,
        )

        pagina = _adicionar_paginas_grupo_economico_pdf(
            pdf,
            montar_tabela_grupo_economico(odonto_grupo_eco),
            "Odonto - Grupos Econômicos",
            pagina,
        )

        # -------------------------------------------------
        # VISITAS
        # -------------------------------------------------
        if visitas_filtradas is not None and not visitas_filtradas.empty:
            fig = _nova_pagina_pdf(pagina)
            _titulo_pagina_pdf(
                fig,
                "Visitas da equipe",
                "Consolidação dos arquivos de visitas carregados no dashboard.",
            )

            total_visitas = len(visitas_filtradas)
            total_resp = visitas_filtradas["Responsável"].nunique()
            total_emp = visitas_filtradas.loc[
                visitas_filtradas["Empresa"].ne("Não informado"), "Empresa"
            ].nunique()
            total_horas = visitas_filtradas["Horas Decimal"].sum()

            _tabela_kpis_pdf(
                fig,
                ["Visitas", "Responsáveis", "Empresas", "Horas"],
                [
                    _formatar_inteiro_br(total_visitas),
                    _formatar_inteiro_br(total_resp),
                    _formatar_inteiro_br(total_emp),
                    _formatar_decimal_br(total_horas, 1),
                ],
                bbox=(0.085, 0.805, 0.83, 0.075),
            )

            ax = fig.add_axes([0.14, 0.535, 0.73, 0.225])
            _grafico_visitas_comp_pdf(ax, visitas_filtradas)

            resumo_resp = (
                visitas_filtradas.groupby("Responsável")
                .size()
                .reset_index(name="Visitas")
                .sort_values("Visitas", ascending=False)
            )
            resumo_resp["Visitas"] = resumo_resp["Visitas"].map(_formatar_inteiro_br)
            _tabela_dataframe_pdf(
                fig,
                resumo_resp,
                bbox=(0.20, 0.30, 0.60, 0.16),
                font_size=6.4,
                col_widths=[0.72, 0.28],
                linhas_maximas=10,
            )

            pdf.savefig(fig, facecolor="white")
            plt.close(fig)
            pagina += 1

            # Página de rankings de visita, no mesmo espírito da página de executivos do Comercial.
            fig = _nova_pagina_pdf(pagina)
            _titulo_pagina_pdf(fig, "Visitas - Distribuição das atividades")

            ax1 = fig.add_axes([0.14, 0.60, 0.73, 0.235])
            ax2 = fig.add_axes([0.14, 0.31, 0.73, 0.235])
            _grafico_barra_horizontal_pdf(
                ax1,
                visitas_filtradas.groupby("Responsável").size(),
                "Visitas por responsável",
                "Quantidade de visitas",
            )
            _grafico_barra_horizontal_pdf(
                ax2,
                visitas_filtradas.groupby("Categoria").size(),
                "Visitas por tipo de atividade",
                "Quantidade de visitas",
            )

            pdf.savefig(fig, facecolor="white")
            plt.close(fig)
            pagina += 1

            fig = _nova_pagina_pdf(pagina)
            _titulo_pagina_pdf(fig, "Visitas - Carteira e evolução diária")

            ax1 = fig.add_axes([0.14, 0.60, 0.73, 0.235])
            _grafico_barra_horizontal_pdf(
                ax1,
                visitas_filtradas.groupby("Carteira").size(),
                "Visitas por carteira",
                "Quantidade de visitas",
            )

            ax2 = fig.add_axes([0.14, 0.31, 0.73, 0.225])
            dia = visitas_filtradas.groupby("Data").size().sort_index()
            ax2.plot(
                dia.index,
                dia.values,
                marker="o",
                linewidth=1.6,
                color=VERDE,
            )
            ax2.set_title(
                "Evolução diária das visitas",
                loc="left",
                fontsize=8.5,
                fontweight="bold",
                pad=4,
            )
            ax2.set_ylabel("Visitas", fontsize=7)
            ax2.tick_params(axis="x", rotation=35, labelsize=5.7)
            _estilo_eixo_pdf(ax2)

            pdf.savefig(fig, facecolor="white")
            plt.close(fig)
            pagina += 1

        # -------------------------------------------------
        # CONFERÊNCIA
        # -------------------------------------------------
        fig = _nova_pagina_pdf(pagina)
        _titulo_pagina_pdf(fig, "Conferência da base")

        _tabela_kpis_pdf(
            fig,
            ["Após regras iniciais", "Base válida", "Diferença > 0 / Ativado"],
            [
                _formatar_inteiro_br(len(bruto)),
                _formatar_inteiro_br(len(valido)),
                _formatar_inteiro_br(len(nao_corretos)),
            ],
            bbox=(0.085, 0.805, 0.83, 0.075),
        )

        fig.text(
            0.085,
            0.745,
            "Os indicadores acima permitem conferir as principais etapas de tratamento antes da formação das visões gerenciais.",
            fontsize=8.0,
            color=GRAFITE,
        )

        pdf.savefig(fig, facecolor="white")
        plt.close(fig)
        pagina += 1

        # -------------------------------------------------
        # REGRAS UTILIZADAS
        # -------------------------------------------------
        fig = _nova_pagina_pdf(pagina)
        _titulo_pagina_pdf(fig, "Regras utilizadas")

        regras = pd.DataFrame(
            [
                ["Empresa", "Somente Empresa = 1000."],
                ["Movimentação", "Somente Tipo Movimentação = MOVIMENTAÇÃO."],
                ["Cancelamento", "SEM JUSTA CAUSA, JUSTA CAUSA, A PEDIDO DO BENEFICIÁRIO (RN561), BENEFÍCIO DEMITIDO/APOSENTADO (RN279), DESLIGAMENTO DA EMPRESA (RN279), INADIMPLENTE ou campo vazio."],
                ["Validação de data", "Data Inclusão Vida - Data Início Beneficiário deve resultar em Diferença menor ou igual a zero."],
                ["Competência", "Definida pela Data Inclusão Vida e considerada a partir de 01/2026."],
                ["Saúde", "Id Acomodação diferente de ODO e SEM."],
                ["Odonto", "Id Acomodação igual a ODO."],
                ["Grupo Econômico", "Registros sem Grupo Econômico informado não são apresentados na tabela."],
                ["Visitas", "Os arquivos carregados são consolidados e o responsável é identificado pela coluna Responsável da própria planilha."],
            ],
            columns=["Regra", "Como o dashboard considera"],
        )

        regras["Como o dashboard considera"] = regras[
            "Como o dashboard considera"
        ].map(lambda valor: _quebrar_texto_pdf(valor, 92))

        _tabela_dataframe_pdf(
            fig,
            regras,
            bbox=(0.085, 0.39, 0.83, 0.47),
            font_size=5.7,
            col_widths=[0.22, 0.78],
        )

        pdf.savefig(fig, facecolor="white")
        plt.close(fig)

    buffer.seek(0)
    return buffer.getvalue()

# Uploads
with st.sidebar:
    st.header("Base de dados")

    arquivo = st.file_uploader(
        "Base Total",
        type=["xlsx", "xls"],
        key="base_total",
    )

    arquivos_visitas = st.file_uploader(
        "Relatórios de visitas da equipe",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="arquivos_visitas",
        help="Pode selecionar vários arquivos. Eles serão consolidados automaticamente.",
    )

    st.divider()
    st.markdown("### Regras")
    st.caption("Empresa = 1000")
    st.caption("Tipo Movimentação = MOVIMENTAÇÃO")
    st.caption("Cancelamento homologado ou vazio")
    st.caption("Diferenca ≤ 0")
    st.caption("Competência a partir de 01/2026")
    st.caption("Saúde: diferente de ODO e SEM")
    st.caption("Odonto: ODO")

logo_bytes = obter_logo_bytes()

# Cabeçalho
c_logo, c_title = st.columns([0.9, 4.1], vertical_alignment="center")

with c_logo:
    if logo_bytes:
        st.image(BytesIO(logo_bytes), width=190)
    else:
        st.markdown(
            f"""
            <div style="font-size:28px;font-weight:700;color:{GRAFITE};line-height:1;">
                MEDIATORIE
            </div>
            <div style="font-size:11px;color:#6B736D;margin-top:5px;">
                Administradora de Benefícios
            </div>
            """,
            unsafe_allow_html=True,
        )

with c_title:
    st.markdown(
        """
        <h1 style="margin-bottom:0;color:#404640">
            Relacionamento Empresarial
        </h1>
        <p style="color:#6B736D;margin-top:4px">
            Dashboard de Movimentações — Mediatorie Administradora de Benefícios
        </p>
        """,
        unsafe_allow_html=True,
    )

st.caption(
    "As visões abaixo seguem as regras homologadas do processo de Relacionamento Empresarial."
)

if arquivo is None:

    st.info("Envie a Base Total no menu lateral.")

    st.stop()

try:

    bruto, valido, nao_corretos = processar(

        arquivo.name,

        arquivo.getvalue(),

    )

except Exception as erro:

    st.error(f"Erro ao processar a base: {erro}")

    st.stop()

visitas_base = pd.DataFrame()
erros_visitas = []

if arquivos_visitas:
    visitas_base, erros_visitas = consolidar_visitas(
        arquivos_visitas
    )

# Filtros

competencias_disponiveis = (

    valido[["Competencia_Data", "Competencia"]]

    .dropna()

    .drop_duplicates()

    .sort_values("Competencia_Data")

    ["Competencia"]

    .tolist()

)

qt_disponiveis = sorted(

    valido["Qt. Vidas Empresa"]

    .dropna()

    .unique()

    .tolist()

)

with st.sidebar:

    st.divider()

    st.markdown("### Filtros")

    competencias = st.multiselect(

        "Competência",

        competencias_disponiveis,

        default=competencias_disponiveis,

    )

    qt_vidas = st.multiselect(

        "Qt. Vidas Empresa",

        qt_disponiveis,

        default=qt_disponiveis,

    )

if not competencias or not qt_vidas:

    st.warning("Selecione ao menos uma competência e um valor de Qt. Vidas Empresa.")

    st.stop()

filtrado = valido[

    valido["Competencia"].isin(competencias)

    & valido["Qt. Vidas Empresa"].isin(qt_vidas)

].copy()

if filtrado.empty:

    st.warning("Nenhum registro corresponde aos filtros selecionados.")

    st.stop()

saude, odonto = separar(filtrado)

saude_agru = agrupar_carteira(saude)

odonto_agru = agrupar_carteira(odonto)

saude_grupo_eco = agrupar_grupo_economico(saude)

odonto_grupo_eco = agrupar_grupo_economico(odonto)

# Metas x realizado

metas_2026 = criar_metas_2026()

comparativo_saude = montar_comparativo_meta(

    saude,

    metas_2026,

    "Meta Saúde",

    competencias,

)

comparativo_odonto = montar_comparativo_meta(

    odonto,

    metas_2026,

    "Meta Odonto",

    competencias,

)

# KPIs

k1, k2, k3, k4 = st.columns(4)

k1.metric("Saúde", f"{len(saude):,}".replace(",", "."))

k2.metric("Odonto", f"{len(odonto):,}".replace(",", "."))

k3.metric(

    "Total",

    f"{len(saude) + len(odonto):,}".replace(",", "."),

)

k4.metric(

    "Diferença > 0",

    f"{len(nao_corretos):,}".replace(",", "."),

)

# Abas

visitas_filtradas = pd.DataFrame()

geral, aba_metas, aba_saude, aba_odonto, aba_visitas, conferencia = st.tabs(

    ["Visão Geral", "Metas", "Saúde", "Odonto", "Visitas", "Conferência"]

)

with geral:

    st.plotly_chart(

        grafico_geral(saude, odonto),

        use_container_width=True,

        config={"displaylogo": False},

    )

    a, b = st.columns(2)

    with a:

        st.plotly_chart(

            ranking_carteira(

                saude_agru,

                "Saúde — maior volume por Qt. Vidas Empresa",

            ),

            use_container_width=True,

            config={"displaylogo": False},

        )

    with b:

        st.plotly_chart(

            ranking_carteira(

                odonto_agru,

                "Odonto — maior volume por Qt. Vidas Empresa",

            ),

            use_container_width=True,

            config={"displaylogo": False},

        )

with aba_metas:

    st.subheader("Realizado x Meta")

    st.caption(

        "O realizado utiliza exatamente as mesmas regras do dashboard. "

        "A meta é a meta mensal corporativa cadastrada para Saúde e Odonto."

    )

    # -----------------------------------------------------

    # Saúde

    # -----------------------------------------------------

    realizado_saude = int(

        comparativo_saude["Realizado"].sum()

    ) if not comparativo_saude.empty else 0

    meta_saude = int(

        comparativo_saude["Meta Saúde"].sum()

    ) if not comparativo_saude.empty else 0

    ating_saude = (

        realizado_saude / meta_saude * 100

        if meta_saude

        else 0

    )

    diferenca_saude = (

        realizado_saude - meta_saude

    )

    st.markdown("#### Saúde")

    s1, s2, s3, s4 = st.columns(4)

    s1.metric(

        "Realizado",

        f"{realizado_saude:,}".replace(",", "."),

    )

    s2.metric(

        "Meta",

        f"{meta_saude:,}".replace(",", "."),

    )

    s3.metric(

        "Atingimento",

        percentual_br(ating_saude),

    )

    s4.metric(

        "Diferença",

        f"{diferenca_saude:+,}".replace(",", "."),

    )

    st.plotly_chart(

        grafico_realizado_meta(

            comparativo_saude,

            "Meta Saúde",

            "Saúde — Realizado x Meta",

        ),

        use_container_width=True,

        config={"displaylogo": False},

    )

    tabela_saude = comparativo_saude.copy()

    tabela_saude["Competencia"] = (

        tabela_saude["Competencia"]

        .dt.strftime("%m/%Y")

    )

    st.dataframe(

        tabela_saude,

        use_container_width=True,

        hide_index=True,

        column_config={

            "Atingimento (%)": st.column_config.NumberColumn(

                format="%.1f%%"

            ),

        },

    )

    st.divider()

    # -----------------------------------------------------

    # Odonto

    # -----------------------------------------------------

    realizado_odonto = int(

        comparativo_odonto["Realizado"].sum()

    ) if not comparativo_odonto.empty else 0

    meta_odonto = int(

        comparativo_odonto["Meta Odonto"].sum()

    ) if not comparativo_odonto.empty else 0

    ating_odonto = (

        realizado_odonto / meta_odonto * 100

        if meta_odonto

        else 0

    )

    diferenca_odonto = (

        realizado_odonto - meta_odonto

    )

    st.markdown("#### Odonto")

    o1, o2, o3, o4 = st.columns(4)

    o1.metric(

        "Realizado",

        f"{realizado_odonto:,}".replace(",", "."),

    )

    o2.metric(

        "Meta",

        f"{meta_odonto:,}".replace(",", "."),

    )

    o3.metric(

        "Atingimento",

        percentual_br(ating_odonto),

    )

    o4.metric(

        "Diferença",

        f"{diferenca_odonto:+,}".replace(",", "."),

    )

    st.plotly_chart(

        grafico_realizado_meta(

            comparativo_odonto,

            "Meta Odonto",

            "Odonto — Realizado x Meta",

        ),

        use_container_width=True,

        config={"displaylogo": False},

    )

    tabela_odonto = comparativo_odonto.copy()

    tabela_odonto["Competencia"] = (

        tabela_odonto["Competencia"]

        .dt.strftime("%m/%Y")

    )

    st.dataframe(

        tabela_odonto,

        use_container_width=True,

        hide_index=True,

        column_config={

            "Atingimento (%)": st.column_config.NumberColumn(

                format="%.1f%%"

            ),

        },

    )

    st.info(

        "Ao filtrar Qt. Vidas Empresa, o Realizado considera apenas o recorte "

        "selecionado. A Meta continua sendo a meta corporativa mensal."

    )

with aba_saude:

    st.subheader("Saúde")

    st.caption("Acomodação diferente de ODO e SEM.")

    st.plotly_chart(

        grafico_carteira(

            saude_agru,

            "Saúde por competência e Qt. Vidas Empresa",

        ),

        use_container_width=True,

        config={"displaylogo": False},

    )

    with st.expander("Ver agrupamento Saúde"):

        tabela = saude_agru.copy()

        tabela["Competencia"] = tabela["Competencia"].dt.strftime("%m/%Y")

        st.dataframe(

            tabela,

            use_container_width=True,

            hide_index=True,

        )

    st.divider()
    st.markdown("#### Grupos Econômicos")
    st.caption(
        "A tabela apresenta apenas registros com Grupo Econômico informado. "
        "As competências são exibidas em colunas e o Total ordena os maiores grupos."
    )

    tabela_grupo_saude = montar_tabela_grupo_economico(saude_grupo_eco)

    st.dataframe(
        tabela_grupo_saude,
        use_container_width=True,
        hide_index=True,
    )

with aba_odonto:

    st.subheader("Odonto")

    st.caption("Acomodação igual a ODO.")

    st.plotly_chart(

        grafico_carteira(

            odonto_agru,

            "Odonto por competência e Qt. Vidas Empresa",

        ),

        use_container_width=True,

        config={"displaylogo": False},

    )

    with st.expander("Ver agrupamento Odonto"):

        tabela = odonto_agru.copy()

        tabela["Competencia"] = tabela["Competencia"].dt.strftime("%m/%Y")

        st.dataframe(

            tabela,

            use_container_width=True,

            hide_index=True,

        )

    st.divider()
    st.markdown("#### Grupos Econômicos")
    st.caption(
        "A tabela apresenta apenas registros com Grupo Econômico informado. "
        "As competências são exibidas em colunas e o Total ordena os maiores grupos."
    )

    tabela_grupo_odonto = montar_tabela_grupo_economico(odonto_grupo_eco)

    st.dataframe(
        tabela_grupo_odonto,
        use_container_width=True,
        hide_index=True,
    )

with aba_visitas:
    st.subheader("Visitas da equipe")

    st.caption(
        "Os relatórios enviados são consolidados automaticamente. "
        "A identificação de cada pessoa é feita pela coluna Responsável."
    )

    if not arquivos_visitas:
        st.info(
            "Envie um ou mais relatórios de visitas no menu lateral "
            "para visualizar esta aba."
        )
    elif visitas_base.empty:
        st.warning(
            "Nenhum registro válido de visita foi encontrado nos arquivos enviados."
        )
    else:
        if erros_visitas:
            for mensagem in erros_visitas:
                st.warning(mensagem)

        competencias_visita = (
            visitas_base[
                ["Competencia_Data", "Competencia"]
            ]
            .drop_duplicates()
            .sort_values("Competencia_Data")
            ["Competencia"]
            .tolist()
        )

        responsaveis_visita = sorted(
            visitas_base["Responsável"]
            .dropna()
            .unique()
            .tolist()
        )

        categorias_visita = sorted(
            visitas_base["Categoria"]
            .dropna()
            .unique()
            .tolist()
        )

        carteiras_visita = sorted(
            visitas_base["Carteira"]
            .dropna()
            .unique()
            .tolist()
        )

        st.markdown("#### Filtros das visitas")

        fv1, fv2, fv3, fv4 = st.columns(4)

        with fv1:
            filtro_comp_visita = st.multiselect(
                "Competência das visitas",
                competencias_visita,
                default=competencias_visita,
                key="filtro_comp_visita",
            )

        with fv2:
            filtro_resp_visita = st.multiselect(
                "Responsável",
                responsaveis_visita,
                default=responsaveis_visita,
                key="filtro_resp_visita",
            )

        with fv3:
            filtro_categoria_visita = st.multiselect(
                "Tipo de atividade",
                categorias_visita,
                default=categorias_visita,
                key="filtro_categoria_visita",
            )

        with fv4:
            filtro_carteira_visita = st.multiselect(
                "Carteira",
                carteiras_visita,
                default=carteiras_visita,
                key="filtro_carteira_visita",
            )

        visitas_filtradas = (
            visitas_base.loc[
                visitas_base["Competencia"].isin(filtro_comp_visita)
                & visitas_base["Responsável"].isin(filtro_resp_visita)
                & visitas_base["Categoria"].isin(filtro_categoria_visita)
                & visitas_base["Carteira"].isin(filtro_carteira_visita)
            ]
            .copy()
            .reset_index(drop=True)
        )

        if visitas_filtradas.empty:
            st.warning(
                "Nenhuma visita corresponde aos filtros selecionados."
            )
        else:
            total_visitas = len(visitas_filtradas)

            total_responsaveis = (
                visitas_filtradas["Responsável"]
                .nunique()
            )

            total_empresas = (
                visitas_filtradas.loc[
                    visitas_filtradas["Empresa"]
                    .ne("Não informado"),
                    "Empresa",
                ]
                .nunique()
            )

            total_horas = (
                visitas_filtradas["Horas Decimal"]
                .sum()
            )

            v1, v2, v3, v4 = st.columns(4)

            v1.metric(
                "Total de visitas",
                f"{total_visitas:,}".replace(",", "."),
            )

            v2.metric(
                "Responsáveis",
                f"{total_responsaveis:,}".replace(",", "."),
            )

            v3.metric(
                "Empresas atendidas",
                f"{total_empresas:,}".replace(",", "."),
            )

            v4.metric(
                "Horas registradas",
                f"{total_horas:,.1f}"
                .replace(",", "X")
                .replace(".", ",")
                .replace("X", "."),
            )

            st.plotly_chart(
                grafico_visitas_competencia(
                    visitas_filtradas
                ),
                use_container_width=True,
                config={"displaylogo": False},
            )

            g1, g2 = st.columns(2)

            with g1:
                st.plotly_chart(
                    grafico_visitas_responsavel(
                        visitas_filtradas
                    ),
                    use_container_width=True,
                    config={"displaylogo": False},
                )

            with g2:
                st.plotly_chart(
                    grafico_visitas_categoria(
                        visitas_filtradas
                    ),
                    use_container_width=True,
                    config={"displaylogo": False},
                )

            g3, g4 = st.columns(2)

            with g3:
                st.plotly_chart(
                    grafico_visitas_carteira(
                        visitas_filtradas
                    ),
                    use_container_width=True,
                    config={"displaylogo": False},
                )

            with g4:
                st.plotly_chart(
                    grafico_visitas_dia(
                        visitas_filtradas
                    ),
                    use_container_width=True,
                    config={"displaylogo": False},
                )

            st.markdown("#### Detalhamento")

            colunas_detalhe = [
                coluna
                for coluna in [
                    "Data",
                    "Responsável",
                    "Empresa",
                    "Categoria",
                    "Tipo da tarefa",
                    "Carteira",
                    "Horas",
                    "CNPJ",
                    "Orientações",
                    "Arquivo Origem",
                ]
                if coluna in visitas_filtradas.columns
            ]

            tabela_visitas = (
                visitas_filtradas[colunas_detalhe]
                .sort_values(
                    ["Data", "Responsável"],
                    ascending=[False, True],
                )
            )

            st.dataframe(
                tabela_visitas,
                use_container_width=True,
                hide_index=True,
            )

with conferencia:

    st.subheader("Conferência")

    c1, c2, c3 = st.columns(3)

    c1.metric(

        "Após regras iniciais",

        f"{len(bruto):,}".replace(",", "."),

    )

    c2.metric(

        "Base válida",

        f"{len(valido):,}".replace(",", "."),

    )

    c3.metric(

        "Diferença > 0 / Ativado",

        f"{len(nao_corretos):,}".replace(",", "."),

    )

    with st.expander("Registros com Diferença > 0"):

        st.dataframe(

            nao_corretos,

            use_container_width=True,

            hide_index=True,

        )

    with st.expander("Base válida filtrada"):

        cols = [

            c for c in [

                "Competencia",

                "Descrição Beneficiário",

                "Qt. Vidas Empresa",

                "Id Acomodação",

                "Data Início Beneficiário",

                "Data Inclusão Vida",

                "Diferenca",

            ]

            if c in filtrado.columns

        ]

        st.dataframe(

            filtrado[cols],

            use_container_width=True,

            hide_index=True,

        )

# Exportação do relatório
with st.sidebar:
    st.divider()
    st.markdown("### Relatório")
    st.caption(
        "O PDF respeita os filtros atuais de movimentações e, quando houver, os filtros da aba Visitas."
    )

    if st.button("Gerar PDF", use_container_width=True, type="primary"):
        try:
            with st.spinner("Gerando relatório em PDF..."):
                st.session_state["pdf_relacionamento"] = gerar_pdf_dashboard(
                    saude=saude,
                    odonto=odonto,
                    saude_agru=saude_agru,
                    odonto_agru=odonto_agru,
                    saude_grupo_eco=saude_grupo_eco,
                    odonto_grupo_eco=odonto_grupo_eco,
                    comparativo_saude=comparativo_saude,
                    comparativo_odonto=comparativo_odonto,
                    bruto=bruto,
                    valido=valido,
                    nao_corretos=nao_corretos,
                    competencias=competencias,
                    qt_vidas=qt_vidas,
                    visitas_filtradas=visitas_filtradas,
                    logo_bytes=logo_bytes,
                    nome_base=arquivo.name,
                    nomes_visitas=[a.name for a in arquivos_visitas] if arquivos_visitas else [],
                )
        except Exception as erro:
            st.error(f"Não foi possível gerar o PDF: {erro}")

    if st.session_state.get("pdf_relacionamento"):
        st.download_button(
            "Baixar PDF",
            data=st.session_state["pdf_relacionamento"],
            file_name="relatorio_gerencial_relacionamento_mediatorie.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

