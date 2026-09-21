import io
from pathlib import Path

import numpy as np
import re
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# ============================================================
# CONFIGURAÇÃO DO APP
# ============================================================

st.set_page_config(
    page_title="Relatório Gerencial | Variação das Avaliações",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

DEFAULT_PATH = (
    Path(__file__).resolve().parent.parent
    / "DADOS"
    / "dados_avaliacao.xlsx"
)

FALLBACK_PATHS = [
    DEFAULT_PATH,
    Path.cwd() / "DADOS" / "dados_avaliacao.xlsx",
    Path.cwd() / "dados_avaliacao.xlsx",
]

NOTAS = ["Muito Bom", "Bom", "Ruim", "Muito Ruim"]

COLS_AVALIACAO = [
    # A1
    "a1_cons", "a1_limp", "a1_mob",

    # A2
    "a2_cons", "a2_mob", "a2_limp",

    # A3 banheiro feminino
    "a3bf_cons", "a3bf_repo", "a3bf_limp",

    # A3 banheiro masculino
    "a3bm_limp", "a3bm_cons", "a3bm_repo",

    # A4
    "a4_cons",

    # A5
    "a5_cons", "a5_limp",

    # A6 cozinha
    "a6coz_equip", "a6coz_limp", "a6coz_armz",

    # A6 refeitório
    "a6ref_armz", "a6ref_cons", "a6ref_mob", "a6ref_limp",

    # A6
    "a6_quali",

    # A7
    "a7_jardinag", "a7_cons", "a7_limp_ext",
    "a7_acesso", "a7_dedet", "a7_pint",
]

NOME_CRITERIO = {
    "a1_cons": "A1 - Conservação",
    "a1_limp": "A1 - Limpeza",
    "a1_mob": "A1 - Mobiliário",

    "a2_cons": "A2 - Conservação",
    "a2_mob": "A2 - Mobiliário",
    "a2_limp": "A2 - Limpeza",

    "a3bf_cons": "A3BF - Conservação",
    "a3bf_repo": "A3BF - Reposição",
    "a3bf_limp": "A3BF - Limpeza",

    "a3bm_cons": "A3BM - Conservação",
    "a3bm_repo": "A3BM - Reposição",
    "a3bm_limp": "A3BM - Limpeza",

    "a4_cons": "A4 - Conservação",

    "a5_cons": "A5 - Conservação",
    "a5_limp": "A5 - Limpeza",

    "a6coz_equip": "A6COZ - Equipamentos",
    "a6coz_limp": "A6COZ - Limpeza",
    "a6coz_armz": "A6COZ - Armazenamento",

    "a6ref_armz": "A6REF - Armazenamento",
    "a6ref_cons": "A6REF - Conservação",
    "a6ref_mob": "A6REF - Mobiliário",
    "a6ref_limp": "A6REF - Limpeza",

    "a6_quali": "A6 - Qualidade",

    "a7_jardinag": "A7 - Jardinagem",
    "a7_cons": "A7 - Conservação",
    "a7_limp_ext": "A7 - Limpeza externa",
    "a7_acesso": "A7 - Acesso",
    "a7_dedet": "A7 - Dedetização",
    "a7_pint": "A7 - Pintura",
}

MAPA_NOTA = {
    "muito bom": "Muito Bom",
    "bom": "Bom",
    "ruim": "Ruim",
    "muito ruim": "Muito Ruim",
}


# ============================================================
# ESTILO
# ============================================================

st.markdown(
    """
    <style>
        .block-container {
            padding-top: 1.4rem;
            padding-bottom: 3rem;
        }

        .dashboard-title {
            font-size: 2rem;
            font-weight: 750;
            margin-bottom: 0.15rem;
        }

        .dashboard-subtitle {
            color: #666;
            margin-bottom: 1.25rem;
        }

        .insight-box {
            border: 1px solid rgba(128,128,128,.20);
            border-radius: 12px;
            padding: 1rem 1.1rem;
            margin-bottom: .75rem;
            background: rgba(128,128,128,.04);
        }

        .small-note {
            font-size: .88rem;
            color: #777;
        }

        div[data-testid="stMetric"] {
            border: 1px solid rgba(128,128,128,.18);
            padding: 14px;
            border-radius: 12px;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# LOTE / EMPRESA
# ============================================================

MAPA_EMPRESA_LOTE = {
    "L1": "LBR ENGENHARIA E CONSULTORIA LTDA",
    "L2": "TUV RHEINLAND DUCTOR LTDA",
    "L3": "QUALITEST CIENCIA E TECNOLOGIA LTDA",
    "L4": "QUALITEST CIENCIA E TECNOLOGIA LTDA",
}


# ============================================================
# FUNÇÕES DE APOIO
# ============================================================

def identificar_bloco(criterio: str) -> str:
    if criterio.startswith("a1_"):
        return "A1"
    if criterio.startswith("a2_"):
        return "A2"
    if criterio.startswith("a3bf_"):
        return "A3BF"
    if criterio.startswith("a3bm_"):
        return "A3BM"
    if criterio.startswith("a4_"):
        return "A4"
    if criterio.startswith("a5_"):
        return "A5"
    if criterio.startswith("a6coz_"):
        return "A6COZ"
    if criterio.startswith("a6ref_"):
        return "A6REF"
    if criterio.startswith("a6_"):
        return "A6"
    if criterio.startswith("a7_"):
        return "A7"
    return "OUTRO"


def normalizar_numero_sala(valor):
    """
    Somente padroniza representação numérica:
    01 e 1 passam a representar a mesma sala.
    Não altera a nota.
    """
    if pd.isna(valor):
        return pd.NA

    texto = str(valor).strip()

    if not texto:
        return pd.NA

    # Extrai a primeira sequência de dígitos encontrada (corrige '8⁰' => '8')
    m = re.search(r"\d+", texto)
    if m:
        try:
            return str(int(m.group(0)))
        except Exception:
            return texto

    return texto


def texto_seguro(valor, padrao="Não informado"):
    if pd.isna(valor):
        return padrao

    valor = str(valor).strip()

    return valor if valor else padrao


def formatar_inteiro(valor):
    try:
        return f"{int(valor):,}".replace(",", ".")
    except Exception:
        return "0"


def formatar_pct(valor):
    if pd.isna(valor):
        return "0,00%"
    return f"{valor:.2f}%".replace(".", ",")


@st.cache_data(show_spinner=False, max_entries=3)
def carregar_dados(fonte):
    """
    Carrega e prepara a base apenas uma vez por arquivo.

    Otimizações:
    - lê somente as colunas usadas pelo dashboard;
    - remove colunas auxiliares após criar o ambiente;
    - converte textos repetidos para category;
    - cria um item_id inteiro para acelerar groupby;
    - deixa a base previamente ordenada para as comparações.
    """
    colunas_desejadas = set(
        COLS_AVALIACAO
        + [
            "id_registro",
            "dt_inicio",
            "cod_esc",
            "cod_chaveav",
            "nm_escola",
            "escola",
            "nomecompleto",
            "avaliador",
            "a1_salaAvaliada",
            "a2_ambienteAvaliado",
            "a4_armz",
        ]
    )

    df = pd.read_excel(
        fonte,
        sheet_name="result",
        usecols=lambda coluna: coluna in colunas_desejadas,
    )

    df = df.rename(
        columns={
            "nm_escola": "escola",
            "nomecompleto": "avaliador",
        }
    )

    obrigatorias = [
        "id_registro",
        "dt_inicio",
        "escola",
        "avaliador",
    ]

    faltantes = [
        coluna
        for coluna in obrigatorias
        if coluna not in df.columns
    ]

    if faltantes:
        raise ValueError(
            f"Colunas obrigatórias ausentes: {faltantes}"
        )

    df["dt_inicio"] = pd.to_datetime(
        df["dt_inicio"],
        errors="coerce",
    )

    # --------------------------------------------------------
    # Lote e Empresa
    # --------------------------------------------------------
    if "cod_chaveav" in df.columns:
        df["Lote"] = (
            df["cod_chaveav"]
            .astype("string")
            .str.strip()
            .str[:2]
            .str.upper()
        )

        df["Empresa"] = (
            df["Lote"]
            .map(MAPA_EMPRESA_LOTE)
            .astype("string")
        )
    else:
        df["Lote"] = pd.NA
        df["Empresa"] = pd.NA

    if "cod_esc" not in df.columns:
        df["cod_esc"] = df["escola"].astype("string")

    for coluna in [
        "a1_salaAvaliada",
        "a2_ambienteAvaliado",
        "a4_armz",
    ]:
        if coluna not in df.columns:
            df[coluna] = pd.NA

    criterios = [
        coluna
        for coluna in COLS_AVALIACAO
        if coluna in df.columns
    ]

    id_vars = [
        "id_registro",
        "dt_inicio",
        "cod_esc",
        "escola",
        "avaliador",
        "Lote",
        "Empresa",
        "a1_salaAvaliada",
        "a2_ambienteAvaliado",
        "a4_armz",
    ]

    long_df = df[
        id_vars + criterios
    ].melt(
        id_vars=id_vars,
        value_vars=criterios,
        var_name="criterio",
        value_name="nota_original",
    )

    long_df["nota"] = (
        long_df["nota_original"]
        .astype("string")
        .str.strip()
        .str.lower()
        .map(MAPA_NOTA)
    )

    long_df = long_df.loc[
        long_df["nota"].isin(NOTAS)
    ]

    long_df["bloco"] = (
        long_df["criterio"]
        .map(identificar_bloco)
    )

    long_df["criterio_nome"] = (
        long_df["criterio"]
        .map(NOME_CRITERIO)
        .fillna(long_df["criterio"])
    )

    # --------------------------------------------------------
    # Ambiente real utilizado na comparação
    # --------------------------------------------------------
    long_df["ambiente"] = long_df["bloco"]

    mask_a1 = long_df["bloco"].eq("A1")
    long_df.loc[
        mask_a1,
        "ambiente",
    ] = (
        long_df.loc[
            mask_a1,
            "a1_salaAvaliada",
        ]
        .map(normalizar_numero_sala)
        .map(
            lambda x:
            f"Sala {x}"
            if pd.notna(x)
            else "Sala não informada"
        )
    )

    mask_a2 = long_df["bloco"].eq("A2")
    long_df.loc[
        mask_a2,
        "ambiente",
    ] = (
        long_df.loc[
            mask_a2,
            "a2_ambienteAvaliado",
        ]
        .map(texto_seguro)
    )

    mask_a4 = long_df["bloco"].eq("A4")
    long_df.loc[
        mask_a4,
        "ambiente",
    ] = (
        long_df.loc[
            mask_a4,
            "a4_armz",
        ]
        .map(texto_seguro)
    )

    long_df["mes"] = (
        long_df["dt_inicio"]
        .dt.to_period("M")
        .astype("string")
    )

    # --------------------------------------------------------
    # Mantém somente o que realmente é usado após a preparação.
    # Isso reduz fortemente o consumo de memória.
    # --------------------------------------------------------
    colunas_finais = [
        "id_registro",
        "dt_inicio",
        "cod_esc",
        "escola",
        "avaliador",
        "Lote",
        "Empresa",
        "criterio",
        "nota",
        "bloco",
        "criterio_nome",
        "ambiente",
        "mes",
    ]

    long_df = long_df[colunas_finais].copy()

    # Colunas com muitos valores repetidos ficam muito menores como category.
    colunas_categoria = [
        "escola",
        "avaliador",
        "Lote",
        "Empresa",
        "criterio",
        "nota",
        "bloco",
        "criterio_nome",
        "ambiente",
        "mes",
    ]

    for coluna in colunas_categoria:
        long_df[coluna] = long_df[coluna].astype("category")

    # --------------------------------------------------------
    # ID numérico do item comparado
    # mesma escola + avaliador + critério + ambiente.
    # O groupby posterior passa a trabalhar com um inteiro,
    # em vez de repetir agrupamentos por várias colunas de texto.
    # --------------------------------------------------------
    colunas_item = [
        "cod_esc",
        "escola",
        "avaliador",
        "criterio",
        "criterio_nome",
        "bloco",
        "ambiente",
    ]

    long_df["item_id"] = (
        long_df
        .groupby(
            colunas_item,
            dropna=False,
            observed=True,
            sort=False,
        )
        .ngroup()
        .astype("int32")
    )

    # Ordena uma única vez. Os filtros apenas retiram linhas,
    # portanto a ordem cronológica de cada item é preservada.
    long_df = (
        long_df
        .sort_values(
            ["item_id", "dt_inicio", "id_registro"],
            kind="mergesort",
        )
        .reset_index(drop=True)
    )

    return long_df

def chaves_comparacao(df_long):
    """
    Colunas descritivas do item. O cálculo pesado usa item_id.
    """
    chaves = [
        "cod_esc",
        "escola",
        "avaliador",
        "criterio",
        "criterio_nome",
        "bloco",
        "ambiente",
    ]

    return [
        coluna
        for coluna in chaves
        if coluna in df_long.columns
    ]


def calcular_grupos_comparaveis(df_long):
    """
    Agrupa por item_id inteiro, evitando repetir groupby por várias
    colunas de texto a cada alteração de filtro.
    """
    if df_long.empty:
        return pd.DataFrame()

    colunas_meta = chaves_comparacao(df_long)

    agregacoes = {
        "qtd_avaliacoes": ("nota", "size"),
        "qtd_notas_diferentes": ("nota", "nunique"),
        "primeira_avaliacao": ("dt_inicio", "min"),
        "ultima_avaliacao": ("dt_inicio", "max"),
    }

    for coluna in colunas_meta:
        agregacoes[coluna] = (coluna, "first")

    grupos = (
        df_long
        .groupby(
            "item_id",
            observed=True,
            sort=False,
        )
        .agg(**agregacoes)
    )

    contagem = (
        df_long
        .groupby(
            ["item_id", "nota"],
            observed=True,
            sort=False,
        )
        .size()
        .unstack(fill_value=0)
        .reindex(
            columns=NOTAS,
            fill_value=0,
        )
        .rename(
            columns={
                "Muito Bom": "qtd_muito_bom",
                "Bom": "qtd_bom",
                "Ruim": "qtd_ruim",
                "Muito Ruim": "qtd_muito_ruim",
            }
        )
    )

    grupos = (
        grupos
        .join(contagem, how="left")
        .reset_index()
    )

    grupos = grupos.loc[
        grupos["qtd_avaliacoes"] >= 2
    ].copy()

    grupos["variou"] = (
        grupos["qtd_notas_diferentes"] > 1
    )

    grupos["houve_variacao"] = np.where(
        grupos["variou"],
        "Sim",
        "Não",
    )

    return grupos


def calcular_transicoes(df_long):
    """
    A base já chega ordenada por item_id/data/id_registro.
    Assim não é necessário ordenar centenas de milhares de linhas
    novamente sempre que um filtro muda.
    """
    if df_long.empty:
        return pd.DataFrame()

    agrupado = df_long.groupby(
        "item_id",
        observed=True,
        sort=False,
    )

    nota_anterior = agrupado["nota"].shift(1)
    data_anterior = agrupado["dt_inicio"].shift(1)

    mask = nota_anterior.notna()

    if not mask.any():
        return pd.DataFrame(
            columns=[
                "item_id",
                "avaliador",
                "nota",
                "nota_anterior",
                "dt_inicio",
                "data_anterior",
                "mudou_nota",
            ]
        )

    transicoes = df_long.loc[
        mask,
        [
            "item_id",
            "avaliador",
            "nota",
            "dt_inicio",
        ],
    ].copy()

    transicoes["nota_anterior"] = (
        nota_anterior.loc[mask]
        .astype("string")
        .to_numpy()
    )

    transicoes["data_anterior"] = (
        data_anterior.loc[mask]
        .to_numpy()
    )

    nota_atual_texto = (
        transicoes["nota"]
        .astype("string")
    )

    transicoes["mudou_nota"] = np.where(
        nota_atual_texto.eq(
            transicoes["nota_anterior"]
        ),
        "Não",
        "Sim",
    )

    return transicoes

def criar_resumo_avaliadores(df_long, df_transicoes):
    """
    Resumo gerencial por avaliador.

    Mostra:
    - quantidade total de notas;
    - quantidade e percentual de Muito Bom, Bom, Ruim e Muito Ruim;
    - quantidade de comparações consecutivas;
    - percentual de comparações em que a nota mudou;
    - percentual de comparações em que a nota permaneceu igual.

    O cálculo é feito somente sobre o recorte já filtrado.
    """
    if df_long.empty:
        return pd.DataFrame()

    # --------------------------------------------------------
    # Distribuição das quatro categorias por avaliador
    # --------------------------------------------------------
    contagem_notas = (
        df_long
        .groupby(
            ["avaliador", "nota"],
            observed=True,
            sort=False,
        )
        .size()
        .unstack(fill_value=0)
        .reindex(columns=NOTAS, fill_value=0)
    )

    contagem_notas["Total de notas"] = (
        contagem_notas[NOTAS].sum(axis=1)
    )

    # Quantidade de registros distintos do avaliador.
    registros = (
        df_long
        .groupby(
            "avaliador",
            observed=True,
            sort=False,
        )["id_registro"]
        .nunique()
        .rename("Registros")
    )

    # Quantidade de escolas distintas.
    escolas = (
        df_long
        .groupby(
            "avaliador",
            observed=True,
            sort=False,
        )["escola"]
        .nunique()
        .rename("Escolas")
    )

    resumo = (
        contagem_notas
        .join(registros, how="left")
        .join(escolas, how="left")
        .reset_index()
    )

    # Percentual de cada categoria.
    for nota in NOTAS:
        resumo[f"% {nota}"] = np.where(
            resumo["Total de notas"] > 0,
            resumo[nota] / resumo["Total de notas"] * 100,
            0.0,
        )

    # --------------------------------------------------------
    # Variação entre avaliações consecutivas
    # --------------------------------------------------------
    if not df_transicoes.empty:
        tmp = df_transicoes[
            ["avaliador", "mudou_nota"]
        ].copy()

        tmp["mudou"] = (
            tmp["mudou_nota"].eq("Sim").astype("int8")
        )

        resumo_transicoes_av = (
            tmp
            .groupby(
                "avaliador",
                observed=True,
                sort=False,
            )
            .agg(
                Comparacoes=("mudou", "size"),
                Mudancas=("mudou", "sum"),
            )
            .reset_index()
        )

        resumo_transicoes_av["% Mudança"] = np.where(
            resumo_transicoes_av["Comparacoes"] > 0,
            resumo_transicoes_av["Mudancas"]
            / resumo_transicoes_av["Comparacoes"]
            * 100,
            0.0,
        )

        resumo_transicoes_av["% Estabilidade"] = (
            100 - resumo_transicoes_av["% Mudança"]
        )

        resumo = resumo.merge(
            resumo_transicoes_av,
            on="avaliador",
            how="left",
        )

    else:
        resumo["Comparacoes"] = 0
        resumo["Mudancas"] = 0
        resumo["% Mudança"] = 0.0
        resumo["% Estabilidade"] = 0.0

    colunas_zero = [
        "Comparacoes",
        "Mudancas",
        "% Mudança",
        "% Estabilidade",
    ]

    for coluna in colunas_zero:
        if coluna in resumo.columns:
            resumo[coluna] = resumo[coluna].fillna(0)

    # Nomes mais amigáveis para exibição.
    resumo = resumo.rename(
        columns={
            "avaliador": "Avaliador",
            "Comparacoes": "Comparações",
            "Mudancas": "Mudanças",
        }
    )

    return (
        resumo
        .sort_values(
            ["Total de notas", "Avaliador"],
            ascending=[False, True],
        )
        .reset_index(drop=True)
    )


def resumo_variacao(grupos, campo, long_df=None):
    if grupos.empty:
        return pd.DataFrame()

    resumo = (
        grupos
        .groupby(
            campo,
            as_index=False,
            dropna=False,
            observed=True,
        )
        .agg(
            grupos_comparaveis=("item_id", "size"),
            grupos_com_variacao=("variou", "sum"),
            avaliacoes_nos_grupos=(
                "qtd_avaliacoes",
                "sum",
            ),
        )
    )

    resumo["pct_variacao"] = (
        resumo["grupos_com_variacao"]
        .div(resumo["grupos_comparaveis"])
        .mul(100)
    )

    if long_df is not None and not long_df.empty:
        notas = (
            long_df
            .groupby(
                [campo, "nota"],
                observed=True,
                dropna=False,
            )
            .size()
            .unstack(fill_value=0)
            .reindex(
                columns=NOTAS,
                fill_value=0,
            )
            .reset_index()
        )

        notas["total_notas"] = (
            notas[NOTAS].sum(axis=1)
        )

        notas["negativas"] = (
            notas["Ruim"]
            + notas["Muito Ruim"]
        )

        notas["pct_ruim_muito_ruim"] = (
            notas["negativas"]
            .div(notas["total_notas"])
            .mul(100)
        )

        resumo = resumo.merge(
            notas[
                [
                    campo,
                    "total_notas",
                    "Muito Bom",
                    "Bom",
                    "Ruim",
                    "Muito Ruim",
                    "pct_ruim_muito_ruim",
                ]
            ],
            on=campo,
            how="left",
        )

    return resumo



def criar_detalhamento_itens(dados, grupos):
    """
    Monta a tabela auditável somente quando o usuário pede.
    Isso evita um groupby com concatenação de textos em todo rerun.
    """
    if dados.empty or grupos.empty:
        return pd.DataFrame()

    ids_comparaveis = grupos["item_id"]

    base = dados.loc[
        dados["item_id"].isin(ids_comparaveis),
        [
            "item_id",
            "nota",
            "Lote",
            "Empresa",
        ],
    ]

    sequencias = (
        base
        .groupby(
            "item_id",
            observed=True,
            sort=False,
        )["nota"]
        .agg(
            lambda serie:
            " → ".join(serie.astype(str).tolist())
        )
        .rename("Notas encontradas")
    )

    lotes = (
        base
        .groupby(
            "item_id",
            observed=True,
            sort=False,
        )["Lote"]
        .agg(
            lambda serie:
            " / ".join(
                pd.Series(serie)
                .dropna()
                .astype(str)
                .drop_duplicates()
                .tolist()
            )
        )
        .rename("Lote")
    )

    empresas = (
        base
        .groupby(
            "item_id",
            observed=True,
            sort=False,
        )["Empresa"]
        .agg(
            lambda serie:
            " / ".join(
                pd.Series(serie)
                .dropna()
                .astype(str)
                .drop_duplicates()
                .tolist()
            )
        )
        .rename("Empresa")
    )

    colunas = [
        "item_id",
        "escola",
        "avaliador",
        "criterio_nome",
        "ambiente",
        "qtd_avaliacoes",
    ]

    itens = (
        grupos[colunas]
        .copy()
        .merge(
            sequencias,
            left_on="item_id",
            right_index=True,
            how="left",
        )
        .merge(
            lotes,
            left_on="item_id",
            right_index=True,
            how="left",
        )
        .merge(
            empresas,
            left_on="item_id",
            right_index=True,
            how="left",
        )
    )

    itens["Comparações geradas"] = (
        itens["qtd_avaliacoes"] - 1
    )

    itens = (
        itens
        .rename(
            columns={
                "escola": "Escola",
                "avaliador": "Avaliador",
                "criterio_nome": "Critério",
                "ambiente": "Ambiente",
                "qtd_avaliacoes": "Quantidade de notas",
            }
        )
        .sort_values(
            [
                "Escola",
                "Avaliador",
                "Critério",
                "Ambiente",
            ]
        )
        .reset_index(drop=True)
    )

    itens.insert(
        0,
        "Item",
        range(1, len(itens) + 1),
    )

    return itens[
        [
            "Item",
            "Escola",
            "Avaliador",
            "Lote",
            "Empresa",
            "Critério",
            "Ambiente",
            "Quantidade de notas",
            "Notas encontradas",
            "Comparações geradas",
        ]
    ]

def criar_excel_download(
    resumo_geral,
    distribuicao,
    criterios,
    blocos,
    avaliadores,
    escolas,
    grupos,
    transicoes_resumo,
    mensal,
):
    buffer = io.BytesIO()

    with pd.ExcelWriter(
        buffer,
        engine="openpyxl",
    ) as writer:
        resumo_geral.to_excel(
            writer,
            sheet_name="00_resumo",
            index=False,
        )

        distribuicao.to_excel(
            writer,
            sheet_name="01_notas",
            index=False,
        )

        criterios.to_excel(
            writer,
            sheet_name="02_criterios",
            index=False,
        )

        blocos.to_excel(
            writer,
            sheet_name="03_blocos",
            index=False,
        )

        avaliadores.to_excel(
            writer,
            sheet_name="04_avaliadores",
            index=False,
        )

        escolas.to_excel(
            writer,
            sheet_name="05_escolas",
            index=False,
        )

        grupos.to_excel(
            writer,
            sheet_name="06_detalhado",
            index=False,
        )

        transicoes_resumo.to_excel(
            writer,
            sheet_name="07_transicoes",
            index=False,
        )

        mensal.to_excel(
            writer,
            sheet_name="08_mensal",
            index=False,
        )

    buffer.seek(0)

    return buffer.getvalue()


# ============================================================
# CABEÇALHO
# ============================================================

st.markdown(
    '<div class="dashboard-title">'
    "Relatório Gerencial — Acompanhamento das Avaliações"
    "</div>",
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="dashboard-subtitle">'
    "Visão gerencial das notas Muito Bom, Bom, Ruim e Muito Ruim. "
    "Para medir mudanças ao longo do tempo, o painel compara somente "
    "avaliações feitas na mesma escola, pelo mesmo avaliador, para o "
    "mesmo critério e o mesmo ambiente."
    "</div>",
    unsafe_allow_html=True,
)




# ============================================================
# FONTE DE DADOS
# ============================================================

with st.sidebar:
    st.header("Dados")

    arquivo_upload = st.file_uploader(
        "Carregar dados_avaliacao.xlsx",
        type=["xlsx"],
    )

    if arquivo_upload is not None:
        fonte = arquivo_upload
        nome_fonte = arquivo_upload.name

    elif any(path.exists() for path in FALLBACK_PATHS):
        for path in FALLBACK_PATHS:
            if path.exists():
                fonte = path
                nome_fonte = str(path)
                break
    else:
        fonte = None
        nome_fonte = None


if fonte is None:
    st.info(
        "Carregue o arquivo **dados_avaliacao.xlsx** na barra lateral. "
        "Ao executar localmente no seu projeto, o app também tenta abrir "
        "automaticamente o caminho configurado em `DEFAULT_PATH`."
    )
    st.stop()


try:
    with st.spinner("Processando avaliações..."):
        long_df = carregar_dados(fonte)

except Exception as erro:
    st.error(f"Não foi possível carregar a planilha: {erro}")
    st.stop()


# ============================================================
# FILTROS
# ============================================================

data_min = long_df["dt_inicio"].min()
data_max = long_df["dt_inicio"].max()

with st.sidebar:
    st.success(f"Fonte carregada: {nome_fonte}")

    st.header("Filtros")

    if pd.notna(data_min) and pd.notna(data_max):
        periodo = st.date_input(
            "Período",
            value=(
                data_min.date(),
                data_max.date(),
            ),
            min_value=data_min.date(),
            max_value=data_max.date(),
        )
    else:
        periodo = None

    escolas_disponiveis = sorted(
        str(valor)
        for valor in long_df["escola"].dropna().unique()
    )

    escolas_selecionadas = st.multiselect(
        "Escolas",
        options=escolas_disponiveis,
        placeholder="Todas as escolas",
    )

    empresas_disponiveis = sorted(
        str(valor)
        for valor in long_df["Empresa"].dropna().unique()
    )

    empresas_selecionadas = st.multiselect(
        "Empresa",
        options=empresas_disponiveis,
        placeholder="Todas as empresas",
        help=(
            "A empresa é identificada automaticamente pelos dois primeiros "
            "caracteres da coluna cod_chaveav: "
            "L1 = LBR ENGENHARIA E CONSULTORIA LTDA; "
            "L2 = TUV RHEINLAND DUCTOR LTDA; "
            "L3 e L4 = QUALITEST CIENCIA E TECNOLOGIA LTDA. "
            "Ao selecionar uma empresa, o filtro de Avaliadores logo abaixo "
            "passa a mostrar somente os avaliadores vinculados a ela."
        ),
    )

    # --------------------------------------------------------
    # Filtro dependente: Empresa -> Avaliadores
    # --------------------------------------------------------
    # Se uma ou mais empresas forem selecionadas, a lista de avaliadores
    # é construída somente com os registros dessas empresas.
    # Sem seleção de empresa, todos os avaliadores ficam disponíveis.
    if empresas_selecionadas:
        serie_avaliadores = long_df.loc[
            long_df["Empresa"].isin(empresas_selecionadas),
            "avaliador",
        ]
    else:
        serie_avaliadores = long_df["avaliador"]

    avaliadores_disponiveis = sorted(
        str(valor)
        for valor in serie_avaliadores.dropna().unique()
    )

    avaliadores_selecionados = st.multiselect(
        "Avaliadores",
        options=avaliadores_disponiveis,
        placeholder=(
            "Avaliadores da empresa selecionada"
            if empresas_selecionadas
            else "Todos os avaliadores"
        ),
        help=(
            "A lista de avaliadores é atualizada automaticamente de acordo "
            "com a empresa selecionada acima."
        ),
    )

    blocos_disponiveis = sorted(
        long_df["bloco"]
        .dropna()
        .unique()
        .tolist()
    )

    blocos_selecionados = st.multiselect(
        "Blocos / ambientes",
        options=blocos_disponiveis,
        placeholder="Todos os blocos",
    )

    criterios_disponiveis = (
        long_df[
            ["criterio", "criterio_nome"]
        ]
        .drop_duplicates()
        .sort_values("criterio_nome")
    )

    mapa_criterios = dict(
        zip(
            criterios_disponiveis["criterio_nome"],
            criterios_disponiveis["criterio"],
        )
    )

    criterios_nomes = st.multiselect(
        "Critérios",
        options=list(mapa_criterios.keys()),
        placeholder="Todos os critérios",
    )

    st.divider()

    top_n = st.slider(
        "Quantidade de itens nos rankings",
        min_value=5,
        max_value=30,
        value=15,
        step=5,
    )


# ============================================================
# APLICAÇÃO DOS FILTROS
# ============================================================

# Um único vetor booleano é montado e a base é recortada apenas uma vez.
# Isso evita várias cópias sucessivas de uma base com centenas de milhares
# de linhas.
mask_filtro = pd.Series(
    True,
    index=long_df.index,
)

if periodo is not None:
    if isinstance(periodo, (tuple, list)) and len(periodo) == 2:
        inicio = pd.Timestamp(periodo[0])
        fim = pd.Timestamp(periodo[1]) + pd.Timedelta(days=1)

        mask_filtro &= (
            long_df["dt_inicio"].ge(inicio)
            & long_df["dt_inicio"].lt(fim)
        )


if escolas_selecionadas:
    mask_filtro &= long_df["escola"].isin(
        escolas_selecionadas
    )


if empresas_selecionadas:
    mask_filtro &= long_df["Empresa"].isin(
        empresas_selecionadas
    )


if avaliadores_selecionados:
    mask_filtro &= long_df["avaliador"].isin(
        avaliadores_selecionados
    )


if blocos_selecionados:
    mask_filtro &= long_df["bloco"].isin(
        blocos_selecionados
    )


if criterios_nomes:
    criterios_codigos = [
        mapa_criterios[nome]
        for nome in criterios_nomes
    ]

    mask_filtro &= long_df["criterio"].isin(
        criterios_codigos
    )


dados = long_df.loc[mask_filtro]


if dados.empty:
    st.warning(
        "Nenhuma avaliação foi encontrada com os filtros selecionados."
    )
    st.stop()


# ============================================================
# CÁLCULOS
# ============================================================

grupos = calcular_grupos_comparaveis(dados)
transicoes = calcular_transicoes(dados)

distribuicao = (
    dados["nota"]
    .value_counts()
    .reindex(
        NOTAS,
        fill_value=0,
    )
    .rename_axis("nota")
    .reset_index(name="quantidade")
)

distribuicao["percentual"] = (
    distribuicao["quantidade"]
    .div(distribuicao["quantidade"].sum())
    .mul(100)
)

resumo_criterios = resumo_variacao(
    grupos,
    "criterio_nome",
    dados,
)

resumo_blocos = resumo_variacao(
    grupos,
    "bloco",
    dados,
)

# Resumo leve usado na nova aba "Avaliadores".
# O agrupamento final possui poucas linhas (uma por avaliador)
# e aproveita a base já filtrada.
resumo_avaliadores_painel = criar_resumo_avaliadores(
    dados,
    transicoes,
)


# Os resumos adicionais de exportação por escola continuam sendo
# calculados somente quando o usuário solicita o Excel.


# ============================================================
# MÉTRICAS EXECUTIVAS
# ============================================================

total_registros = dados["id_registro"].nunique()
total_escolas = dados["escola"].nunique()
total_avaliadores = dados["avaliador"].nunique()
total_notas = len(dados)

qtd_positivas = dados["nota"].isin(
    ["Bom", "Muito Bom"]
).sum()

qtd_negativas = dados["nota"].isin(
    ["Ruim", "Muito Ruim"]
).sum()

pct_positivas = (
    qtd_positivas / total_notas * 100
    if total_notas
    else 0
)

pct_negativas = (
    qtd_negativas / total_notas * 100
    if total_notas
    else 0
)

qtd_grupos = len(grupos)

qtd_grupos_variacao = (
    grupos["houve_variacao"]
    .eq("Sim")
    .sum()
    if not grupos.empty
    else 0
)

pct_grupos_variacao = (
    qtd_grupos_variacao / qtd_grupos * 100
    if qtd_grupos
    else 0
)

total_transicoes = len(transicoes)

qtd_mudancas = (
    transicoes["mudou_nota"]
    .eq("Sim")
    .sum()
    if not transicoes.empty
    else 0
)

pct_mudancas = (
    qtd_mudancas / total_transicoes * 100
    if total_transicoes
    else 0
)

pct_estabilidade = (
    100 - pct_mudancas
    if total_transicoes
    else 0
)


# Quantidade de critérios efetivamente presentes no recorte atual.
# Usada apenas para explicar de forma simples o cartão de avaliações categóricas.
qtd_criterios_ativos = dados["criterio"].nunique()

avaliacoes_possiveis_exemplo = (
    total_registros * qtd_criterios_ativos
)


# Números usados nas explicações dinâmicas dos cartões.
# 'notas_em_itens_repetidos' representa a soma das notas pertencentes
# aos itens que aparecem pelo menos duas vezes no recorte atual.
notas_em_itens_repetidos = (
    int(grupos["qtd_avaliacoes"].sum())
    if not grupos.empty
    else 0
)


# Explicação dinâmica de como o total de notas foi formado.
if not grupos.empty:
    distribuicao_notas_por_item = (
        grupos["qtd_avaliacoes"]
        .value_counts()
        .sort_index()
    )

    partes_calculo_notas = []
    linhas_calculo_notas = []

    for qtd_notas_item, qtd_itens in distribuicao_notas_por_item.items():
        qtd_notas_item = int(qtd_notas_item)
        qtd_itens = int(qtd_itens)
        subtotal = qtd_notas_item * qtd_itens

        trecho = (
            f"{formatar_inteiro(qtd_itens)} "
            f"{'item' if qtd_itens == 1 else 'itens'} × "
            f"{formatar_inteiro(qtd_notas_item)} "
            f"{'nota' if qtd_notas_item == 1 else 'notas'} = "
            f"{formatar_inteiro(subtotal)} notas"
        )

        partes_calculo_notas.append(trecho)
        linhas_calculo_notas.append(f"- {trecho}")

    calculo_detalhado_notas = "; ".join(partes_calculo_notas)
    calculo_detalhado_notas_multilinha = "\n".join(linhas_calculo_notas)

    # Validação interna: a soma dos subtotais precisa ser igual ao total.
    total_recalculado_notas = int(
        sum(
            int(qtd_notas) * int(qtd_itens)
            for qtd_notas, qtd_itens in distribuicao_notas_por_item.items()
        )
    )

else:
    calculo_detalhado_notas = ""
    calculo_detalhado_notas_multilinha = ""
    total_recalculado_notas = 0


# O detalhamento linha a linha dos itens é montado somente sob demanda
# na aba "Mudanças entre Avaliações".


qtd_sem_mudanca = (
    total_transicoes - qtd_mudancas
)

qtd_muito_bom = int(
    dados["nota"].eq("Muito Bom").sum()
)

qtd_bom = int(
    dados["nota"].eq("Bom").sum()
)

qtd_ruim = int(
    dados["nota"].eq("Ruim").sum()
)

qtd_muito_ruim = int(
    dados["nota"].eq("Muito Ruim").sum()
)

qtd_positivas_atual = (
    qtd_muito_bom + qtd_bom
)

qtd_negativas_atual = (
    qtd_ruim + qtd_muito_ruim
)


# ============================================================
# TABELAS TEMPORAIS / TRANSIÇÕES
# ============================================================

mensal = (
    dados
    .groupby(
        ["mes", "nota"],
        observed=True,
    )
    .size()
    .unstack(fill_value=0)
    .reindex(
        columns=NOTAS,
        fill_value=0,
    )
    .reset_index()
)

mensal["total"] = mensal[NOTAS].sum(axis=1)

for nota in NOTAS:
    nome = (
        "pct_"
        + nota.lower().replace(" ", "_")
    )

    mensal[nome] = (
        mensal[nota]
        .div(mensal["total"])
        .mul(100)
    )

mensal["pct_ruim_muito_ruim"] = (
    mensal["Ruim"]
    .add(mensal["Muito Ruim"])
    .div(mensal["total"])
    .mul(100)
)


if not transicoes.empty:
    resumo_transicoes = (
        transicoes
        .groupby(
            ["nota_anterior", "nota"],
            as_index=False,
            observed=True,
            sort=False,
        )
        .size()
        .rename(columns={"size": "quantidade"})
    )

    resumo_transicoes["transicao"] = (
        resumo_transicoes["nota_anterior"].astype(str)
        + " → "
        + resumo_transicoes["nota"].astype(str)
    )

    resumo_transicoes = (
        resumo_transicoes[
            [
                "transicao",
                "quantidade",
            ]
        ]
        .sort_values(
            "quantidade",
            ascending=False,
        )
        .reset_index(drop=True)
    )

    resumo_transicoes["percentual"] = (
        resumo_transicoes["quantidade"]
        .div(resumo_transicoes["quantidade"].sum())
        .mul(100)
    )
else:
    resumo_transicoes = pd.DataFrame(
        columns=[
            "transicao",
            "quantidade",
            "percentual",
        ]
    )


# ============================================================
# RESUMO PARA DOWNLOAD
# ============================================================

resumo_geral_download = pd.DataFrame(
    {
        "indicador": [
            "Registros",
            "Escolas",
            "Avaliadores",
            "Avaliações categóricas",
            "Bom + Muito Bom (%)",
            "Ruim + Muito Ruim (%)",
            "Grupos comparáveis",
            "Grupos com variação",
            "Grupos com variação (%)",
            "Transições consecutivas",
            "Mudanças consecutivas (%)",
            "Estabilidade consecutiva (%)",
        ],
        "valor": [
            total_registros,
            total_escolas,
            total_avaliadores,
            total_notas,
            pct_positivas,
            pct_negativas,
            qtd_grupos,
            qtd_grupos_variacao,
            pct_grupos_variacao,
            total_transicoes,
            pct_mudancas,
            pct_estabilidade,
        ],
    }
)


# ============================================================
# ABAS
# ============================================================

tab_executivo, tab_criterios, tab_avaliadores, tab_transicoes = st.tabs(
    [
        "Visão Executiva",
        "Critérios e Ambientes",
        "Avaliadores",
        "Mudanças entre Avaliações",
    ]
)


# ============================================================
# VISÃO EXECUTIVA
# ============================================================

with tab_executivo:

    st.subheader("Indicadores principais")

    with st.expander("Entenda os números deste filtro", expanded=False):
        if total_transicoes:
            st.markdown(
                f"""
                **O recorte atual possui {formatar_inteiro(total_registros)} registros,**
                feitos por **{formatar_inteiro(total_avaliadores)} avaliadores**, com
                **{formatar_inteiro(total_notas)} respostas válidas** nos critérios selecionados.

                Para acompanhar mudanças ao longo do tempo, foram encontrados
                **{formatar_inteiro(qtd_grupos)} itens avaliados mais de uma vez**.
                Somando todas as vezes em que esses itens receberam uma classificação, temos
                **{formatar_inteiro(notas_em_itens_repetidos)} notas**.

                De onde vem esse total?

                {calculo_detalhado_notas_multilinha}

                Somando todos os subtotais, chegamos exatamente a
                **{formatar_inteiro(total_recalculado_notas)} notas**.

                Como a primeira nota de cada item é apenas o ponto de partida,
                a conta das comparações é:

                **{formatar_inteiro(notas_em_itens_repetidos)} notas
                - {formatar_inteiro(qtd_grupos)} primeiras notas
                = {formatar_inteiro(total_transicoes)} comparações.**

                Dessas comparações:

                - **{formatar_inteiro(qtd_sem_mudanca)}** mantiveram a mesma nota
                  (**{formatar_pct(pct_estabilidade)}**);
                - **{formatar_inteiro(qtd_mudancas)}** mudaram de nota
                  (**{formatar_pct(pct_mudancas)}**).
                """
            )
        else:
            st.markdown(
                f"""
                O recorte atual possui **{formatar_inteiro(total_registros)} registros**
                e **{formatar_inteiro(total_notas)} respostas válidas**.

                Neste filtro ainda não existem avaliações repetidas suficientes do mesmo item
                para comparar uma avaliação com a seguinte.
                """
            )

    # Top metrics: remove 'Escolas' per request
    c1, c2, c3 = st.columns(3)

    c1.metric(
        "Registros",
        formatar_inteiro(total_registros),
        help=(
            f"Na seleção atual existem {formatar_inteiro(total_registros)} registros diferentes. "
            "Cada registro é contado apenas uma vez. "
            f"Portanto, para os filtros escolhidos agora, o cartão mostra "
            f"{formatar_inteiro(total_registros)} registros."
        ),
    )

    c2.metric(
        "Avaliadores",
        formatar_inteiro(total_avaliadores),
        help=(
            f"Na seleção atual existem {formatar_inteiro(total_avaliadores)} avaliadores diferentes. "
            "Se a mesma pessoa aparece em vários registros, ela é contada apenas uma vez."
        ),
    )

    c3.metric(
        "Avaliações categóricas",
        formatar_inteiro(total_notas),
        help=(
            f"Na seleção atual existem {formatar_inteiro(total_registros)} registros e "
            f"{formatar_inteiro(qtd_criterios_ativos)} critérios diferentes. "
            f"Se todos os critérios estivessem preenchidos em todos os registros, haveria até "
            f"{formatar_inteiro(total_registros)} × {formatar_inteiro(qtd_criterios_ativos)} = "
            f"{formatar_inteiro(avaliacoes_possiveis_exemplo)} respostas. "
            f"Foram encontradas {formatar_inteiro(total_notas)} respostas válidas com as notas "
            "Muito Bom, Bom, Ruim ou Muito Ruim. "
            f"Por isso o cartão mostra {formatar_inteiro(total_notas)}."
        ),
    )

    # Secondary metrics: remove 'Grupos com variação' per request
    c5, c6, c7 = st.columns(3)

    c5.metric(
        "Bom + Muito Bom",
        formatar_pct(pct_positivas),
        help=(
            f"Na seleção atual há {formatar_inteiro(total_notas)} respostas válidas. "
            f"Dessas, {formatar_inteiro(qtd_bom)} são Bom e "
            f"{formatar_inteiro(qtd_muito_bom)} são Muito Bom. "
            f"Somando: {formatar_inteiro(qtd_bom)} + {formatar_inteiro(qtd_muito_bom)} = "
            f"{formatar_inteiro(qtd_positivas_atual)} respostas positivas. "
            f"Isso representa {formatar_pct(pct_positivas)} do total."
        ),
    )

    c6.metric(
        "Ruim + Muito Ruim",
        formatar_pct(pct_negativas),
        help=(
            f"Na seleção atual há {formatar_inteiro(total_notas)} respostas válidas. "
            f"Dessas, {formatar_inteiro(qtd_ruim)} são Ruim e "
            f"{formatar_inteiro(qtd_muito_ruim)} são Muito Ruim. "
            f"Somando: {formatar_inteiro(qtd_ruim)} + {formatar_inteiro(qtd_muito_ruim)} = "
            f"{formatar_inteiro(qtd_negativas_atual)} respostas negativas. "
            f"Isso representa {formatar_pct(pct_negativas)} do total."
        ),
    )

    c7.metric(
        "Estabilidade entre visitas",
        formatar_pct(pct_estabilidade),
        help=(
            f"Na seleção atual foram feitas {formatar_inteiro(total_transicoes)} comparações "
            f"entre uma avaliação e a seguinte do mesmo item. "
            f"Em {formatar_inteiro(qtd_sem_mudanca)} comparações a nota permaneceu igual. "
            f"Assim, {formatar_inteiro(qtd_sem_mudanca)} ÷ "
            f"{formatar_inteiro(total_transicoes)} = {formatar_pct(pct_estabilidade)}. "
            "A comparação considera sempre a mesma escola, o mesmo avaliador, "
            "o mesmo critério e o mesmo ambiente."
            if total_transicoes
            else
            "Não existem avaliações repetidas suficientes no filtro atual para calcular estabilidade."
        ),
    )

    st.divider()

    col_a, col_b = st.columns(
        [1, 1.3]
    )

    with col_a:
        st.subheader("Distribuição das notas")

        fig = px.bar(
            distribuicao,
            x="nota",
            y="percentual",
            text="percentual",
            category_orders={
                "nota": NOTAS
            },
            labels={
                "nota": "",
                "percentual": "Percentual (%)",
            },
        )

        fig.update_traces(
            texttemplate="%{text:.1f}%",
            textposition="outside",
        )

        fig.update_layout(
            showlegend=False,
            margin=dict(
                l=10,
                r=10,
                t=20,
                b=10,
            ),
        )

        st.plotly_chart(
            fig,
            use_container_width=True,
            key="grafico_executivo_distribuicao_notas",
        )

    with col_b:
        st.subheader("Evolução mensal")

        mensal_long = mensal.melt(
            id_vars=["mes"],
            value_vars=[
                "pct_muito_bom",
                "pct_bom",
                "pct_ruim",
                "pct_muito_ruim",
            ],
            var_name="categoria",
            value_name="percentual",
        )

        mensal_long["categoria"] = (
            mensal_long["categoria"]
            .str.replace("pct_", "", regex=False)
            .str.replace("_", " ", regex=False)
            .str.title()
        )

        fig = px.line(
            mensal_long,
            x="mes",
            y="percentual",
            color="categoria",
            markers=True,
            labels={
                "mes": "Mês",
                "percentual": "Percentual (%)",
                "categoria": "Nota",
            },
        )

        fig.update_layout(
            margin=dict(
                l=10,
                r=10,
                t=20,
                b=10,
            ),
        )

        st.plotly_chart(
            fig,
            use_container_width=True,
            key="grafico_executivo_evolucao_mensal",
        )

    st.divider()

    st.subheader("Principais insights do recorte")

    insights = []

    if not distribuicao.empty:
        linha = distribuicao.sort_values(
            "quantidade",
            ascending=False,
        ).iloc[0]

        insights.append(
            f"A classificação mais frequente é **{linha['nota']}**, "
            f"representando **{linha['percentual']:.2f}%** das avaliações."
        )

    if not resumo_criterios.empty:
        linha = (
            resumo_criterios
            .sort_values(
                [
                    "pct_variacao",
                    "grupos_comparaveis",
                ],
                ascending=[False, False],
            )
            .iloc[0]
        )

        insights.append(
            f"O critério que mais mudou de nota foi **{linha['criterio_nome']}**. "
            f"Em **{linha['pct_variacao']:.2f}%** dos itens acompanhados mais de uma vez, "
            f"a classificação mudou em algum momento."
        )

        linha_neg = (
            resumo_criterios
            .dropna(
                subset=["pct_ruim_muito_ruim"]
            )
            .sort_values(
                [
                    "pct_ruim_muito_ruim",
                    "total_notas",
                ],
                ascending=[False, False],
            )
            .iloc[0]
        )

        insights.append(
            f"A maior participação de **Ruim + Muito Ruim** ocorre em "
            f"**{linha_neg['criterio_nome']}**, com "
            f"**{linha_neg['pct_ruim_muito_ruim']:.2f}%** em "
            f"**{int(linha_neg['total_notas']):,}** classificações."
        )

    if total_transicoes:
        insights.append(
            f"De uma avaliação para a seguinte, **{pct_estabilidade:.2f}%** das notas "
            f"permaneceram iguais e **{pct_mudancas:.2f}%** mudaram."
        )

    if not resumo_transicoes.empty:
        linha = resumo_transicoes.iloc[0]

        insights.append(
            f"A transição mais frequente é **{linha['transicao']}**, "
            f"com **{int(linha['quantidade']):,}** ocorrências "
            f"(**{linha['percentual']:.2f}%** das transições)."
        )

    if not resumo_blocos.empty:
        linha = (
            resumo_blocos
            .sort_values(
                [
                    "pct_variacao",
                    "grupos_comparaveis",
                ],
                ascending=[False, False],
            )
            .iloc[0]
        )

        insights.append(
            f"Entre os blocos analisados, **{linha['bloco']}** foi o que apresentou "
            f"mais mudanças de classificação: **{linha['pct_variacao']:.2f}%**."
        )

    for texto in insights:
        st.markdown(
            f'<div class="insight-box">{texto}</div>',
            unsafe_allow_html=True,
        )

    st.caption(
        "As notas continuam sendo Muito Bom, Bom, Ruim e Muito Ruim. "
        "O painel não cria nota numérica: ele apenas verifica se a classificação "
        "permaneceu igual ou mudou ao longo das avaliações."
    )


# ============================================================
# CRITÉRIOS E AMBIENTES
# ============================================================

with tab_criterios:

    st.subheader(
        "Como as notas mudam por critério"
    )

    if resumo_criterios.empty:
        st.info(
            "Não há avaliações repetidas suficientes para comparar "
            "no período e filtros selecionados."
        )

    else:
        col1, col2 = st.columns(2)

        with col1:
            dados_grafico = (
                resumo_criterios
                .nlargest(
                    top_n,
                    "pct_variacao",
                )
                .sort_values(
                    "pct_variacao"
                )
            )

            fig = px.bar(
                dados_grafico,
                x="pct_variacao",
                y="criterio_nome",
                orientation="h",
                text="pct_variacao",
                labels={
                    "pct_variacao":
                        "% de itens que mudaram",
                    "criterio_nome":
                        "",
                },
                hover_data=[
                    "grupos_comparaveis",
                    "grupos_com_variacao",
                ],
            )

            fig.update_traces(
                texttemplate="%{text:.1f}%",
                textposition="outside",
            )

            fig.update_layout(
                title=(
                    f"{top_n} critérios que mais mudaram de nota"
                ),
                margin=dict(
                    l=10,
                    r=10,
                    t=50,
                    b=10,
                ),
            )

            st.plotly_chart(
                fig,
                use_container_width=True,
                key="grafico_criterios_variacao",
            )

        with col2:
            dados_grafico = (
                resumo_criterios
                .dropna(
                    subset=[
                        "pct_ruim_muito_ruim"
                    ]
                )
                .nlargest(
                    top_n,
                    "pct_ruim_muito_ruim",
                )
                .sort_values(
                    "pct_ruim_muito_ruim"
                )
            )

            fig = px.bar(
                dados_grafico,
                x="pct_ruim_muito_ruim",
                y="criterio_nome",
                orientation="h",
                text="pct_ruim_muito_ruim",
                labels={
                    "pct_ruim_muito_ruim":
                        "% Ruim + Muito Ruim",
                    "criterio_nome":
                        "",
                },
                hover_data=[
                    "total_notas",
                    "Ruim",
                    "Muito Ruim",
                ],
            )

            fig.update_traces(
                texttemplate="%{text:.1f}%",
                textposition="outside",
            )

            fig.update_layout(
                title=(
                    f"{top_n} critérios com maior "
                    "participação negativa"
                ),
                margin=dict(
                    l=10,
                    r=10,
                    t=50,
                    b=10,
                ),
            )

            st.plotly_chart(
                fig,
                use_container_width=True,
                key="grafico_criterios_notas_negativas",
            )

        


# ============================================================
# AVALIADORES
# ============================================================

with tab_avaliadores:

    st.subheader("Acompanhamento das notas por avaliador")

    st.markdown(
        """
        Esta visão mostra **como cada avaliador distribui suas classificações**
        entre **Muito Bom, Bom, Ruim e Muito Ruim**.

        Também mostra com que frequência a classificação **muda de uma avaliação
        para a seguinte**, sempre comparando o mesmo item: mesma escola,
        mesmo critério e mesmo ambiente.

        Uma variação maior ou menor não significa, por si só, que um avaliador
        está certo ou errado. O indicador apenas mostra o comportamento das
        classificações ao longo das avaliações.
        """
    )

    if resumo_avaliadores_painel.empty:
        st.info(
            "Não existem avaliações para os filtros selecionados."
        )

    else:
        st.subheader("Resumo dos avaliadores")

        st.caption(
            "A tabela abaixo respeita todos os filtros atuais da aplicação, "
            "inclusive Empresa. Os percentuais mostram como cada avaliador "
            "distribuiu suas notas e com que frequência a classificação mudou "
            "entre avaliações consecutivas."
        )

        colunas_resumo_av = [
            "Avaliador",
            "Registros",
            "Escolas",
            "Total de notas",
            "Muito Bom",
            "% Muito Bom",
            "Bom",
            "% Bom",
            "Ruim",
            "% Ruim",
            "Muito Ruim",
            "% Muito Ruim",
            "Comparações",
            "Mudanças",
            "% Mudança",
            "% Estabilidade",
        ]

        resumo_exibicao_av = (
            resumo_avaliadores_painel[
                [
                    coluna
                    for coluna in colunas_resumo_av
                    if coluna in resumo_avaliadores_painel.columns
                ]
            ]
            .copy()
        )

        st.dataframe(
            resumo_exibicao_av,
            use_container_width=True,
            hide_index=True,
            column_config={
                "% Muito Bom": st.column_config.NumberColumn(
                    "% Muito Bom",
                    format="%.2f%%",
                ),
                "% Bom": st.column_config.NumberColumn(
                    "% Bom",
                    format="%.2f%%",
                ),
                "% Ruim": st.column_config.NumberColumn(
                    "% Ruim",
                    format="%.2f%%",
                ),
                "% Muito Ruim": st.column_config.NumberColumn(
                    "% Muito Ruim",
                    format="%.2f%%",
                ),
                "% Mudança": st.column_config.NumberColumn(
                    "% Mudança",
                    format="%.2f%%",
                ),
                "% Estabilidade": st.column_config.NumberColumn(
                    "% Estabilidade",
                    format="%.2f%%",
                ),
            },
        )

        st.caption(
            "Importante: a taxa de mudança é descritiva. Ela mostra quantas vezes "
            "a nota mudou em relação à avaliação anterior do mesmo item e não deve "
            "ser interpretada isoladamente como medida de desempenho do avaliador."
        )


# ============================================================
# TRANSIÇÕES E QUALIDADE
# ============================================================

with tab_transicoes:

    st.subheader(
        "O que mudou de uma avaliação para a seguinte"
    )

    c1, c2, c3 = st.columns(3)

    c1.metric(
        "Comparações consecutivas",
        formatar_inteiro(total_transicoes),
        help=(
            f"No filtro atual foram encontrados {formatar_inteiro(qtd_grupos)} itens "
            f"com pelo menos duas notas. "
            f"\n\nO total de {formatar_inteiro(notas_em_itens_repetidos)} notas vem exatamente desta conta: "
            f"\n{calculo_detalhado_notas_multilinha}"
            f"\n\nSomando os subtotais: "
            f"{formatar_inteiro(total_recalculado_notas)} notas."
            f"\n\nCada item é definido por: mesma escola + mesmo avaliador + mesmo critério + mesmo ambiente."
            f"\n\nA lista completa dos {formatar_inteiro(qtd_grupos)} itens pode ser carregada logo abaixo, "
            f"em 'Ver exatamente quais itens formaram as comparações'. "
            f"Nessa tabela, a coluna 'Notas encontradas' mostra exatamente cada nota "
            f"que compõe a quantidade apresentada para o item."
            f"\n\nConta final: "
            f"{formatar_inteiro(notas_em_itens_repetidos)} - "
            f"{formatar_inteiro(qtd_grupos)} = "
            f"{formatar_inteiro(total_transicoes)} comparações consecutivas."
            if qtd_grupos
            else
            "No filtro atual não existem itens com pelo menos duas notas. "
            "Por isso não há comparações consecutivas."
        ),
    )

    c2.metric(
        "Mudaram de categoria",
        formatar_pct(pct_mudancas),
        help=(
            f"Na seleção atual existem {formatar_inteiro(total_transicoes)} comparações. "
            f"Em {formatar_inteiro(qtd_mudancas)} delas a nota mudou em relação à avaliação anterior. "
            f"Assim, {formatar_inteiro(qtd_mudancas)} ÷ "
            f"{formatar_inteiro(total_transicoes)} = {formatar_pct(pct_mudancas)}. "
            "Exemplos de mudança: Bom → Ruim, Ruim → Bom ou Bom → Muito Bom."
            if total_transicoes
            else
            "Não existem comparações suficientes no filtro atual para calcular mudanças de categoria."
        ),
    )

    c3.metric(
        "Mantiveram a categoria",
        formatar_pct(pct_estabilidade),
        help=(
            f"Na seleção atual existem {formatar_inteiro(total_transicoes)} comparações. "
            f"Em {formatar_inteiro(qtd_sem_mudanca)} delas a nota permaneceu igual. "
            f"Assim, {formatar_inteiro(qtd_sem_mudanca)} ÷ "
            f"{formatar_inteiro(total_transicoes)} = {formatar_pct(pct_estabilidade)}. "
            "Exemplos: Bom → Bom ou Muito Bom → Muito Bom."
            if total_transicoes
            else
            "Não existem comparações suficientes no filtro atual para calcular a manutenção da categoria."
        ),
    )

    with st.expander(
        f"Ver exatamente quais itens formaram as comparações ({formatar_inteiro(qtd_grupos)} itens)",
        expanded=False,
    ):
        st.markdown(
            """
            Para manter o dashboard rápido, a tabela detalhada é montada
            somente quando você solicitar. Ela mostra escola, avaliador,
            lote, empresa, critério, ambiente, quantidade de notas,
            as notas na ordem em que ocorreram e as comparações geradas.
            """
        )

        carregar_detalhamento = st.checkbox(
            "Carregar tabela detalhada",
            value=False,
            key="carregar_detalhamento_comparacoes",
        )

        if carregar_detalhamento:
            with st.spinner("Montando detalhamento dos itens..."):
                itens_comparacoes = criar_detalhamento_itens(
                    dados,
                    grupos,
                )

            if not itens_comparacoes.empty:
                st.dataframe(
                    itens_comparacoes,
                    use_container_width=True,
                    hide_index=True,
                )

                st.markdown(
                    f"""
                    **Conferência do cálculo**

                    Total de itens: **{formatar_inteiro(qtd_grupos)}**

                    Soma da coluna "Quantidade de notas":
                    **{formatar_inteiro(notas_em_itens_repetidos)}**

                    Soma da coluna "Comparações geradas":
                    **{formatar_inteiro(total_transicoes)}**

                    Portanto:

                    **{formatar_inteiro(notas_em_itens_repetidos)} notas
                    - {formatar_inteiro(qtd_grupos)} primeiras notas
                    = {formatar_inteiro(total_transicoes)} comparações consecutivas.**
                    """
                )
            else:
                st.info(
                    "Não existem itens com pelo menos duas notas no filtro atual."
                )

    if not transicoes.empty:

        matriz = pd.crosstab(
            transicoes["nota_anterior"],
            transicoes["nota"],
        ).reindex(
            index=NOTAS,
            columns=NOTAS,
            fill_value=0,
        )

        fig = go.Figure(
            data=go.Heatmap(
                z=matriz.values,
                x=matriz.columns,
                y=matriz.index,
                text=matriz.values,
                texttemplate="%{text}",
                hovertemplate=(
                    "Anterior: %{y}<br>"
                    "Atual: %{x}<br>"
                    "Quantidade: %{z}<extra></extra>"
                ),
            )
        )

        fig.update_layout(
            title="Matriz de transição",
            xaxis_title="Nota atual",
            yaxis_title="Nota anterior",
        )

        st.plotly_chart(
            fig,
            use_container_width=True,
            key="grafico_matriz_transicao",
        )

        st.subheader(
            "Transições mais frequentes"
        )

        st.dataframe(
            resumo_transicoes,
            use_container_width=True,
            hide_index=True,
        )

    # (Seção 'Qualidade e cobertura' removida por solicitação.)


# ============================================================
# DOWNLOAD
# ============================================================

st.divider()

st.subheader("Exportar recorte gerencial")

st.caption(
    "Para manter o dashboard rápido, o Excel é gerado somente quando "
    "você clicar em 'Preparar relatório Excel'."
)

assinatura_filtros = (
    tuple(periodo) if isinstance(periodo, (tuple, list)) else str(periodo),
    tuple(escolas_selecionadas),
    tuple(empresas_selecionadas),
    tuple(avaliadores_selecionados),
    tuple(blocos_selecionados),
    tuple(criterios_nomes),
)

if (
    st.session_state.get("excel_assinatura_filtros")
    != assinatura_filtros
):
    st.session_state.pop("excel_bytes", None)
    st.session_state["excel_assinatura_filtros"] = (
        assinatura_filtros
    )


if st.button(
    "Preparar relatório Excel",
    type="secondary",
):
    with st.spinner("Preparando arquivo Excel..."):

        resumo_avaliadores_export = resumo_variacao(
            grupos,
            "avaliador",
            dados,
        )

        if not resumo_avaliadores_export.empty:
            escolas_avaliador = (
                dados
                .groupby(
                    "avaliador",
                    observed=True,
                )["escola"]
                .nunique()
                .rename("qtd_escolas")
                .reset_index()
            )

            resumo_avaliadores_export = (
                resumo_avaliadores_export
                .merge(
                    escolas_avaliador,
                    on="avaliador",
                    how="left",
                )
            )

        resumo_escolas_export = resumo_variacao(
            grupos,
            "escola",
            dados,
        )

        if not resumo_escolas_export.empty:
            avaliadores_escola = (
                dados
                .groupby(
                    "escola",
                    observed=True,
                )["avaliador"]
                .nunique()
                .rename("qtd_avaliadores")
                .reset_index()
            )

            resumo_escolas_export = (
                resumo_escolas_export
                .merge(
                    avaliadores_escola,
                    on="escola",
                    how="left",
                )
            )

        st.session_state["excel_bytes"] = (
            criar_excel_download(
                resumo_geral=resumo_geral_download,
                distribuicao=distribuicao,
                criterios=resumo_criterios,
                blocos=resumo_blocos,
                avaliadores=resumo_avaliadores_export,
                escolas=resumo_escolas_export,
                grupos=grupos,
                transicoes_resumo=resumo_transicoes,
                mensal=mensal,
            )
        )


if "excel_bytes" in st.session_state:
    st.download_button(
        label="Baixar relatório gerencial em Excel",
        data=st.session_state["excel_bytes"],
        file_name="relatorio_gerencial_variacao.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )

st.caption(
    "O arquivo exportado respeita todos os filtros atualmente aplicados "
    "no dashboard."
)
