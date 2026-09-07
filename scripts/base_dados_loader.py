"""
scripts/base_dados_loader.py — SISPLAN v2
==============================================
Loader único para o módulo Contratos. Lê UMA planilha só
(data/base_dados_sisplan.xlsx), com 6 abas:

  - dados_reais        : um município por linha, valores REAIS atuais
  - metas_contrato     : uma linha por (município, indicador), meta atual
  - historico_mensal   : uma linha por (município, indicador), jan-jun/2026
  - evolucao_metas     : uma linha por (município, indicador), 2029-2049
  - contratos_programa : um município por linha, dados cadastrais do
                         Contrato de Programa e do Convênio AGEMS
  - historico_aditivos : um aditivo/revisão por linha (vários por município)

As duas últimas abas alimentam a tela "Dados Gerais" e vieram da
extração oficial dos 68 PDFs do SIGIS (scripts/extrair_contratos_pdf.py).

═══════════════════════════════════════════════════════════════════
PERFORMANCE: cache em memória (o ponto principal deste arquivo)
═══════════════════════════════════════════════════════════════════
Antes: cada tela (Por Município, Por Indicador, Análise Geral) abria
e reprocessava o Excel do ZERO a cada clique — medimos ~250ms por
navegação só nisso, porque abrir um .xlsx é lento (é um zip com XML
dentro, não um formato binário direto como Parquet).

Como os dados só mudam quando alguém troca o arquivo manualmente
(não a cada request), não faz sentido reler o disco toda hora. Agora:
  1. As 4 abas são lidas UMA VEZ (na primeira chamada, "lazy") usando
     pd.ExcelFile — que abre o arquivo .xlsx uma única vez e lê todas
     as abas dele, em vez de pd.read_excel() 4 vezes seguidas (cada
     chamada de pd.read_excel abre e fecha o arquivo do zero).
  2. Os resultados computados (comparação real x meta, evolução das
     metas) também ficam em cache — não é só o Excel bruto que fica
     memorizado, é o resultado FINAL já pronto pra tela.
  3. Se o arquivo for atualizado (novo deploy com xlsx novo), o cache
     é invalidado automaticamente comparando a data de modificação do
     arquivo (mtime) — não precisa reiniciar nada manualmente.

Resultado: a primeira requisição depois de o servidor subir paga o
custo de ~250ms uma única vez; todas as seguintes ficam na casa de
frações de milissegundo (é só um dicionário Python já pronto).
"""

import pandas as pd
from pathlib import Path

DATA_DIR = Path(__file__).parent.parent / "data"
ARQUIVO_BASE = DATA_DIR / "base_dados_sisplan.xlsx"

INDICADORES = {
    "agua":       {"label": "Cobertura de Água"},
    "esgoto":     {"label": "Cobertura de Esgoto"},
    "ipl":        {"label": "IPL (Perdas)"},
    "tratamento": {"label": "Tratamento de Esgoto"},
    "iqa":        {"label": "IQA"},
    "dbo":        {"label": "DBO (Remoção %)"},
}

CAMPO_REAL = {
    "agua": "cobertura_agua",
    "esgoto": "cobertura_esgoto",
    "ipl": "ipl",
    "tratamento": "tratamento_esgoto",
    "iqa": "iqa",
    "dbo": "dbo5",
}


# ── Cache em memória — ver docstring do módulo ──────────────────────
_cache = {
    "mtime": None,       # data de modificação do arquivo na última leitura
    "dfs": None,          # dict {nome_aba: DataFrame}
    "comparacao": None,   # resultado pronto de carregar_comparacao()
    "evolucao": None,     # resultado pronto de carregar_evolucao_metas()
    "metas_lookup": None, # {(MUNICIPIO_MAIUSCULO, indicador): {"meta":..., "operador":...}}
}


def _garantir_cache_atualizado() -> None:
    """
    Recarrega o Excel só se: (a) é a primeira vez que alguém pede um
    dado, ou (b) o arquivo no disco mudou desde a última leitura
    (comparando mtime). Caso contrário, não toca no disco — usa o
    que já está em memória.
    """
    mtime_atual = ARQUIVO_BASE.stat().st_mtime

    if _cache["dfs"] is not None and _cache["mtime"] == mtime_atual:
        return  # cache válido, nada a fazer

    # pd.ExcelFile abre o .xlsx UMA vez; .parse() por aba reaproveita
    # esse mesmo handle já aberto, em vez de reabrir o arquivo a cada
    # pd.read_excel(caminho, sheet_name=...) como fazíamos antes.
    with pd.ExcelFile(ARQUIVO_BASE) as excel:
        dfs = {
            "dados_reais": excel.parse("dados_reais"),
            "metas_contrato": excel.parse("metas_contrato"),
            "historico_mensal": excel.parse("historico_mensal"),
            "evolucao_metas": excel.parse("evolucao_metas"),
            "contratos_programa": excel.parse("contratos_programa"),
            "historico_aditivos": excel.parse("historico_aditivos"),
        }

    _cache["dfs"] = dfs
    _cache["mtime"] = mtime_atual
    _cache["comparacao"] = None  # invalida os resultados computados também
    _cache["evolucao"] = None
    _cache["metas_lookup"] = None


def _obter_metas_lookup() -> dict:
    """
    Monta (uma vez, cacheado) um lookup {(MUNICIPIO_MAIUSCULO, indicador):
    {"meta":..., "operador":...}} a partir da aba 'metas_contrato' — a
    ÚNICA fonte de verdade para meta/operador no sistema.

    Por quê isso existe: a aba 'historico_mensal' também tem suas
    próprias colunas 'meta'/'operador' (redundantes), e na prática elas
    ficam desatualizadas quando alguém atualiza só a aba 'metas_contrato'
    (foi o que aconteceu em ago/2026 — 85 combinações de município x
    indicador ficaram com meta antiga em historico_mensal enquanto
    metas_contrato já tinha o valor novo). Em vez de reeducar quem
    atualiza a planilha a sempre editar as duas abas, o código passa a
    ignorar meta/operador de historico_mensal e sempre sobrescrever pelo
    valor de metas_contrato — assim as duas abas nunca mais divergem,
    mesmo que a de histórico fique com dado velho.
    """
    if _cache["metas_lookup"] is not None:
        return _cache["metas_lookup"]

    df_metas = _cache["dfs"]["metas_contrato"]
    lookup = {}
    for _, row in df_metas.iterrows():
        chave = (row["municipio"].upper(), row["indicador"])
        lookup[chave] = {"meta": row["meta"], "operador": row["operador"]}

    _cache["metas_lookup"] = lookup
    return lookup


def _cumpre_meta(real: float | None, meta: float | None, operador: str, indicador: str | None = None) -> str:
    """
    Regra especial para DBO: quando não existe valor real de remoção
    de DBO, é porque o município ainda não tem cobertura de esgoto
    (não há o que tratar, logo não há remoção nenhuma). Isso é
    reprovação, não "falta de dado" — por isso classificamos como
    "nao_cumpre" em vez de "sem_dado" só pra esse indicador.
    """
    if real is None or pd.isna(real):
        if indicador == "dbo":
            return "nao_cumpre"
        return "sem_dado"
    if meta is None or pd.isna(meta):
        return "sem_dado"
    comparacoes = {
        ">=": real >= meta,
        ">":  real > meta,
        "<":  real < meta,
        "<=": real <= meta,
    }
    return "cumpre" if comparacoes.get(operador, False) else "nao_cumpre"


def carregar_comparacao() -> dict:
    """
    Retorna a comparação real x meta por município x indicador:

      { "Agua Clara": {
          "agua": {"real": 99.0, "meta": 99.0, "operador": ">=",
                    "label": "Cobertura de Água", "status": "cumpre"},
          ...
        }, ... }

    Computado uma vez e reaproveitado (ver cache no topo do arquivo).
    """
    _garantir_cache_atualizado()
    if _cache["comparacao"] is not None:
        return _cache["comparacao"]

    df_reais = _cache["dfs"]["dados_reais"]
    df_metas = _cache["dfs"]["metas_contrato"]

    reais_por_municipio = {row["municipio"]: row for _, row in df_reais.iterrows()}

    resultado = {}
    for _, linha_meta in df_metas.iterrows():
        municipio = linha_meta["municipio"].title()
        indicador = linha_meta["indicador"]

        row_real = reais_por_municipio.get(municipio)
        campo = CAMPO_REAL[indicador]
        valor_real = row_real[campo] if row_real is not None else None
        if valor_real is not None and pd.isna(valor_real):
            valor_real = None

        meta = linha_meta["meta"]
        if pd.isna(meta):
            meta = None

        resultado.setdefault(municipio, {})[indicador] = {
            "real": valor_real,
            "meta": meta,
            "operador": linha_meta["operador"],
            "label": linha_meta["label"],
            "status": _cumpre_meta(valor_real, meta, linha_meta["operador"], indicador),
        }

    _cache["comparacao"] = resultado
    return resultado


def carregar_evolucao_metas() -> dict:
    """
    Evolução da meta contratual por ano-alvo:

      { ("Agua Clara", "agua"): [("2029", 99.0), ("2033", 99.0), ...], ... }

    Desde ago/2026, 'evolucao_metas' está em formato LONGO — uma linha
    por (município, indicador, ano) — em vez de 6 colunas fixas
    (y2029, y2033, ..., y2049) iguais pra todo mundo. Motivo: os anos
    de revisão contratual variam por município (ex.: alguns têm 2031,
    2038, 2039... que não existiam nas colunas fixas antigas) — forçar
    isso em colunas fixas ou perdia ano ou misturava anos diferentes
    na mesma coluna. Formato longo aceita qualquer ano sem mudar a
    estrutura da planilha.

    Meta 0 ainda é omitida por segurança (mesma regra de antes: pode
    indicar fim de contrato antes desse horizonte) — hoje a fonte nova
    não usa mais esse padrão, mas o filtro não faz mal manter.
    """
    _garantir_cache_atualizado()
    if _cache["evolucao"] is not None:
        return _cache["evolucao"]

    df = _cache["dfs"]["evolucao_metas"]

    resultado = {}
    for _, row in df.iterrows():
        valor = row["valor"]
        if pd.isna(valor) or valor == 0:
            continue
        chave = (row["municipio"], row["indicador"])
        resultado.setdefault(chave, []).append((str(int(row["ano"])), float(valor)))

    # garante ordem cronológica dentro de cada município/indicador
    for chave in resultado:
        resultado[chave].sort(key=lambda ponto: ponto[0])

    _cache["evolucao"] = resultado
    return resultado


def carregar_historico_municipio(municipio: str) -> list[dict]:
    """
    Histórico mensal (jan a jun/2026) de todos os indicadores de UM
    município. Filtra o DataFrame já em cache — não toca no disco.

    O campo 'meta' de cada linha é sobrescrito pelo valor vindo de
    metas_contrato (ver _obter_metas_lookup) em vez do valor que já vem
    dentro de historico_mensal — essa última coluna existe na planilha
    mas fica desatualizada quando só metas_contrato é editada, então
    não confiamos mais nela.
    """
    _garantir_cache_atualizado()
    df = _cache["dfs"]["historico_mensal"]
    df_municipio = df[df["municipio"] == municipio]
    metas_lookup = _obter_metas_lookup()

    linhas = []
    for _, row in df_municipio.iterrows():
        linha = row.to_dict()
        for campo in ["meta", "jan", "fev", "mar", "abr", "mai", "jun"]:
            if campo in linha and pd.isna(linha[campo]):
                linha[campo] = None

        meta_correta = metas_lookup.get((municipio.upper(), linha["indicador"]))
        if meta_correta is not None:
            linha["meta"] = meta_correta["meta"]
            linha["operador"] = meta_correta["operador"]

        linhas.append(linha)
    return linhas


def _vazio_para_none(valor):
    """
    A aba contratos_programa usa NaN (pandas) ou strings tipo '--' / '-- 0'
    pra indicar campo ausente (município sem Contrato de Programa vigente
    — ex.: Aparecida do Taboado). Padroniza tudo pra None, mais fácil de
    checar no template com {% if %}.
    """
    if pd.isna(valor):
        return None
    if isinstance(valor, str) and valor.strip().replace("0", "").strip("- ") == "":
        return None
    return valor


def carregar_dados_gerais() -> dict:
    """
    Dados "cadastrais" do Contrato de Programa por município — número,
    datas, duração e Convênio de Cooperação AGEMS — para a tela Dados
    Gerais. Filtra o DataFrame já em cache — não toca no disco.

    Retorna { "ITAPORA": {
        "municipio": "ITAPORA",
        "num_contrato_programa": "006/2008",
        "data_assinatura_cp": "18/12/2008",
        "duracao_anos_cp": 30.0,
        "data_expiracao_cp": "18/12/2038",
        "num_convenio_agems": "006/2008",
        "data_assinatura_convenio": "18/12/2008",
        "data_expiracao_convenio": "18/12/2038",
        "tem_contrato": True,
      }, ... }

    "tem_contrato" é um atalho pro template: False quando o município
    não tem Contrato de Programa vigente (demais campos vêm None).
    """
    _garantir_cache_atualizado()
    df = _cache["dfs"]["contratos_programa"]

    resultado = {}
    for _, row in df.iterrows():
        municipio = row["municipio"].title()
        num_contrato = _vazio_para_none(row["num_contrato_programa"])
        resultado[municipio] = {
            "municipio": municipio,
            "num_contrato_programa": num_contrato,
            "data_assinatura_cp": _vazio_para_none(row["data_assinatura_cp"]),
            "duracao_anos_cp": _vazio_para_none(row["duracao_anos_cp"]),
            "data_expiracao_cp": _vazio_para_none(row["data_expiracao_cp"]),
            "num_convenio_agems": _vazio_para_none(row["num_convenio_agems"]),
            "data_assinatura_convenio": _vazio_para_none(row["data_assinatura_convenio"]),
            "data_expiracao_convenio": _vazio_para_none(row["data_expiracao_convenio"]),
            "tem_contrato": num_contrato is not None,
        }
    return resultado


def carregar_aditivos_municipio(municipio: str) -> list[dict]:
    """
    Histórico de aditivos/revisões de UM município, em ordem
    cronológica (mais antigo primeiro). Filtra o DataFrame já em
    cache — não toca no disco.
    """
    _garantir_cache_atualizado()
    df = _cache["dfs"]["historico_aditivos"]
    sub = df[df["municipio"].str.title() == municipio].copy()
    if sub.empty:
        return []

    # a coluna "data" é string "dd/mm/aaaa" — convertemos só pra
    # ordenar, mas devolvemos a string original (evita problema de
    # timezone/formato ao exibir no template).
    sub["data_ord"] = pd.to_datetime(sub["data"], format="%d/%m/%Y", errors="coerce")
    sub = sub.sort_values("data_ord")

    return sub[["instrumento", "data"]].to_dict("records")
