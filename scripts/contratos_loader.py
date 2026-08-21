"""
scripts/contratos_loader.py
============================
Carrega os dados da aba "Lista de Contratos" a partir de:
  - data/Indicadores01_0001.xls  (Cobertura Água, Cobertura Esgoto, IQA)
  - data/Informações01_0001.xls  (DBO afluente/efluente da ETE)
  - data/contratos_numero.csv    (número do contrato — fonte temporária,
                                   extraída manualmente do PDF exemplo até
                                   existir uma origem melhor)

Formato dos .xls do SINISA (por que o parsing é assim):
  Cada planilha vem como um "relatório impresso" exportado do sistema,
  não como uma tabela normal. A estrutura é:

    <código>-<REGIONAL>                    ← linha de regional (ignorar)
    <código>-<MUNICÍPIO>                   ← início do bloco do município
    8300-COBERTURA DE ÁGUA - AGEMS  Un.  06/2025  07/2025 ... 06/2026
    8301-COBERTURA DE ESGOTO - AGEMS ...
    8450-INDICE QUALIDADE DA AGUA (IQA) ...

  Ou seja, não dá pra usar pd.read_excel "normal" com cabeçalho fixo:
  precisamos varrer linha a linha, identificando pelo padrão do texto
  na coluna 0 se é município, regional ou indicador.

Qual mês usamos como "valor atual"?
  Pegamos a ÚLTIMA coluna de mês disponível (a mais recente — hoje é
  06/2026). Se no futuro quiser trocar para média do período, é só
  mudar MODO_VALOR abaixo.
"""

import re
import unicodedata
import pandas as pd
from pathlib import Path

from scripts.ipl_loader import carregar_ipl

DATA_DIR = Path(__file__).parent.parent / "data"

# Trocar para "media" se um dia quiser usar a média do período em vez
# do mês mais recente.
MODO_VALOR = "ultimo_mes"

# Códigos dos indicadores que nos interessam nesta tela
COD_COBERTURA_AGUA = "8300"
COD_COBERTURA_ESGOTO = "8301"
COD_IQA = "8450"
COD_DBO_AFLUENTE = "8402"
COD_DBO_EFLUENTE = "8403"

TRATAMENTO_ESGOTO_FIXO = 100.0  # confirmado com o Bruno: sempre 100


def _calcular_remocao_dbo(afluente: float | None, efluente: float | None) -> float | None:
    """
    Remoção (%) = (Le - Ls) / Le, onde Le = DBO na entrada da ETE
    (afluente) e Ls = DBO na saída (efluente). Fórmula confirmada
    pelo Bruno (documento técnico da SANESUL).

    Se afluente for None, 0, ou o município não tratar esgoto (ambos
    zerados — ver Água Clara, cobertura de esgoto 0%), não dá pra
    calcular remoção: devolve None em vez de gerar ZeroDivisionError
    ou um "100% de remoção" enganoso.
    """
    if afluente is None or efluente is None or afluente == 0:
        return None
    return (afluente - efluente) / afluente * 100


def _normalizar(nome: str) -> str:
    """
    Remove acentos e padroniza maiúsculas/espaços, para conseguir
    casar 'ANGÉLICA' (se aparecer assim em algum arquivo) com
    'ANGELICA' (como está no CSV do PDF), por exemplo.
    """
    nome = str(nome).strip().upper()
    nome = unicodedata.normalize("NFKD", nome).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", nome)


def _e_linha_municipio(valor: str) -> bool:
    """
    Linha de município tem o padrão '<código numérico>-<NOME>',
    ex.: '25020-AGUA CLARA'. Diferencia de:
      - linha de regional: '20-Regional de Dourados' (tem a palavra Regional)
      - linha de indicador: '8300-COBERTURA...' (código sempre começa com 8)
    """
    m = re.match(r"^(\d+)-(.+)$", valor)
    if not m:
        return False
    codigo, resto = m.groups()
    if "regional" in resto.lower():
        return False
    if codigo.startswith("8"):  # códigos de indicador começam com 8xxx
        return False
    return True


def _extrair_ultimo_valor(linha: pd.Series, colunas_mes: list) -> float | None:
    """Pega o valor da última coluna de mês que não seja NaN."""
    for col in reversed(colunas_mes):
        val = linha[col]
        if pd.notna(val):
            try:
                return float(val)
            except (TypeError, ValueError):
                continue
    return None


def _carregar_indicadores() -> dict:
    """
    Lê Indicadores01_0001.xls e retorna:
      { "AGUA CLARA": {"cobertura_agua": 99.0, "cobertura_esgoto": 0.0, "iqa": 98.04}, ... }
    """
    caminho = DATA_DIR / "Indicadores01_0001.xls"
    df = pd.read_excel(caminho, sheet_name="Indicadores_Municipio", header=None)

    # Linha 3 (índice 3) tem os cabeçalhos de mês, a partir da coluna 5
    # (colunas 0-4 são código/nome/unidade). Pegamos só as colunas que
    # têm um cabeçalho tipo "MM/AAAA".
    linha_cabecalho = df.iloc[3]
    colunas_mes = [
        c for c in df.columns
        if isinstance(linha_cabecalho[c], str) and re.match(r"^\d{2}/\d{4}$", linha_cabecalho[c])
    ]

    resultado = {}
    municipio_atual = None

    for i in range(4, len(df)):
        celula = df.iloc[i, 0]
        if pd.isna(celula):
            continue
        celula = str(celula).strip()

        if _e_linha_municipio(celula):
            nome = celula.split("-", 1)[1]
            municipio_atual = _normalizar(nome)
            resultado[municipio_atual] = {
                "cobertura_agua": None,
                "cobertura_esgoto": None,
                "iqa": None,
            }
            continue

        if municipio_atual is None:
            continue  # ainda não achamos o primeiro município (cabeçalho/regional)

        linha = df.iloc[i]
        valor = _extrair_ultimo_valor(linha, colunas_mes)

        if celula.startswith(COD_COBERTURA_AGUA + "-"):
            resultado[municipio_atual]["cobertura_agua"] = valor
        elif celula.startswith(COD_COBERTURA_ESGOTO + "-"):
            resultado[municipio_atual]["cobertura_esgoto"] = valor
        elif celula.startswith(COD_IQA + "-") and "INDICE" in celula.upper():
            resultado[municipio_atual]["iqa"] = valor

    return resultado


def _carregar_dbo() -> dict:
    """
    Lê Informações01_0001.xls e retorna:
      { "AGUA CLARA": {"dbo_afluente": 0.0, "dbo_efluente": ...}, ... }

    Esse arquivo tem uma estrutura diferente do de indicadores: cada
    bloco de indicador (8402, 8403) já vem com TODOS os municípios
    listados em sequência (uma linha por município), em vez de vir
    agrupado por município como no arquivo de Indicadores.
    """
    caminho = DATA_DIR / "Informações01_0001.xls"
    df = pd.read_excel(caminho, sheet_name="Informacoes_Municipio", header=None)

    linha_cabecalho = df.iloc[1]
    colunas_mes = [
        c for c in df.columns
        if isinstance(linha_cabecalho[c], str) and re.match(r"^\d{2}/\d{4}$", linha_cabecalho[c])
    ]
    col_municipio = 2  # confirmado pela inspeção: coluna "Município"

    resultado = {}
    bloco_atual = None  # "afluente" ou "efluente"

    for i in range(2, len(df)):
        celula0 = df.iloc[i, 0]
        celula0 = str(celula0).strip() if pd.notna(celula0) else ""

        if celula0.startswith("8402-"):
            bloco_atual = "dbo_afluente"
            continue
        if celula0.startswith("8403-"):
            bloco_atual = "dbo_efluente"
            continue

        if bloco_atual is None:
            continue

        nome_municipio = df.iloc[i, col_municipio]
        if pd.isna(nome_municipio):
            continue
        nome_municipio = _normalizar(nome_municipio)

        linha = df.iloc[i]
        valor = _extrair_ultimo_valor(linha, colunas_mes)

        resultado.setdefault(nome_municipio, {"dbo_afluente": None, "dbo_efluente": None})
        resultado[nome_municipio][bloco_atual] = valor

    return resultado


def _carregar_numeros_contrato() -> dict:
    caminho = DATA_DIR / "contratos_numero.csv"
    df = pd.read_csv(caminho, dtype=str, keep_default_na=False)
    return {
        _normalizar(row["municipio"]): row["numero_contrato"]
        for _, row in df.iterrows()
    }


def carregar_contratos() -> list[dict]:
    """
    Função principal chamada pela rota. Junta os 4 arquivos por
    município e devolve uma lista de dicts pronta para a tabela:

      [{"municipio": "AGUA CLARA", "numero_contrato": "001/2020",
        "cobertura_agua": 99.0, "cobertura_esgoto": 0.0,
        "tratamento_esgoto": 100.0, "ipl": 143.64,
        "iqa": 98.04, "dbo5": 0.0}, ...]

    Sobre o IPL: calculado em scripts/ipl_loader.py a partir dos 10
    componentes brutos do SINISA (IPL_todos.xls), fórmula validada
    reproduzindo o indicador 8096 oficial (ver docstring do loader).
    """
    indicadores = _carregar_indicadores()
    dbo = _carregar_dbo()
    numeros = _carregar_numeros_contrato()
    ipl = carregar_ipl()

    # União de todos os municípios encontrados em qualquer uma das fontes,
    # ordenada alfabeticamente (mesma ordem do PDF exemplo).
    todos_municipios = sorted(set(indicadores) | set(dbo) | set(numeros) | set(ipl))

    linhas = []
    for municipio in todos_municipios:
        ind = indicadores.get(municipio, {})
        d = dbo.get(municipio, {})
        linhas.append({
            "municipio": municipio.title(),
            "numero_contrato": numeros.get(municipio) or "—",
            "cobertura_agua": ind.get("cobertura_agua"),
            "cobertura_esgoto": ind.get("cobertura_esgoto"),
            "tratamento_esgoto": TRATAMENTO_ESGOTO_FIXO,
            "ipl": ipl.get(municipio),
            "iqa": ind.get("iqa"),
            "dbo5": _calcular_remocao_dbo(d.get("dbo_afluente"), d.get("dbo_efluente")),
        })

    return linhas
