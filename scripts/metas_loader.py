"""
scripts/metas_loader.py
=========================
Carrega as METAS contratuais (Contrato de Programa) por município,
a partir de Metas_Contrato_de_Programa_Atual_2025.xlsx.

Cada indicador tem sua própria aba com o mesmo formato:
  ORDEM | MUNICÍPIO | ATUAL | 2029 | 2033 | 2037 | 2041 | 2045 | 2049

Por enquanto usamos só a coluna ATUAL (meta vigente hoje) — as colunas
de horizonte (2029...2049) ficam pra quando fizermos a tela de Projeções.

Sobre o operador de comparação (>=, >, <):
  A planilha não repete o operador em cada linha — ele aparece só uma
  vez no cabeçalho da aba "TABELA GERAL" (ex.: '>= 99', '> 90', '< 216').
  Conferimos manualmente que o operador é FIXO por indicador (nunca
  muda de município pra município), então hardcodamos aqui. Se um dia
  a SANESUL mudar essa regra contratual, é só ajustar o dicionário
  OPERADOR_POR_INDICADOR abaixo.
"""

import re
import unicodedata
import pandas as pd
from pathlib import Path

DATA_DIR = Path(__file__).parent.parent / "data"
ARQUIVO_METAS = DATA_DIR / "Metas_Contrato_de_Programa_Atual_2025.xlsx"

# Nome interno do indicador -> (aba no Excel, operador da meta, rótulo bonito)
INDICADORES = {
    "agua":       {"aba": "ÁGUA",        "operador": ">=", "label": "Cobertura de Água"},
    "esgoto":     {"aba": "ESGOTO",      "operador": ">",  "label": "Cobertura de Esgoto"},
    "ipl":        {"aba": "PERDAS",      "operador": "<",  "label": "IPL (Perdas)"},
    "tratamento": {"aba": "TRAT-ESGOTO", "operador": ">=", "label": "Tratamento de Esgoto"},
    "iqa":        {"aba": "IQA",         "operador": ">",  "label": "IQA"},
    "dbo":        {"aba": "DBO",         "operador": ">=", "label": "DBO (Remoção %)"},
}


def _normalizar(nome: str) -> str:
    nome = str(nome).strip().upper()
    nome = unicodedata.normalize("NFKD", nome).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", nome)


# Correções manuais para municípios cujo nome vem escrito de forma
# diferente na planilha de metas do contrato em relação à base SINISA
# (a normalização de acento/maiúscula sozinha não resolve esses casos —
# são erro de digitação ou nome oficial completo vs. nome curto).
# Achados ao comparar as duas fontes: "Bataypora" (SINISA) x "Bataipora" (metas)
# e "Rio Verde" (SINISA) x "Rio Verde De Mato Grosso" (metas).
APELIDOS_MUNICIPIO = {
    "BATAYPORA": "BATAIPORA",
    "RIO VERDE DE MATO GROSSO": "RIO VERDE",
}


def _normalizar_com_apelido(nome: str) -> str:
    """Normaliza e, se for um nome conhecido com grafia divergente, já
    devolve a versão usada na base SINISA — para casar corretamente
    com carregar_contratos() no comparacao_loader.py."""
    normalizado = _normalizar(nome)
    return APELIDOS_MUNICIPIO.get(normalizado, normalizado)


def _ler_aba(aba: str) -> dict:
    """Lê uma aba de indicador e devolve {municipio_normalizado: valor_meta_atual}."""
    df = pd.read_excel(ARQUIVO_METAS, sheet_name=aba, header=None)
    resultado = {}
    for i in range(2, len(df)):  # linha 0 = título da aba, linha 1 = cabeçalho
        nome_municipio = df.iloc[i, 2]
        valor_atual = df.iloc[i, 3]
        if pd.isna(nome_municipio) or pd.isna(valor_atual):
            continue
        resultado[_normalizar_com_apelido(nome_municipio)] = float(valor_atual)
    return resultado


def carregar_metas() -> dict:
    """
    Retorna:
      { "AGUA CLARA": {
          "agua":       {"meta": 99.0, "operador": ">=", "label": "Cobertura de Água"},
          "esgoto":     {"meta": 0.0,  "operador": ">",  "label": "Cobertura de Esgoto"},
          ...
        }, ... }
    """
    dados_por_indicador = {
        chave: _ler_aba(info["aba"]) for chave, info in INDICADORES.items()
    }

    todos_municipios = set()
    for dados in dados_por_indicador.values():
        todos_municipios |= set(dados)

    resultado = {}
    for municipio in todos_municipios:
        resultado[municipio] = {}
        for chave, info in INDICADORES.items():
            meta = dados_por_indicador[chave].get(municipio)
            resultado[municipio][chave] = {
                "meta": meta,
                "operador": info["operador"],
                "label": info["label"],
            }
    return resultado
