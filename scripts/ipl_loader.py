"""
scripts/ipl_loader.py
=======================
Calcula o IPL (Índice de Perda por Ligação) por município a partir de
IPL_todos.xls, que traz os 10 componentes brutos da fórmula (não o IPL
já pronto — o SINISA não expõe o indicador 8096 agregado por município
num relatório exportável, só município por município).

FÓRMULA (validada reproduzindo o indicador 8096 oficial de Corumbá,
disponível em mês a mês; erro < 0,3 sobre um índice na casa dos 1000+,
compatível com arredondamento interno do SINISA):

    IPL = 1000 × Σ(Produzido + Importado − Exportado − Uso Operacional
                    − Consumido − Recuperado de Irregularidades)
          ────────────────────────────────────────────────────────────
              Ligações do mês mais recente × Dias acumulados desde janeiro

  - O somatório (Σ) é ACUMULADO desde janeiro do ano corrente até o
    mês mais recente disponível — não é só o mês isolado. Foi assim
    que bateu com o valor oficial do SINISA nos testes.
  - "Dias acumulados desde janeiro" = soma dos dias de cada mês do
    ano corrente até o mês mais recente (ex.: até junho = 31+28+31+30+31+30 = 181).
  - Ligações usa o valor do ÚLTIMO mês (não soma, não é volume — é uma
    contagem de ligações ativas, então faz sentido pegar o valor mais
    recente, não acumular).

Uso operacional = soma dos códigos 29 (descarga de rede) + 30 (quebra
de rede por terceiros) + 31 (limpeza de reservatórios) + 32 (uso pelo
corpo de bombeiros).
"""

import re
import unicodedata
import pandas as pd
from pathlib import Path
from calendar import monthrange

DATA_DIR = Path(__file__).parent.parent / "data"
ARQUIVO_IPL = DATA_DIR / "IPL_todos.xls"

# Códigos SINISA que compõem a fórmula, na ordem em que aparecem no arquivo
CODIGOS = {
    "produzido":  "1-",
    "op_29":      "29-",
    "op_30":      "30-",
    "op_31":      "31-",
    "op_32":      "32-",
    "importado":  "67-",
    "exportado":  "68-",
    "recuperado": "9261-",
    "ligacoes":   "9603-",
    "consumido":  "9642-",
}


def _normalizar(nome: str) -> str:
    nome = str(nome).strip().upper()
    nome = unicodedata.normalize("NFKD", nome).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", nome)


def _dias_acumulados_ate(ano: int, mes: int) -> int:
    """Soma os dias de janeiro até o mês informado (inclusive)."""
    return sum(monthrange(ano, m)[1] for m in range(1, mes + 1))


def _ler_bloco(df: pd.DataFrame, celula_inicio: str, colunas_mes: list, indice_mes_alvo: int) -> dict:
    """
    Lê um bloco de indicador (ex.: '1-VOLUME PRODUZIDO TOTAL') e retorna
    {municipio: soma_acumulada_ate_o_mes_alvo}.

    'indice_mes_alvo' é a posição (0 = primeiro mês da lista de colunas
    do ano corrente, ...) até onde acumular.
    """
    resultado = {}
    dentro_do_bloco = False
    for i in range(len(df)):
        celula = df.iloc[i, 0]
        celula = str(celula).strip() if pd.notna(celula) else ""

        if celula.startswith(celula_inicio):
            dentro_do_bloco = True
            continue
        if dentro_do_bloco and re.match(r"^\d+-", celula):
            break  # começou o próximo bloco de indicador — para por aqui

        if not dentro_do_bloco:
            continue

        nome_municipio = df.iloc[i, 2]
        if pd.isna(nome_municipio):
            continue
        municipio = _normalizar(nome_municipio)

        valores = [df.iloc[i, c] for c in colunas_mes[: indice_mes_alvo + 1]]
        valores = [float(v) for v in valores if pd.notna(v)]
        resultado[municipio] = sum(valores)

    return resultado


def _ler_bloco_ultimo_valor(df: pd.DataFrame, celula_inicio: str, colunas_mes: list, indice_mes_alvo: int) -> dict:
    """Igual a _ler_bloco, mas pega só o valor do mês alvo (não soma) — usado para Ligações."""
    resultado = {}
    dentro_do_bloco = False
    for i in range(len(df)):
        celula = df.iloc[i, 0]
        celula = str(celula).strip() if pd.notna(celula) else ""

        if celula.startswith(celula_inicio):
            dentro_do_bloco = True
            continue
        if dentro_do_bloco and re.match(r"^\d+-", celula):
            break

        if not dentro_do_bloco:
            continue

        nome_municipio = df.iloc[i, 2]
        if pd.isna(nome_municipio):
            continue
        municipio = _normalizar(nome_municipio)

        col_alvo = colunas_mes[indice_mes_alvo]
        valor = df.iloc[i, col_alvo]
        if pd.notna(valor):
            resultado[municipio] = float(valor)

    return resultado


def carregar_ipl(ano_corrente: int = 2026, mes_alvo: int = 6) -> dict:
    """
    Retorna { "AGUA CLARA": 1144.35, ... } com o IPL acumulado de
    janeiro até 'mes_alvo' de 'ano_corrente' (padrão: até junho/2026,
    que é o mês mais recente disponível no export do SINISA hoje).

    Se no futuro o Bruno exportar um mês mais novo (ex.: julho/2026),
    é só chamar carregar_ipl(2026, 7) — desde que as colunas de mês do
    Excel também incluam julho, senão o índice usado vai ser o último
    disponível mesmo (ver 'indice_mes_alvo' abaixo).
    """
    df = pd.read_excel(ARQUIVO_IPL, sheet_name="Informacoes_Municipio", header=None)

    linha_cabecalho = df.iloc[1]
    colunas_mes = [
        c for c in df.columns
        if isinstance(linha_cabecalho[c], str) and re.match(r"^\d{2}/\d{4}$", linha_cabecalho[c])
    ]
    # Filtra só as colunas do ANO CORRENTE (ex.: 01/2026 a 06/2026) —
    # ignora os meses de 2025 que aparecem no início do relatório,
    # já que o acumulado é "desde janeiro do ano corrente".
    colunas_ano_corrente = [
        c for c in colunas_mes
        if linha_cabecalho[c].endswith(f"/{ano_corrente}")
    ]
    # índice do mês alvo dentro dessa lista filtrada (mes_alvo=6 → posição 5, base 0)
    indice_mes_alvo = min(mes_alvo, len(colunas_ano_corrente)) - 1

    componentes = {}
    for chave, prefixo in CODIGOS.items():
        if chave == "ligacoes":
            componentes[chave] = _ler_bloco_ultimo_valor(df, prefixo, colunas_ano_corrente, indice_mes_alvo)
        else:
            componentes[chave] = _ler_bloco(df, prefixo, colunas_ano_corrente, indice_mes_alvo)

    dias_acumulados = _dias_acumulados_ate(ano_corrente, mes_alvo)

    todos_municipios = set(componentes["produzido"])
    resultado = {}
    for municipio in todos_municipios:
        try:
            produzido  = componentes["produzido"].get(municipio, 0.0)
            importado  = componentes["importado"].get(municipio, 0.0)
            exportado  = componentes["exportado"].get(municipio, 0.0)
            op         = (componentes["op_29"].get(municipio, 0.0)
                          + componentes["op_30"].get(municipio, 0.0)
                          + componentes["op_31"].get(municipio, 0.0)
                          + componentes["op_32"].get(municipio, 0.0))
            consumido  = componentes["consumido"].get(municipio, 0.0)
            recuperado = componentes["recuperado"].get(municipio, 0.0)
            ligacoes   = componentes["ligacoes"].get(municipio)

            if not ligacoes:  # None ou 0 — não dá pra dividir
                continue

            numerador = 1000 * (produzido + importado - exportado - op - consumido - recuperado)
            denominador = ligacoes * dias_acumulados
            resultado[municipio] = numerador / denominador
        except (TypeError, ZeroDivisionError):
            continue  # município com dado faltante — fica de fora em vez de quebrar tudo

    return resultado
