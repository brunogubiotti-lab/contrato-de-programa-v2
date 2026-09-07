"""
scripts/comparacao_loader.py
==============================
Junta o valor REAL (SINISA, scripts/contratos_loader.py) com a META
contratual (scripts/metas_loader.py) por município e por indicador,
e calcula se o município está CUMPRINDO ou NÃO a meta.

Como funciona a comparação:
  O operador (>=, >, <) usado para decidir "cumpre" ou "não cumpre" vem
  direto da meta carregada por metas_loader.carregar_metas() — não é
  fixo neste arquivo. Ex.: Cobertura de Água normalmente usa '>=', então
  cumpre se REAL >= META. IPL normalmente usa '<' (quanto menor a perda,
  melhor). Mas o operador pode variar entre municípios ou entre o valor
  "atual" e a meta final do mesmo indicador (já visto em Novo Horizonte
  do Sul, por exemplo) — por isso não hardcodamos aqui.

HISTÓRICO — DBO e IPL já têm dado real:
  Até uma versão anterior deste arquivo, "dbo" e "ipl" sempre voltavam
  "sem_dado" porque faltava a fórmula de remoção de DBO e o cálculo de
  IPL (ver ipl_loader.py). Isso já foi implementado — INDICADORES_SEM_REAL_AINDA
  está vazio hoje (linha abaixo) e ambos são comparados normalmente.
  Se algum indicador precisar voltar a ficar de fora da comparação (por
  falta de dado real), inclua sua chave nesse set — é o único lugar que
  precisa mudar.
"""

from scripts.contratos_loader import carregar_contratos
from scripts.metas_loader import carregar_metas, INDICADORES

# Todos os indicadores já têm valor real desde a implementação do
# ipl_loader.py (IPL) e da fórmula de remoção de DBO em contratos_loader.py.
INDICADORES_SEM_REAL_AINDA = set()

# Mapeia o nome do indicador (usado nas metas) para o campo correspondente
# no dict que vem de carregar_contratos().
CAMPO_REAL = {
    "agua": "cobertura_agua",
    "esgoto": "cobertura_esgoto",
    "ipl": "ipl",
    "tratamento": "tratamento_esgoto",
    "iqa": "iqa",
    "dbo": "dbo5",
}


def _cumpre_meta(real: float | None, meta: float | None, operador: str) -> str:
    """Retorna 'cumpre', 'nao_cumpre' ou 'sem_dado'."""
    if real is None or meta is None:
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
    Retorna um dict por município, já cruzando real x meta x status:

      { "Agua Clara": {
          "agua":   {"real": 99.0, "meta": 99.0, "operador": ">=",
                     "label": "Cobertura de Água", "status": "cumpre"},
          "dbo":    {"real": 0.0, "meta": 60.0, "operador": ">=",
                     "label": "DBO (Remoção %)", "status": "sem_dado"},
          ...
        }, ... }
    """
    reais = {c["municipio"].upper(): c for c in carregar_contratos()}
    metas = carregar_metas()

    resultado = {}
    for municipio_norm, indicadores_meta in metas.items():
        # carregar_contratos() já devolve nome em Title Case (ex: "Agua Clara");
        # carregar_metas() normaliza em CAIXA ALTA sem acento. Usamos o nome
        # "bonito" do lado real como chave, quando existir; senão caímos no
        # nome normalizado mesmo (ainda em caixa alta).
        real_dados = reais.get(municipio_norm)
        nome_exibicao = real_dados["municipio"] if real_dados else municipio_norm.title()

        if real_dados is None:
            # Sinal de que o nome desse município na planilha de metas não
            # bate com nenhum nome da base SINISA (real). Já aconteceu com
            # "Bataypora"/"Bataipora" e "Rio Verde De Mato Grosso"/"Rio Verde" —
            # corrigidos em metas_loader.APELIDOS_MUNICIPIO. Se aparecer de
            # novo aqui, é sinal de outro caso parecido que falta mapear.
            print(f"[comparacao_loader] Aviso: '{nome_exibicao}' tem meta mas nenhum dado real "
                  f"correspondente — confira se o nome bate com a base SINISA "
                  f"(ver metas_loader.APELIDOS_MUNICIPIO).")

        linha = {}
        for chave, info_meta in indicadores_meta.items():
            campo = CAMPO_REAL[chave]
            valor_real = real_dados.get(campo) if real_dados else None

            if chave in INDICADORES_SEM_REAL_AINDA:
                status = "sem_dado"
            else:
                status = _cumpre_meta(valor_real, info_meta["meta"], info_meta["operador"])

            linha[chave] = {
                "real": valor_real,
                "meta": info_meta["meta"],
                "operador": info_meta["operador"],
                "label": info_meta["label"],
                "status": status,
            }
        resultado[nome_exibicao] = linha

    return resultado
