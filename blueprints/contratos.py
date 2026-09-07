"""
blueprints/contratos.py — SISPLAN v2
========================================================
Módulo Contratos de Programa: Por Município, Por Indicador e
Análise Geral (em construção).
"""

from flask import Blueprint, render_template, request

from scripts.base_dados_loader import (
    carregar_comparacao,
    carregar_historico_municipio,
    carregar_evolucao_metas,
    carregar_dados_gerais,
    carregar_aditivos_municipio,
    INDICADORES,
)


contratos_bp = Blueprint("contratos", __name__, url_prefix="/contratos")


@contratos_bp.route("/dados-gerais")
def dados_gerais():
    """
    Ficha "cadastral" do Contrato de Programa: número, datas, duração,
    convênio AGEMS e histórico de aditivos — um município por vez,
    mesmo padrão de seleção da tela Por Município.
    """
    todos_dados = carregar_dados_gerais()
    municipios_disponiveis = sorted(todos_dados.keys())

    municipio_selecionado = request.args.get("municipio") or municipios_disponiveis[0]
    dados_municipio = todos_dados.get(municipio_selecionado, {})
    aditivos = carregar_aditivos_municipio(municipio_selecionado)

    return render_template(
        "contratos/dados_gerais.html",
        municipios=municipios_disponiveis,
        municipio_selecionado=municipio_selecionado,
        dados=dados_municipio,
        aditivos=aditivos,
        pagina_ativa="contratos",
        sub_ativa="dados_gerais",
    )


@contratos_bp.route("/por-municipio")
def por_municipio():
    comparacao = carregar_comparacao()
    municipios_disponiveis = sorted(comparacao.keys())

    municipio_selecionado = request.args.get("municipio") or municipios_disponiveis[0]
    dados_municipio = comparacao.get(municipio_selecionado, {})
    historico = carregar_historico_municipio(municipio_selecionado)

    return render_template(
        "contratos/por_municipio.html",
        municipios=municipios_disponiveis,
        municipio_selecionado=municipio_selecionado,
        dados=dados_municipio,
        historico=historico,
        pagina_ativa="contratos",
        sub_ativa="por_municipio",
    )


@contratos_bp.route("/por-indicador")
def por_indicador():
    """
    Mostra TODOS os 6 indicadores ao mesmo tempo, uma linha por
    município (68 linhas) e uma coluna por indicador.

    Também monta, para cada indicador, a lista de municípios "dentro"
    e "fora" da meta (usada no modal que abre ao clicar no card de
    resumo — ver ponto 5 do pedido do Bruno).
    """
    comparacao = carregar_comparacao()
    evolucao = carregar_evolucao_metas()

    linhas = []
    for municipio in sorted(comparacao.keys()):
        indicadores_municipio = comparacao[municipio]
        celulas = {}
        for chave in INDICADORES:
            info = indicadores_municipio.get(chave, {})
            celulas[chave] = {
                "real": info.get("real"),
                "meta": info.get("meta"),
                "operador": info.get("operador"),
                "status": info.get("status"),
                "evolucao": evolucao.get((municipio, chave), []),
            }
        linhas.append({"municipio": municipio, "indicadores": celulas})

    # ── Cards de resumo + listas dentro/fora (para o modal) ──
    resumo_por_indicador = {}
    for chave in INDICADORES:
        dentro = []
        fora = []
        for l in linhas:
            item = l["indicadores"][chave]
            entrada = {"municipio": l["municipio"], "real": item["real"], "meta": item["meta"]}
            if item["status"] == "cumpre":
                dentro.append(entrada)
            elif item["status"] == "nao_cumpre":
                fora.append(entrada)
        resumo_por_indicador[chave] = {
            "cumpre": len(dentro),
            "nao_cumpre": len(fora),
            "dentro": dentro,
            "fora": fora,
        }

    return render_template(
        "contratos/por_indicador.html",
        indicadores=INDICADORES,
        linhas=linhas,
        resumo_por_indicador=resumo_por_indicador,
        pagina_ativa="contratos",
        sub_ativa="por_indicador",
    )


@contratos_bp.route("/analise-geral")
def analise_geral():
    """
    Visão geral dos 6 indicadores: um gráfico de rosca por indicador
    (quantos municípios atingiram/não atingiram a meta) + uma lista
    "cascata" expansível abaixo de cada um, com o nome de cada
    município separado por status.
    """
    comparacao = carregar_comparacao()

    resumo = {}
    for chave, info in INDICADORES.items():
        atingiram = []
        nao_atingiram = []
        for municipio, indicadores_municipio in comparacao.items():
            item = indicadores_municipio.get(chave, {})
            if item.get("status") == "cumpre":
                atingiram.append(municipio)
            elif item.get("status") == "nao_cumpre":
                nao_atingiram.append(municipio)
        atingiram.sort()
        nao_atingiram.sort()
        total = len(atingiram) + len(nao_atingiram)
        # percentual de municípios que atingiram a meta (0 se não houver dados,
        # para não quebrar a divisão quando total == 0)
        pct = round((len(atingiram) / total) * 100) if total else 0

        # nível de desempenho usado para colorir o card no template:
        # >=90% = bom (verde) | 70-89% = atenção (âmbar) | <70% = crítico (vermelho)
        if pct >= 90:
            nivel = "bom"
        elif pct >= 70:
            nivel = "atencao"
        else:
            nivel = "critico"

        resumo[chave] = {
            "label": info["label"],
            "atingiram": atingiram,
            "nao_atingiram": nao_atingiram,
            "total": total,
            "pct": pct,
            "nivel": nivel,
        }

    return render_template(
        "contratos/analise_geral.html",
        resumo=resumo,
        pagina_ativa="contratos",
        sub_ativa="analise_geral",
    )
