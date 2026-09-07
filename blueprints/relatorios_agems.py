"""
blueprints/relatorios_agems.py — SISPLAN v2
================================================
Módulo Relatórios AGEMS: apresentação do RAD (Relatório Anual de
Desempenho) por município, com:

  - Filtro de Município (inclui a opção "Todos os Municípios") e
    filtro de Ano, lado a lado na tela
  - Exportar Excel / Exportar HTML / Relatório de Validação — cada um
    decide sozinho, pelo filtro de Município, se gera 1 arquivo
    (município específico) ou um .zip com 1 arquivo por município
    (quando "Todos os Municípios" está selecionado)

Segue o mesmo padrão de blueprints/contratos.py: sem login (esta
versão do SISPLAN não usa Flask-Login), render_template + url_for.
"""

import io

from flask import Blueprint, render_template, request, send_file, abort

from scripts.rad_loader import (
    carregar_rad_municipio,
    carregar_rad_todos_municipios,
    listar_municipios_rad,
    obter_regional_por_municipio,
    AREAS_VALIDACAO,
    ANO_PADRAO,
    ANOS_DISPONIVEIS,
    TODOS_MUNICIPIOS,
    LABEL_TODOS_MUNICIPIOS,
)
from scripts.rad_exporters import (
    gerar_excel_rad,
    gerar_zip_excel_rad_todos,
    gerar_html_rad,
    gerar_zip_html_rad_todos,
    gerar_excel_validacao_area,
    gerar_zip_validacao_area_todos,
)

relatorios_agems_bp = Blueprint(
    "relatorios_agems", __name__, url_prefix="/relatorios-agems"
)


def _resolver_ano() -> int:
    return request.args.get("ano", type=int) or ANO_PADRAO


def _municipio_eh_todos(municipio: str) -> bool:
    return not municipio or municipio == TODOS_MUNICIPIOS


# ═══════════════════════════════ TELA ═══════════════════════════════

@relatorios_agems_bp.route("/rad")
def rad():
    municipios = listar_municipios_rad()
    municipio_selecionado = request.args.get("municipio") or municipios[0]
    ano = _resolver_ano()

    exibir_todos = _municipio_eh_todos(municipio_selecionado)
    if exibir_todos:
        dados, regional = {}, ""
    else:
        dados = carregar_rad_municipio(municipio_selecionado) or {}
        regional = dados.get("regional") or obter_regional_por_municipio(municipio_selecionado)

    return render_template(
        "relatorios_agems/index.html",
        municipios=municipios,
        municipio_selecionado=municipio_selecionado,
        municipio=municipio_selecionado,
        exibir_todos=exibir_todos,
        todos_municipios_valor=TODOS_MUNICIPIOS,
        label_todos_municipios=LABEL_TODOS_MUNICIPIOS,
        ano=ano,
        anos_disponiveis=ANOS_DISPONIVEIS,
        regional=regional,
        dados=dados,
        areas_validacao=AREAS_VALIDACAO,
        pagina_ativa="relatorios_agems",
        sub_ativa="rad",
    )


# ══════════════════════ EXPORTAÇÃO — RAD COMPLETO ═══════════════════

@relatorios_agems_bp.route("/exportar/excel")
def exportar_excel():
    """Ficha de 1 município, OU — se o filtro de Município estiver em
    'Todos os Municípios' — um .zip com 1 .xlsx por município."""
    municipio = request.args.get("municipio")
    ano = _resolver_ano()

    if _municipio_eh_todos(municipio):
        todos = carregar_rad_todos_municipios()
        buffer = gerar_zip_excel_rad_todos(ano, todos)
        return send_file(
            buffer,
            as_attachment=True,
            download_name=f"RAD_Todos_Municipios_{ano}.zip",
            mimetype="application/zip",
        )

    dados = carregar_rad_municipio(municipio)
    if dados is None:
        abort(404, "Município não encontrado.")

    regional = dados.get("regional") or obter_regional_por_municipio(municipio)
    buffer = gerar_excel_rad(municipio, ano, regional, dados)

    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"RAD_{municipio.replace(' ', '_')}_{ano}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@relatorios_agems_bp.route("/exportar/html")
def exportar_html():
    """Relatório visual de 1 município como .html, OU — se o filtro de
    Município estiver em 'Todos os Municípios' — um .zip com 1 .html
    por município."""
    municipio = request.args.get("municipio")
    ano = _resolver_ano()

    if _municipio_eh_todos(municipio):
        todos = carregar_rad_todos_municipios()
        buffer = gerar_zip_html_rad_todos(ano, todos)
        return send_file(
            buffer,
            as_attachment=True,
            download_name=f"RAD_Todos_Municipios_{ano}.zip",
            mimetype="application/zip",
        )

    dados = carregar_rad_municipio(municipio)
    if dados is None:
        abort(404, "Município não encontrado.")

    regional = dados.get("regional") or obter_regional_por_municipio(municipio)
    html = gerar_html_rad(municipio, ano, regional, dados)

    return send_file(
        io.BytesIO(html.encode("utf-8")),
        as_attachment=True,
        download_name=f"RAD_{municipio.replace(' ', '_')}_{ano}.html",
        mimetype="text/html",
    )


# ═══════════════════ RELATÓRIO DE VALIDAÇÃO (Função 2) ══════════════

@relatorios_agems_bp.route("/exportar/validacao/<area>")
def exportar_validacao_area(area):
    """Excel de validação de 1 área (água/esgoto/contábil/investimentos)
    para 1 município, OU — se o filtro de Município estiver em 'Todos
    os Municípios' — um .zip com 1 ficha da área por município."""
    if area not in AREAS_VALIDACAO:
        abort(404, f"Área de validação desconhecida: {area}")

    municipio = request.args.get("municipio")
    ano = _resolver_ano()
    nome_area = AREAS_VALIDACAO[area]

    if _municipio_eh_todos(municipio):
        todos = carregar_rad_todos_municipios()
        buffer = gerar_zip_validacao_area_todos(area, ano, todos)
        return send_file(
            buffer,
            as_attachment=True,
            download_name=f"Validacao_{nome_area}_Todos_Municipios_{ano}.zip",
            mimetype="application/zip",
        )

    dados = carregar_rad_municipio(municipio)
    if dados is None:
        abort(404, "Município não encontrado.")

    buffer = gerar_excel_validacao_area(area, municipio, ano, dados)

    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"Validacao_{nome_area}_{municipio.replace(' ', '_')}_{ano}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
