"""
scripts/rad_loader.py — SISPLAN v2
=======================================
Loader do módulo Relatórios AGEMS (RAD — Relatório Anual de Desempenho).

Lê data/base_dados_rad.xlsx, com DUAS abas:
  - "dados_rad": um município por linha, com todos os campos do RAD já
    calculados (automáticos) ou em branco (manuais/Laudenise, ainda não
    integrados a uma fonte de dados própria — ver CAMPOS_RAD abaixo).
  - "etes_municipio": 1 linha por ETE (município, nome, ordem,
    tratamento_pct, eficiencia_tratamento_pct). Existe porque 9 dos 68
    municípios têm mais de uma ETE (Dourados tem 5) — a antiga coluna
    única "ete" em dados_rad virou esta aba normalizada. tratamento_pct
    é sempre 100 quando a ETE existe (regra de negócio: a SANESUL trata
    100% do esgoto que coleta); eficiencia_tratamento_pct (eficiência de
    remoção de DBO) ainda não é conhecida por ETE — fica vazia até
    alguém preencher manualmente (ex.: relatório SIBO).

Segue o mesmo padrão de cache por mtime do scripts/base_dados_loader.py:
o Excel só é relido quando o arquivo muda no disco, não a cada request.

NOTA sobre o filtro de ano: a base atual é um snapshot único (não há
uma coluna de ano nos dados de origem). O parâmetro `ano` já circula
pelo sistema (rótulos, nomes de arquivo) desde as primeiras versões;
esta versão adiciona o SELETOR na tela. Quando houver extrações
anuais separadas, ARQUIVO_RAD pode virar `base_dados_rad_{ano}.xlsx`
e `_garantir_cache_atualizado` passa a receber o ano como parâmetro.
"""

import pandas as pd
from pathlib import Path
from datetime import datetime

DATA_DIR = Path(__file__).parent.parent / "data"
ARQUIVO_RAD = DATA_DIR / "base_dados_rad.xlsx"

ANO_ATUAL = datetime.now().year
ANO_PADRAO = 2025  # ano corrente de dados alimentados na base
ANOS_DISPONIVEIS = [2025, 2026]  # 2026 será alimentado futuramente; sem histórico anterior

# Valor especial usado no filtro de Município para significar "todos os
# 68 municípios" — usado pelas rotas de exportação para decidir entre
# gerar 1 arquivo (município específico) ou um .zip (todos).
TODOS_MUNICIPIOS = "__todos__"
LABEL_TODOS_MUNICIPIOS = "Todos os Municípios"

# ── Áreas do Relatório de Validação (Função 2) ──────────────────────
# Cada campo do RAD pertence a exatamente uma área de validação, para
# que o Excel de cada área saia limpo e sem mistura de assunto.
# Só 3 áreas (Alterações no RAD 2025, item 5): "Contábil"/GECONT foi
# fundida de volta dentro de "Investimentos" — os campos que antes
# tinham area_validacao="contabil" agora são "investimento" (ver
# abaixo), e o relatório segregado de Investimentos (gerar_excel_
# validacao_area) já mostra esses campos dentro da própria banda ÁGUA.
AREAS_VALIDACAO = {
    "agua":         "Água",
    "esgoto":       "Esgoto",
    "investimento": "Investimentos",
}

# ── Metadados de cada campo do RAD ──────────────────────────────────
# classificacao: 'auto' (código INF/IND do SINISA), 'manual' (valor
# manual/fixo/GECONT), 'laud' (informação descritiva da Laudenise) ou
# 'ext' (base externa — contratos).
# secao: agrupamento usado na tela e no export "ficha" (ordem do RAD).
# area_validacao: agrupamento usado no Relatório de Validação (Função
# 2) — água / esgoto / contábil / investimento. Só existe para campos
# que fazem sentido mandar para uma área validar; campos de contexto
# (Informações do Município) não têm area_validacao — eles aparecem
# como cabeçalho em todo relatório de validação, não como linha própria.
# codigo: texto curto do código/fonte, mostrado como "hint" na tela e
# na planilha exportada (ex.: "GTA1001", "IAG2013"). Atualizados em
# 18/09/2026 para os códigos do novo SINISA (RAD 2026), a partir do
# glossário de indicadores (Glossário_SINISA_Água_e_Esgoto.xlsx) e do
# relatório gerencial mensal enviado ao SINISA (RelGer): os campos que
# usavam os códigos antigos (INF/IND/AG/ES/FN, do sistema anterior)
# passaram para os equivalentes SINISA quando existia correspondência
# direta; os que não têm correspondência (Poços, Ticket Médio,
# Ligações Faturadas/Não Faturadas, Habitantes por Domicílio,
# Crescimento Vegetativo, Investimento Total Água, Outros
# Investimentos) continuam com a fonte antiga — não são informações
# que a SANESUL envia ao SINISA.
CAMPOS_RAD = {
    "pop_urbana":                     {"label": "População Urbana",                  "unidade": "hab",        "classificacao": "auto",   "secao": "municipio",     "codigo": "INF8001",   "area_validacao": None},
    "cresc_vegetativo":               {"label": "Crescimento Vegetativo",             "unidade": "hab",        "classificacao": "auto",   "secao": "municipio",     "codigo": "INF8004 (dez/atual − dez/anterior)", "area_validacao": None},
    "regional":                       {"label": "Regional",                          "unidade": "",           "classificacao": "auto",   "secao": "municipio",     "codigo": "Planilha Bruno SANESUL_Regionais_v1", "area_validacao": None},
    "habitantes_domicilio":           {"label": "Habitantes por Domicílio",          "unidade": "hab/economia",  "classificacao": "auto",   "secao": "municipio",     "codigo": "INF8005",   "area_validacao": None},
    "localidades_atendidas":          {"label": "Localidades Atendidas",             "unidade": "",           "classificacao": "manual", "secao": "municipio",     "codigo": "Planilha Bruno", "area_validacao": None},

    "pop_atendida_agua_hab":          {"label": "População Atendida — Água",              "unidade": "hab",        "classificacao": "auto",   "secao": "agua",          "codigo": "INF8006 → AG026 · SINISA: GTA0019", "area_validacao": "agua"},
    "pop_atendida_agua_pct":          {"label": "População Atendida — Água (%)",          "unidade": "%",          "classificacao": "auto",   "secao": "agua",          "codigo": "INF8006 → AG026 · SINISA: GTA0019/DFE0001 (IAG0001)", "area_validacao": "agua"},
    "captacao":                       {"label": "Captação",                          "unidade": "",           "classificacao": "manual", "secao": "agua",          "codigo": "Planilha Bruno SANESUL_Regionais_v1", "area_validacao": "agua"},
    "eta":                            {"label": "ETA",                               "unidade": "",           "classificacao": "manual", "secao": "agua",          "codigo": "Planilha Bruno SANESUL_Regionais_v1", "area_validacao": "agua"},
    "pocos":                          {"label": "Poços",                             "unidade": "unidades",      "classificacao": "auto",   "secao": "agua",          "codigo": "INF26",     "area_validacao": "agua"},
    "extensao_rede_agua_km":          {"label": "Extensão de Rede",           "unidade": "km",         "classificacao": "auto",   "secao": "agua",          "codigo": "INF33 → AG005 · SINISA: GTA1102", "area_validacao": "agua"},
    "reservacao_m3":                  {"label": "Reservação",                        "unidade": "m³",         "classificacao": "auto",   "secao": "agua",          "codigo": "INF54 + INF55 · SINISA: GTA1208", "area_validacao": "agua"},

    "pop_atendida_esgoto_hab":        {"label": "População Atendida — Esgoto",            "unidade": "hab",        "classificacao": "auto",   "secao": "esgoto",        "codigo": "IND8301 · SINISA: GTE0023", "area_validacao": "esgoto"},
    "pop_atendida_esgoto_pct":        {"label": "População Atendida — Esgoto (%)",        "unidade": "%",          "classificacao": "auto",   "secao": "esgoto",        "codigo": "IND8301 · SINISA: GTE0023/DFE0001 (IES0001)", "area_validacao": "esgoto"},
    "tratamento_esgoto_pct":          {"label": "Tratamento de Esgoto",              "unidade": "%",          "classificacao": "manual", "secao": "esgoto",        "codigo": "valor fixo", "area_validacao": "esgoto"},
    # "ete" NÃO está mais aqui: município pode ter 1+ ETEs (ver aba
    # etes_municipio / listar_etes_municipio()). Como não é mais um
    # campo escalar, é tratado à parte em vez de por field_row genérico
    # — ver bloco dedicado em rad_exporters._escrever_ficha_municipio
    # e em templates/relatorios_agems/_ficha_macro.html.
    "extensao_rede_esgoto_km":        {"label": "Extensão de Rede",         "unidade": "km",         "classificacao": "auto",   "secao": "esgoto",        "codigo": "INF34 → ES004 · SINISA: GTE1001", "area_validacao": "esgoto"},
    # Adicionados em 17/09 (Bruno pediu para verificar se "todo esgoto
    # coletado é tratado" — confirmado: os 2 indicadores do SINISA
    # (8008 e 8009) batem exatamente em todos os 68 municípios na
    # extração fornecida). volume_esgoto_tratado_m3ano é sempre igual a
    # volume_esgoto_coletado_m3ano hoje — mantidos como 2 campos
    # (em vez de 1) porque são 2 indicadores distintos do SINISA e podem
    # divergir no futuro (ex.: parte do esgoto coletado, mas ainda não
    # tratado, num município específico).
    "volume_esgoto_coletado_m3ano":   {"label": "Volume Coletado",          "unidade": "m³/ano",     "classificacao": "auto",   "secao": "operacional",   "codigo": "8008 — VOL.ESGOTO COLETADO · SINISA: GTE1002", "area_validacao": "esgoto"},
    "volume_esgoto_tratado_m3ano":    {"label": "Volume Tratado",           "unidade": "m³/ano",     "classificacao": "auto",   "secao": "operacional",   "codigo": "8009 — VOL.ESGOTO COLETADO E TRATADO · SINISA: GTE1014", "area_validacao": "esgoto"},

    "volume_produzido_m3ano":         {"label": "Volume Produzido",                  "unidade": "m³/ano",     "classificacao": "auto",   "secao": "operacional",   "codigo": "INF1 → AG006 · SINISA: GTA1001", "area_validacao": "agua"},
    "indice_perdas_distribuicao_pct": {"label": "Índice de Perdas na Distribuição",  "unidade": "%",          "classificacao": "auto",   "secao": "operacional",   "codigo": "fórmula INF1/67/9642/71/78 · SINISA: IAG2013", "area_validacao": "agua"},
    "indice_hidrometracao_pct":       {"label": "Índice de Hidrometração",           "unidade": "%",          "classificacao": "auto",   "secao": "operacional",   "codigo": "INF1003 / INF1065 · SINISA: IAG1003", "area_validacao": "agua"},
    "indice_macromedicao_pct":        {"label": "Índice de Macromedição",            "unidade": "%",          "classificacao": "auto",   "secao": "operacional",   "codigo": "IND8008 · SINISA: IAG2003", "area_validacao": "agua"},
    "consumo_medio_economia":         {"label": "Consumo Médio por Economia",        "unidade": "m³/economia",   "classificacao": "auto",   "secao": "operacional",   "codigo": "fórmula AG010–AG019/AG003 · SINISA: IAG2009", "area_validacao": "agua"},
    # Tarifas/ticket médio: indicadores operacionais de tarifa, não fazem
    # mais parte do Relatório de Validação da área Contábil (que agora
    # cobre apenas fonte de investimento + ligações, ver print 04) — por
    # isso area_validacao=None. Continuam aparecendo normalmente na ficha
    # (seção "Indicadores Operacionais"), pois isso é controlado por
    # "secao", não por "area_validacao".
    "tarifa_media_agua":              {"label": "Tarifa Média de Água",              "unidade": "R$/m³",      "classificacao": "auto",   "secao": "operacional",   "codigo": "FN002/AG011 · SINISA: IFA1001", "area_validacao": None},
    "tarifa_media_esgoto":            {"label": "Tarifa Média de Esgoto",            "unidade": "R$/m³",      "classificacao": "auto",   "secao": "operacional",   "codigo": "FN003/ES007 · SINISA: IFE1001", "area_validacao": None},
    "ticket_medio_total":             {"label": "Ticket Médio Total",                "unidade": "R$/economia","classificacao": "auto",   "secao": "operacional",   "codigo": "fórmula INF9652/3164/3170", "area_validacao": None},

    "cobertura_agua_atual_pct":       {"label": "Cobertura Água — Situação Atual",   "unidade": "%",          "classificacao": "auto",   "secao": "metas",         "codigo": "IND8300 · SINISA: SAN08300", "area_validacao": "agua"},
    "perdas_lig_dia_atual":           {"label": "Perdas — Situação Atual",           "unidade": "l/lig/dia",  "classificacao": "auto",   "secao": "metas",         "codigo": "fórmula AG/INF · SINISA: IAG2015", "area_validacao": "agua"},
    "iqa":                            {"label": "IQA",                               "unidade": "",           "classificacao": "manual", "secao": "metas",         "codigo": "planilha de controle própria Acompanhamento IQA", "area_validacao": "agua"},
    "cobertura_esgoto_atual_pct":     {"label": "Cobertura Esgoto — Situação Atual", "unidade": "%",          "classificacao": "auto",   "secao": "metas",         "codigo": "IND8301 · SINISA: SAN08301", "area_validacao": "esgoto"},
    "tratamento_esgoto_atual_pct":    {"label": "Tratamento Esgoto — Situação Atual","unidade": "%",          "classificacao": "manual", "secao": "metas",         "codigo": "valor fixo", "area_validacao": "esgoto"},
    "eficiencia_tratamento":          {"label": "Eficiência no Tratamento",          "unidade": "%",          "classificacao": "manual", "secao": "metas",         "codigo": "IES2102 — Relatório SIBO (DBO) até integrar ao SINISA", "area_validacao": "esgoto"},

    # Ligações faturadas/não faturadas: vivem na SEÇÃO "operacional" da
    # ficha (ficam ao lado de tarifa_media_agua/esgoto e ticket_medio_
    # total — mesmo "assunto" de faturamento/comercial), mas continuam
    # pertencendo à ÁREA DE VALIDAÇÃO "investimento" (Alterações no RAD
    # 2025: "retirar dos investimentos [a EXIBIÇÃO na ficha] e realocar
    # em local condizente" — a validação continua no relatório único de
    # Investimentos, junto da tabela de fonte de água).
    "investimento_total_esgoto":      {"label": "Investimento Total — Esgoto",       "unidade": "R$",         "classificacao": "auto",   "secao": "investimentos", "codigo": "FN024/INF8389", "area_validacao": "investimento"},
    "outros_investimentos":           {"label": "Outros Investimentos",              "unidade": "R$",         "classificacao": "auto",   "secao": "investimentos", "codigo": "FN025/INF8390", "area_validacao": "investimento"},
    # Itens 6/7 (17/09): "Principais Obras" era 1 campo só; agora é
    # dividido em Obras de Cobertura + Obras de Produção (água) / Obras
    # de Cobertura + Obras de Tratamento (esgoto), cada um podendo levar
    # uma descrição de texto livre (relato de obra) — igual ao RAD
    # antigo. Item 14: código/fonte de todos os 6 campos = "Laudenise".
    "obras_cobertura_agua":            {"label": "Principais Obras — Cobertura (Água)",     "unidade": "",           "classificacao": "laud",   "secao": "investimentos", "codigo": "Laudenise", "area_validacao": "investimento"},
    "obras_producao_agua":             {"label": "Principais Obras — Produção (Água)",      "unidade": "",           "classificacao": "laud",   "secao": "investimentos", "codigo": "Laudenise", "area_validacao": "investimento"},
    "melhorias_agua":                 {"label": "Melhorias Operacionais — Água",     "unidade": "",           "classificacao": "laud",   "secao": "investimentos", "codigo": "Laudenise", "area_validacao": "investimento"},
    "obras_cobertura_esgoto":          {"label": "Principais Obras — Cobertura (Esgoto)",   "unidade": "",           "classificacao": "laud",   "secao": "investimentos", "codigo": "Laudenise", "area_validacao": "investimento"},
    "obras_tratamento_esgoto":         {"label": "Principais Obras — Tratamento (Esgoto)",  "unidade": "",           "classificacao": "laud",   "secao": "investimentos", "codigo": "Laudenise", "area_validacao": "investimento"},
    "melhorias_esgoto":               {"label": "Melhorias Operacionais — Esgoto",   "unidade": "",           "classificacao": "laud",   "secao": "investimentos", "codigo": "Laudenise", "area_validacao": "investimento"},

    # ── Investimento por fonte (própria/onerosa/não onerosa) — GECONT ──
    # "Contábil"/GECONT deixou de ser uma área de validação separada
    # (Alterações no RAD 2025, item 5) — todos esses campos agora são
    # area_validacao="investimento" e entram na banda ÁGUA do relatório
    # segregado de Investimentos (gerar_excel_validacao_area).
    "invest_fonte_propria_agua":      {"label": "Investimento Água — Fonte Própria",       "unidade": "R$", "classificacao": "manual", "secao": "investimentos", "codigo": "GFI2021 — Relatório de Fontes de Recursos GECONT", "area_validacao": "investimento"},
    "invest_fonte_onerosa_agua":      {"label": "Investimento Água — Fonte Onerosa",       "unidade": "R$", "classificacao": "manual", "secao": "investimentos", "codigo": "GFI2022 — Relatório de Fontes de Recursos GECONT", "area_validacao": "investimento"},
    "invest_fonte_nao_onerosa_agua":  {"label": "Investimento Água — Fonte Não Onerosa",   "unidade": "R$", "classificacao": "manual", "secao": "investimentos", "codigo": "GFI2023 — Relatório de Fontes de Recursos GECONT", "area_validacao": "investimento"},
    "investimento_total_agua":        {"label": "Investimento Total — Água",         "unidade": "R$",         "classificacao": "auto",   "secao": "investimentos", "codigo": "FN023/INF8388", "area_validacao": "investimento"},
    "ligacoes_faturadas":             {"label": "Ligações Reais Faturadas",          "unidade": "unidades",      "classificacao": "auto",   "secao": "operacional",   "codigo": "INF9603",   "area_validacao": "investimento"},
    "ligacoes_nao_faturadas":         {"label": "Ligações Reais Não Faturadas",      "unidade": "unidades",      "classificacao": "auto",   "secao": "operacional",   "codigo": "INF3001",   "area_validacao": "investimento"},
    "invest_fonte_propria_esgoto":    {"label": "Investimento Esgoto — Fonte Própria",     "unidade": "R$", "classificacao": "manual", "secao": "investimentos", "codigo": "GFI2121 — Relatório de Fontes de Recursos GECONT", "area_validacao": "investimento"},
    "invest_fonte_onerosa_esgoto":    {"label": "Investimento Esgoto — Fonte Onerosa",     "unidade": "R$", "classificacao": "manual", "secao": "investimentos", "codigo": "GFI2122 — Relatório de Fontes de Recursos GECONT", "area_validacao": "investimento"},
    "invest_fonte_nao_onerosa_esgoto":{"label": "Investimento Esgoto — Fonte Não Onerosa", "unidade": "R$", "classificacao": "manual", "secao": "investimentos", "codigo": "GFI2123 — Relatório de Fontes de Recursos GECONT", "area_validacao": "investimento"},

    # ── Ativos de uso compartilhado — GECONT (print 06) ─────────────
    "construcao_compartilhado":       {"label": "Ativos Compartilhados — Construção",              "unidade": "R$", "classificacao": "manual", "secao": "investimentos", "codigo": "Relatório de Fontes de Recursos GECONT", "area_validacao": "investimento"},
    "veiculos_compartilhado":         {"label": "Ativos Compartilhados — Aquisição de Veículos",   "unidade": "R$", "classificacao": "manual", "secao": "investimentos", "codigo": "Relatório de Fontes de Recursos GECONT", "area_validacao": "investimento"},
    "maquinarios_compartilhado":      {"label": "Ativos Compartilhados — Aquisição de Maquinários","unidade": "R$", "classificacao": "manual", "secao": "investimentos", "codigo": "Relatório de Fontes de Recursos GECONT", "area_validacao": "investimento"},
    "reforma_compartilhado":          {"label": "Ativos Compartilhados — Reforma/Conservação",     "unidade": "R$", "classificacao": "manual", "secao": "investimentos", "codigo": "Relatório de Fontes de Recursos GECONT", "area_validacao": "investimento"},
    "obras_andamento_compartilhado":  {"label": "Ativos Compartilhados — Obras em Andamento",      "unidade": "R$", "classificacao": "manual", "secao": "investimentos", "codigo": "Relatório de Fontes de Recursos GECONT", "area_validacao": "investimento"},

    # Dados contratuais: saem do Relatório de Validação da área Contábil
    # (que agora segue só o print 04 — fonte de investimento em água +
    # ligações). Continuam aparecendo normalmente na ficha (seção "Dados
    # Contratuais"), pois isso é controlado por "secao", não por
    # "area_validacao".
    "num_convenio_agems":             {"label": "Convênio nº",                       "unidade": "",           "classificacao": "ext",    "secao": "contratual",    "codigo": "base_dados_sisplan / contratos_programa", "area_validacao": None},
    "data_assinatura_convenio":       {"label": "Convênio — Data de Assinatura",     "unidade": "",           "classificacao": "ext",    "secao": "contratual",    "codigo": "base_dados_sisplan / contratos_programa", "area_validacao": None},
    "data_expiracao_convenio":        {"label": "Convênio — Vigência (expiração)",   "unidade": "",           "classificacao": "ext",    "secao": "contratual",    "codigo": "base_dados_sisplan / contratos_programa", "area_validacao": None},
    "num_contrato_programa":          {"label": "Contrato de Programa nº",           "unidade": "",           "classificacao": "ext",    "secao": "contratual",    "codigo": "base_dados_sisplan / contratos_programa", "area_validacao": None},
    "data_assinatura_cp":             {"label": "Contrato de Programa — Data de Assinatura", "unidade": "",    "classificacao": "ext",    "secao": "contratual",    "codigo": "base_dados_sisplan / contratos_programa", "area_validacao": None},
    "duracao_anos_cp":                {"label": "Contrato de Programa — Duração",    "unidade": "anos",       "classificacao": "ext",    "secao": "contratual",    "codigo": "base_dados_sisplan / contratos_programa", "area_validacao": None},
    "data_expiracao_cp":              {"label": "Contrato de Programa — Vigência (expiração)", "unidade": "",  "classificacao": "ext",    "secao": "contratual",    "codigo": "base_dados_sisplan / contratos_programa", "area_validacao": None},
    "revisao_metas_recente":          {"label": "Revisão de Metas Contratuais — Data Mais Recente", "unidade": "", "classificacao": "ext", "secao": "contratual",    "codigo": "base_dados_sisplan / historico_aditivos", "area_validacao": None},
}

SECOES = {
    "municipio":     "Informações do Município",
    "agua":          "Sistemas de Água",
    "esgoto":        "Sistemas de Esgoto",
    "operacional":   "Indicadores Operacionais",
    "metas":         "Metas Contratuais",
    "investimentos": "Investimentos Realizados",
    "contratual":    "Dados Contratuais",
}

# Campos de contexto (Informações do Município) — usados como cabeçalho
# em todo relatório de validação, independentemente da área.
CAMPOS_CONTEXTO = [c for c, m in CAMPOS_RAD.items() if m["secao"] == "municipio"]

# ── Cache em memória — mesmo padrão de scripts/base_dados_loader.py ──
# "etes" guarda, por município (chave = nome exatamente como na coluna
# "municipio" de dados_rad), a lista de nomes de ETE em ordem — vazia
# quando o município não tem ETE própria registrada.
# "iqa" guarda, por município, o valor calculado a partir da planilha
# de controle própria (Acompanhamento_IQA.xlsx) — ver aba iqa_municipio
# em base_dados_rad.xlsx e a nota abaixo sobre a lógica do cálculo.
_cache = {"mtime": None, "df": None, "etes": None, "iqa": None}


def _garantir_cache_atualizado() -> None:
    mtime_atual = ARQUIVO_RAD.stat().st_mtime
    if _cache["df"] is not None and _cache["mtime"] == mtime_atual:
        return
    df = pd.read_excel(ARQUIVO_RAD, sheet_name="dados_rad")
    df_etes = pd.read_excel(ARQUIVO_RAD, sheet_name="etes_municipio")
    df_iqa = pd.read_excel(ARQUIVO_RAD, sheet_name="iqa_municipio")

    # Cada ETE vira um dict {nome, tratamento_pct, eficiencia_pct} — não só
    # o nome — porque desde a regra dos 100% de tratamento cada ETE tem seu
    # próprio tratamento_pct (sempre 100, pela regra de negócio da SANESUL)
    # e, quando disponível, sua própria eficiência de tratamento (DBO).
    etes_por_municipio: dict[str, list[dict]] = {}
    for municipio, grupo in df_etes.sort_values("ordem").groupby("municipio"):
        etes_por_municipio[municipio] = [
            {
                "nome": linha["ete_nome"],
                "tratamento_pct": None if pd.isna(linha["tratamento_pct"]) else float(linha["tratamento_pct"]),
                "eficiencia_pct": None if pd.isna(linha["eficiencia_tratamento_pct"]) else float(linha["eficiencia_tratamento_pct"]),
            }
            for _, linha in grupo.iterrows()
        ]

    # IQA do município = média das localidades atendidas (aba
    # iqa_municipio, já calculada — ver scripts/... nota de cálculo).
    # "incompleto"/"sem_dado" ficam disponíveis no dict de retorno para
    # uso futuro na tela (hoje não alteram nada além do valor do IQA em
    # si) — casos atuais: Itaporã, Porto Murtinho, Sidrolândia, Terenos.
    iqa_por_municipio: dict[str, dict] = {}
    for _, linha in df_iqa.iterrows():
        iqa_por_municipio[linha["municipio"]] = {
            "iqa": None if pd.isna(linha["iqa"]) else float(linha["iqa"]),
            "incompleto": str(linha["iqa_incompleto"]).strip().lower() == "sim",
            "sem_dado": None if pd.isna(linha["localidades_sem_dado"]) else linha["localidades_sem_dado"],
        }

    _cache["df"] = df
    _cache["etes"] = etes_por_municipio
    _cache["iqa"] = iqa_por_municipio
    _cache["mtime"] = mtime_atual


def _aplicar_iqa_calculado(dados: dict) -> dict:
    """Sobrescreve dados['iqa'] com o valor calculado (aba iqa_municipio),
    quando existir para o município. Também expõe 'iqa_incompleto' e
    'iqa_localidades_sem_dado' no dict — hoje não usados pela tela/Excel,
    mas disponíveis caso se queira sinalizar a pendência no futuro."""
    info = _cache["iqa"].get(dados["municipio"])
    if info is not None:
        dados["iqa"] = info["iqa"]
        dados["iqa_incompleto"] = info["incompleto"]
        dados["iqa_localidades_sem_dado"] = info["sem_dado"]
    return dados


def _aplicar_eficiencia_calculada(dados: dict) -> dict:
    """Preenche dados['eficiencia_tratamento'] (campo escalar, 1 por
    município, seção Metas Contratuais) a partir da(s) ETE(s) do
    município (aba etes_municipio), quando ainda vazio na base: 1 ETE
    só → usa o valor dela; 2+ ETEs → média simples entre as que já têm
    valor (decisão validada com Bruno em 18/09/2026). Sem ETE, ou
    nenhuma ETE com valor disponível, o campo permanece None e cai em
    'a preencher'/'Não aplicável' via _aplicar_explicacoes_pendencias,
    como já acontecia antes desta função existir."""
    if dados.get("eficiencia_tratamento") is not None:
        return dados
    valores = [e["eficiencia_pct"] for e in dados.get("etes", []) if e["eficiencia_pct"] is not None]
    if valores:
        dados["eficiencia_tratamento"] = round(sum(valores) / len(valores), 2)
    return dados


def listar_etes_municipio(municipio: str) -> list[dict]:
    """[{'nome', 'tratamento_pct', 'eficiencia_pct'}, ...] do município,
    na ordem cadastrada. Lista vazia quando não há ETE registrada."""
    _garantir_cache_atualizado()
    return list(_cache["etes"].get(municipio, []))


def listar_municipios_rad() -> list[str]:
    _garantir_cache_atualizado()
    return sorted(_cache["df"]["municipio"].tolist())


def _aplicar_regra_tratamento_100(dados: dict) -> dict:
    """Regra de negócio: a SANESUL trata 100% do esgoto que coleta (ver
    tratamento_pct em etes_municipio, sempre 100 quando a ETE existe).
    'tratamento_esgoto_atual_pct' (campo usado no card Metas Contratuais)
    é conceitualmente o MESMO valor de 'tratamento_esgoto_pct' (campo da
    seção Esgoto) — mas essa coluna vem sempre vazia da planilha, porque
    nunca foi preenchida manualmente. Em vez de mostrar uma pendência
    para os 63 municípios que já têm o valor em tratamento_esgoto_pct,
    replicamos o valor aqui."""
    if dados.get("tratamento_esgoto_atual_pct") is None:
        dados["tratamento_esgoto_atual_pct"] = dados.get("tratamento_esgoto_pct")
    return dados


# ── Explicações para campos vazios (17/09, análise com Bruno) ───────
# Toda a análise de base_dados_rad.xlsx (68 municípios) mostrou que
# "em branco" tem 3 causas bem diferentes, cada uma com o texto certo:
#
# 1) Município SEM sistema de esgoto (6: Agua Clara, Itaquirai, Mundo
#    Novo, Sete Quedas, Sonora, Taquarussu — pop_atendida_esgoto_pct
#    == 0) → "Não aplicável" nos campos de esgoto que dependem de ter
#    serviço ativo. Itaquirai É um caso especial: tem ETE cadastrada e
#    tratamento_esgoto_pct=100 mesmo sem população atendida ainda —
#    intencional (norma da empresa: quando o serviço entrar em
#    operação, o tratamento já nasce 100%), não é inconsistência.
# 2) Município com água ativa mas 100% captação por poços (53 de 68 —
#    confirmado 1:1 com pocos > 0) → "Não aplicável (sistema por
#    poços)" em Captação/ETA, porque esses municípios nunca tiveram (e
#    não precisam de) uma fonte de superfície nem uma ETA centralizada.
# 3) Todo o resto (Rio Verde sem regional, Aparecida Do Taboado sem
#    contrato, IQA/Eficiência, Obras/Melhorias — Laudenise,
#    Investimento por Fonte e Ativos Compartilhados — GECONT) é
#    pendência real → "a preencher".
#
# Implementação: em vez de deixar o valor None (que ficava em branco
# na tela — Alterações no RAD 2025, item 9 — revertido em 17/09 a
# pedido do Bruno: "informação em branco sem justificativa não pode"),
# o valor vira uma string marcada com EXPL_PREFIX. Os templates
# (fmt() macro) e o Excel (field_row) reconhecem o marcador e mostram
# o texto no estilo "pendência" (itálico cinza) em vez de valor real.
EXPL_PREFIX = "§EXPL§"


def eh_explicacao(v) -> bool:
    """True se `v` é um texto explicativo de pendência (não um valor
    real) — usado por rad_exporters.py para aplicar o estilo certo."""
    return isinstance(v, str) and v.startswith(EXPL_PREFIX)


def texto_explicacao(v) -> str:
    """Remove o marcador, devolvendo só o texto a exibir."""
    return v[len(EXPL_PREFIX):]


def _aplicar_explicacoes_pendencias(dados: dict) -> dict:
    """Preenche todo campo de CAMPOS_RAD ainda vazio com um texto
    explicativo marcado (ver nota acima) — nunca deixa None passar
    para a tela/Excel sem justificativa."""
    sem_esgoto = (dados.get("pop_atendida_esgoto_pct") or 0) == 0
    dados["_sem_esgoto"] = sem_esgoto  # usado pelo bloco de ETE (macro/index/exporter)

    if sem_esgoto:
        for campo in ("tratamento_esgoto_pct", "tratamento_esgoto_atual_pct",
                      "tarifa_media_esgoto", "extensao_rede_esgoto_km"):
            if dados.get(campo) is None:
                dados[campo] = EXPL_PREFIX + "Não aplicável"
        # Volume Coletado/Tratado vêm da base como 0 (não None) nesses
        # municípios — 0 aqui não é "zero m³ coletados", é "não há
        # serviço", então troca por "Não aplicável" mesmo sem ser None.
        for campo in ("volume_esgoto_coletado_m3ano", "volume_esgoto_tratado_m3ano"):
            if not dados.get(campo):
                dados[campo] = EXPL_PREFIX + "Não aplicável"

    capta_por_pocos = dados.get("captacao") is None and (dados.get("pocos") or 0) > 0
    if capta_por_pocos:
        if dados.get("captacao") is None:
            dados["captacao"] = EXPL_PREFIX + "Não aplicável (sistema por poços)"
        if dados.get("eta") is None:
            dados["eta"] = EXPL_PREFIX + "Não aplicável (sistema por poços)"

    for campo in CAMPOS_RAD:
        if dados.get(campo) is None:
            dados[campo] = EXPL_PREFIX + "a preencher"

    return dados


def carregar_rad_municipio(municipio: str) -> dict | None:
    """Retorna {campo: valor} para 1 município, ou None se não achar.
    Inclui 'etes' (lista de nomes) e 'qtd_etes' (int) — ver aba
    etes_municipio / listar_etes_municipio()."""
    _garantir_cache_atualizado()
    df = _cache["df"]
    linha = df[df["municipio"] == municipio]
    if linha.empty:
        return None
    dados = linha.iloc[0].to_dict()
    for k, v in dados.items():
        if pd.isna(v):
            dados[k] = None
    etes = listar_etes_municipio(municipio)
    dados["etes"] = etes
    dados["qtd_etes"] = len(etes)
    dados = _aplicar_regra_tratamento_100(dados)
    dados = _aplicar_iqa_calculado(dados)
    dados = _aplicar_eficiencia_calculada(dados)
    return _aplicar_explicacoes_pendencias(dados)


def obter_regional_por_municipio(municipio: str) -> str:
    """Fallback direto — a coluna 'regional' já vem em cada linha de
    carregar_rad_municipio(); esta função existe para os casos em que
    só se tem o nome do município (sem ter chamado a função acima)."""
    dados = carregar_rad_municipio(municipio)
    return (dados or {}).get("regional", "")


def carregar_rad_todos() -> pd.DataFrame:
    """Retorna o DataFrame completo (para exportação em massa)."""
    _garantir_cache_atualizado()
    return _cache["df"].copy()


def carregar_rad_todos_municipios() -> list[tuple[str, dict]]:
    """Retorna [(municipio, dados_dict), ...] para os 68 municípios, na
    ordem alfabética — usado pelas exportações 'Todos os Municípios'
    (Função 3) e pelo Relatório de Validação 'todos' (Função 2).
    Cada dados_dict inclui 'etes'/'qtd_etes', igual carregar_rad_municipio."""
    df = carregar_rad_todos().sort_values("municipio")
    resultado = []
    for _, row in df.iterrows():
        dados = row.to_dict()
        for k, v in dados.items():
            if pd.isna(v):
                dados[k] = None
        etes = listar_etes_municipio(dados["municipio"])
        dados["etes"] = etes
        dados["qtd_etes"] = len(etes)
        dados = _aplicar_regra_tratamento_100(dados)
        dados = _aplicar_iqa_calculado(dados)
        dados = _aplicar_eficiencia_calculada(dados)
        dados = _aplicar_explicacoes_pendencias(dados)
        resultado.append((dados["municipio"], dados))
    return resultado


def agrupar_por_secao(dados: dict) -> list[dict]:
    """Agrupa os campos do RAD por seção, na ordem de SECOES, cada um
    já com label/unidade/classificação/valor prontos para exibição —
    usado tanto pela tela (index) quanto pelo HTML exportado (mesma
    estrutura)."""
    secoes_com_campos = []
    for chave_secao, titulo_secao in SECOES.items():
        campos_da_secao = []
        for campo, meta in CAMPOS_RAD.items():
            if meta["secao"] != chave_secao:
                continue
            campos_da_secao.append({
                "campo": campo,
                "label": meta["label"],
                "unidade": meta["unidade"],
                "classificacao": meta["classificacao"],
                "valor": dados.get(campo),
            })
        secoes_com_campos.append({"titulo": titulo_secao, "campos": campos_da_secao})
    return secoes_com_campos


def campos_por_area(area: str) -> list[tuple[str, dict]]:
    """Retorna [(campo, meta), ...] pertencentes à área de validação
    informada ('agua', 'esgoto', 'contabil', 'investimento'), na ordem
    em que aparecem em CAMPOS_RAD."""
    if area not in AREAS_VALIDACAO:
        raise ValueError(f"Área de validação desconhecida: {area}")
    return [(c, m) for c, m in CAMPOS_RAD.items() if m.get("area_validacao") == area]


# ── Grupos água/esgoto/geral da seção "Indicadores Operacionais" ────
# (Alterações no RAD 2025, item 8) Só essa seção mistura água e esgoto
# sem uma chave própria de agrupamento (area_validacao não serve aqui:
# tarifa_media_agua/esgoto e ticket_medio_total são area_validacao=None,
# e ligações faturadas/não faturadas são area_validacao="investimento"
# mesmo sendo indicador operacional comercial). Mesmo padrão de lista
# fixa já usado em rad_exporters._GRUPOS_INVESTIMENTOS_FICHA.
GRUPOS_OPERACIONAL = {
    "agua": [
        "volume_produzido_m3ano", "indice_perdas_distribuicao_pct",
        "indice_hidrometracao_pct", "indice_macromedicao_pct",
        "consumo_medio_economia", "tarifa_media_agua",
    ],
    # RESOLVIDO (17/09): Volume Coletado e Volume Tratado (esgoto) foram
    # adicionados a base_dados_rad.xlsx a partir da extração SINISA
    # 8008/8009 (ver CAMPOS_RAD). Ticket Médio por Economia — Esgoto
    # continua sem existir na base (só existe ticket_medio_total,
    # sem separação água/esgoto) — segue em "geral".
    "esgoto": ["volume_esgoto_coletado_m3ano", "volume_esgoto_tratado_m3ano", "tarifa_media_esgoto"],
    "geral": [
        "ticket_medio_total", "ligacoes_faturadas", "ligacoes_nao_faturadas",
    ],
}


# ── Glossário (Alterações no RAD 2025, itens 1/2) ────────────────────
# Fonte única: CAMPOS_RAD. Não duplica nada — lista todos os ~68 campos
# do RAD com código/fonte, classificação e seção, na mesma ordem de
# SECOES. Usado pela nova rota /relatorios-agems/glossario.
_LABEL_CLASSIFICACAO = {
    "auto": "Automático (código SINISA: GTA/GTE/IAG/IES/IFA/IFE/SAN0/GFI)",
    "manual": "Manual / valor fixo / GECONT",
    "laud": "Laudenise",
    "ext": "Base externa (contratos)",
}


def listar_glossario() -> list[dict]:
    """[{secao, campo, label, codigo, classificacao_label}, ...] de
    TODOS os campos de CAMPOS_RAD, agrupados por seção na ordem de
    SECOES — inclui campos sem código de origem (ex.: Captação)."""
    itens = []
    for chave_secao, titulo_secao in SECOES.items():
        for campo, meta in CAMPOS_RAD.items():
            if meta["secao"] != chave_secao:
                continue
            itens.append({
                "secao": titulo_secao,
                "campo": campo,
                "label": meta["label"],
                "codigo": meta["codigo"],
                "classificacao": meta["classificacao"],  # 'auto'/'manual'/'laud'/'ext' — usado só para colorir a tag
                "classificacao_label": _LABEL_CLASSIFICACAO[meta["classificacao"]],
            })
    return itens
