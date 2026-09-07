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
# na planilha exportada (ex.: "INF8001", "fórmula AG010-AG019/AG003").
CAMPOS_RAD = {
    "pop_urbana":                     {"label": "População Urbana",                  "unidade": "hab",        "classificacao": "auto",   "secao": "municipio",     "codigo": "INF8001",   "area_validacao": None},
    "cresc_vegetativo":               {"label": "Crescimento Vegetativo",             "unidade": "hab",        "classificacao": "auto",   "secao": "municipio",     "codigo": "INF8004 (dez/atual − dez/anterior)", "area_validacao": None},
    "regional":                       {"label": "Regional",                          "unidade": "",           "classificacao": "auto",   "secao": "municipio",     "codigo": "SANESUL_Regionais_v1", "area_validacao": None},
    "habitantes_domicilio":           {"label": "Habitantes por Domicílio",          "unidade": "hab/econ.",  "classificacao": "auto",   "secao": "municipio",     "codigo": "INF8005",   "area_validacao": None},
    "localidades_atendidas":          {"label": "Localidades Atendidas",             "unidade": "",           "classificacao": "manual", "secao": "municipio",     "codigo": "sem fonte identificada", "area_validacao": None},

    "pop_atendida_agua_hab":          {"label": "Pop. Atendida — Água",              "unidade": "hab",        "classificacao": "auto",   "secao": "agua",          "codigo": "INF8006 → AG026", "area_validacao": "agua"},
    "pop_atendida_agua_pct":          {"label": "Pop. Atendida — Água (%)",          "unidade": "%",          "classificacao": "auto",   "secao": "agua",          "codigo": "INF8006 → AG026", "area_validacao": "agua"},
    "captacao":                       {"label": "Captação",                          "unidade": "",           "classificacao": "manual", "secao": "agua",          "codigo": "nome da fonte", "area_validacao": "agua"},
    "eta":                            {"label": "ETA",                               "unidade": "",           "classificacao": "manual", "secao": "agua",          "codigo": "identificação", "area_validacao": "agua"},
    "pocos":                          {"label": "Poços",                             "unidade": "unid.",      "classificacao": "auto",   "secao": "agua",          "codigo": "INF26",     "area_validacao": "agua"},
    "extensao_rede_agua_km":          {"label": "Extensão de Rede — Água",           "unidade": "km",         "classificacao": "auto",   "secao": "agua",          "codigo": "INF33 → AG005", "area_validacao": "agua"},
    "reservacao_m3":                  {"label": "Reservação",                        "unidade": "m³",         "classificacao": "auto",   "secao": "agua",          "codigo": "INF54 + INF55", "area_validacao": "agua"},

    "pop_atendida_esgoto_hab":        {"label": "Pop. Atendida — Esgoto",            "unidade": "hab",        "classificacao": "auto",   "secao": "esgoto",        "codigo": "IND8301",   "area_validacao": "esgoto"},
    "pop_atendida_esgoto_pct":        {"label": "Pop. Atendida — Esgoto (%)",        "unidade": "%",          "classificacao": "auto",   "secao": "esgoto",        "codigo": "IND8301",   "area_validacao": "esgoto"},
    "tratamento_esgoto_pct":          {"label": "Tratamento de Esgoto",              "unidade": "%",          "classificacao": "manual", "secao": "esgoto",        "codigo": "valor fixo", "area_validacao": "esgoto"},
    # "ete" NÃO está mais aqui: município pode ter 1+ ETEs (ver aba
    # etes_municipio / listar_etes_municipio()). Como não é mais um
    # campo escalar, é tratado à parte em vez de por field_row genérico
    # — ver bloco dedicado em rad_exporters._escrever_ficha_municipio
    # e em templates/relatorios_agems/_ficha_macro.html.
    "extensao_rede_esgoto_km":        {"label": "Extensão de Rede — Esgoto",         "unidade": "km",         "classificacao": "auto",   "secao": "esgoto",        "codigo": "INF34 → ES004", "area_validacao": "esgoto"},

    "volume_produzido_m3ano":         {"label": "Volume Produzido",                  "unidade": "m³/ano",     "classificacao": "auto",   "secao": "operacional",   "codigo": "INF1 → AG006", "area_validacao": "agua"},
    "indice_perdas_distribuicao_pct": {"label": "Índice de Perdas na Distribuição",  "unidade": "%",          "classificacao": "auto",   "secao": "operacional",   "codigo": "fórmula INF1/67/9642/71/78", "area_validacao": "agua"},
    "indice_hidrometracao_pct":       {"label": "Índice de Hidrometração",           "unidade": "%",          "classificacao": "auto",   "secao": "operacional",   "codigo": "INF1003 / INF1065", "area_validacao": "agua"},
    "indice_macromedicao_pct":        {"label": "Índice de Macromedição",            "unidade": "%",          "classificacao": "auto",   "secao": "operacional",   "codigo": "IND8008",   "area_validacao": "agua"},
    "consumo_medio_economia":         {"label": "Consumo Médio por Economia",        "unidade": "m³/econ.",   "classificacao": "auto",   "secao": "operacional",   "codigo": "fórmula AG010–AG019/AG003", "area_validacao": "agua"},
    # Tarifas/ticket médio: indicadores operacionais de tarifa, não fazem
    # mais parte do Relatório de Validação da área Contábil (que agora
    # cobre apenas fonte de investimento + ligações, ver print 04) — por
    # isso area_validacao=None. Continuam aparecendo normalmente na ficha
    # (seção "Indicadores Operacionais"), pois isso é controlado por
    # "secao", não por "area_validacao".
    "tarifa_media_agua":              {"label": "Tarifa Média de Água",              "unidade": "R$/m³",      "classificacao": "auto",   "secao": "operacional",   "codigo": "FN002/AG011", "area_validacao": None},
    "tarifa_media_esgoto":            {"label": "Tarifa Média de Esgoto",            "unidade": "R$/m³",      "classificacao": "auto",   "secao": "operacional",   "codigo": "FN003/ES007", "area_validacao": None},
    "ticket_medio_total":             {"label": "Ticket Médio Total",                "unidade": "R$/economia","classificacao": "auto",   "secao": "operacional",   "codigo": "fórmula INF9652/3164/3170", "area_validacao": None},

    "cobertura_agua_atual_pct":       {"label": "Cobertura Água — Situação Atual",   "unidade": "%",          "classificacao": "auto",   "secao": "metas",         "codigo": "IND8300",   "area_validacao": "agua"},
    "perdas_lig_dia_atual":           {"label": "Perdas — Situação Atual",           "unidade": "l/lig/dia",  "classificacao": "auto",   "secao": "metas",         "codigo": "fórmula AG/INF", "area_validacao": "agua"},
    "iqa":                            {"label": "IQA",                               "unidade": "",           "classificacao": "manual", "secao": "metas",         "codigo": "planilha de controle própria", "area_validacao": "agua"},
    "cobertura_esgoto_atual_pct":     {"label": "Cobertura Esgoto — Situação Atual", "unidade": "%",          "classificacao": "auto",   "secao": "metas",         "codigo": "IND8301",   "area_validacao": "esgoto"},
    "tratamento_esgoto_atual_pct":    {"label": "Tratamento Esgoto — Situação Atual","unidade": "%",          "classificacao": "manual", "secao": "metas",         "codigo": "valor fixo", "area_validacao": "esgoto"},
    "eficiencia_tratamento":          {"label": "Eficiência no Tratamento",          "unidade": "%",          "classificacao": "manual", "secao": "metas",         "codigo": "Relatório SIBO (DBO)", "area_validacao": "esgoto"},

    # Ligações faturadas/não faturadas: vivem na SEÇÃO "operacional" da
    # ficha (ficam ao lado de tarifa_media_agua/esgoto e ticket_medio_
    # total — mesmo "assunto" de faturamento/comercial), mas continuam
    # pertencendo à ÁREA DE VALIDAÇÃO "investimento" (Alterações no RAD
    # 2025: "retirar dos investimentos [a EXIBIÇÃO na ficha] e realocar
    # em local condizente" — a validação continua no relatório único de
    # Investimentos, junto da tabela de fonte de água).
    "investimento_total_esgoto":      {"label": "Investimento Total — Esgoto",       "unidade": "R$",         "classificacao": "auto",   "secao": "investimentos", "codigo": "FN024/INF8389", "area_validacao": "investimento"},
    "outros_investimentos":           {"label": "Outros Investimentos",              "unidade": "R$",         "classificacao": "auto",   "secao": "investimentos", "codigo": "FN025/INF8390", "area_validacao": "investimento"},
    "obras_agua":                     {"label": "Principais Obras — Água",           "unidade": "",           "classificacao": "laud",   "secao": "investimentos", "codigo": "relato descritivo", "area_validacao": "investimento"},
    "melhorias_agua":                 {"label": "Melhorias Operacionais — Água",     "unidade": "",           "classificacao": "laud",   "secao": "investimentos", "codigo": "relato descritivo", "area_validacao": "investimento"},
    "obras_esgoto":                   {"label": "Principais Obras — Esgoto",         "unidade": "",           "classificacao": "laud",   "secao": "investimentos", "codigo": "relato descritivo", "area_validacao": "investimento"},
    "melhorias_esgoto":               {"label": "Melhorias Operacionais — Esgoto",   "unidade": "",           "classificacao": "laud",   "secao": "investimentos", "codigo": "relato descritivo", "area_validacao": "investimento"},

    # ── Investimento por fonte (própria/onerosa/não onerosa) — GECONT ──
    # "Contábil"/GECONT deixou de ser uma área de validação separada
    # (Alterações no RAD 2025, item 5) — todos esses campos agora são
    # area_validacao="investimento" e entram na banda ÁGUA do relatório
    # segregado de Investimentos (gerar_excel_validacao_area).
    "invest_fonte_propria_agua":      {"label": "Investimento Água — Fonte Própria",       "unidade": "R$", "classificacao": "manual", "secao": "investimentos", "codigo": "relato descritivo", "area_validacao": "investimento"},
    "invest_fonte_onerosa_agua":      {"label": "Investimento Água — Fonte Onerosa",       "unidade": "R$", "classificacao": "manual", "secao": "investimentos", "codigo": "relato descritivo", "area_validacao": "investimento"},
    "invest_fonte_nao_onerosa_agua":  {"label": "Investimento Água — Fonte Não Onerosa",   "unidade": "R$", "classificacao": "manual", "secao": "investimentos", "codigo": "relato descritivo", "area_validacao": "investimento"},
    "investimento_total_agua":        {"label": "Investimento Total — Água",         "unidade": "R$",         "classificacao": "auto",   "secao": "investimentos", "codigo": "FN023/INF8388", "area_validacao": "investimento"},
    "ligacoes_faturadas":             {"label": "Ligações Reais Faturadas",          "unidade": "unid.",      "classificacao": "auto",   "secao": "operacional",   "codigo": "INF9603",   "area_validacao": "investimento"},
    "ligacoes_nao_faturadas":         {"label": "Ligações Reais Não Faturadas",      "unidade": "unid.",      "classificacao": "auto",   "secao": "operacional",   "codigo": "INF3001",   "area_validacao": "investimento"},
    "invest_fonte_propria_esgoto":    {"label": "Investimento Esgoto — Fonte Própria",     "unidade": "R$", "classificacao": "manual", "secao": "investimentos", "codigo": "relato descritivo", "area_validacao": "investimento"},
    "invest_fonte_onerosa_esgoto":    {"label": "Investimento Esgoto — Fonte Onerosa",     "unidade": "R$", "classificacao": "manual", "secao": "investimentos", "codigo": "relato descritivo", "area_validacao": "investimento"},
    "invest_fonte_nao_onerosa_esgoto":{"label": "Investimento Esgoto — Fonte Não Onerosa", "unidade": "R$", "classificacao": "manual", "secao": "investimentos", "codigo": "relato descritivo", "area_validacao": "investimento"},

    # ── Ativos de uso compartilhado — GECONT (print 06) ─────────────
    "construcao_compartilhado":       {"label": "Ativos Compartilhados — Construção",              "unidade": "R$", "classificacao": "manual", "secao": "investimentos", "codigo": "relato descritivo", "area_validacao": "investimento"},
    "veiculos_compartilhado":         {"label": "Ativos Compartilhados — Aquisição de Veículos",   "unidade": "R$", "classificacao": "manual", "secao": "investimentos", "codigo": "relato descritivo", "area_validacao": "investimento"},
    "maquinarios_compartilhado":      {"label": "Ativos Compartilhados — Aquisição de Maquinários","unidade": "R$", "classificacao": "manual", "secao": "investimentos", "codigo": "relato descritivo", "area_validacao": "investimento"},
    "reforma_compartilhado":          {"label": "Ativos Compartilhados — Reforma/Conservação",     "unidade": "R$", "classificacao": "manual", "secao": "investimentos", "codigo": "relato descritivo", "area_validacao": "investimento"},
    "obras_andamento_compartilhado":  {"label": "Ativos Compartilhados — Obras em Andamento",      "unidade": "R$", "classificacao": "manual", "secao": "investimentos", "codigo": "relato descritivo", "area_validacao": "investimento"},

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
_cache = {"mtime": None, "df": None, "etes": None}


def _garantir_cache_atualizado() -> None:
    mtime_atual = ARQUIVO_RAD.stat().st_mtime
    if _cache["df"] is not None and _cache["mtime"] == mtime_atual:
        return
    df = pd.read_excel(ARQUIVO_RAD, sheet_name="dados_rad")
    df_etes = pd.read_excel(ARQUIVO_RAD, sheet_name="etes_municipio")

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

    _cache["df"] = df
    _cache["etes"] = etes_por_municipio
    _cache["mtime"] = mtime_atual


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
    nunca foi preenchida manualmente. Em vez de mostrar "a preencher"
    para os 63 municípios que já têm o valor em tratamento_esgoto_pct,
    replicamos o valor aqui. Continua "a preencher" só para os 5
    municípios sem ETE (Agua Clara, Mundo Novo, Sete Quedas, Sonora,
    Taquarussu), que é o único caso real de pendência."""
    if dados.get("tratamento_esgoto_atual_pct") is None:
        dados["tratamento_esgoto_atual_pct"] = dados.get("tratamento_esgoto_pct")
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
    return _aplicar_regra_tratamento_100(dados)


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
