"""
scripts/rad_exporters.py — SISPLAN v2
==========================================
Geração das saídas do módulo Relatórios AGEMS:

  RAD completo (1 município OU 'Todos os Municípios', decidido pelo
  filtro de Município na tela — ver TODOS_MUNICIPIOS em rad_loader.py):
    - gerar_excel_rad(municipio, ano, regional, dados) -> BytesIO
    - gerar_zip_excel_rad_todos(ano, lista_municipios_dados) -> BytesIO
      (1 arquivo .xlsx por município, todos dentro de um .zip)
    - gerar_html_rad(municipio, ano, regional, dados) -> str
    - gerar_zip_html_rad_todos(ano, lista_municipios_dados) -> BytesIO
      (1 arquivo .html por município, todos dentro de um .zip)

  Relatório de Validação (por área — água / esgoto / contábil /
  investimentos — para mandar cada Excel à área responsável validar).
  Mesma lógica de 1-município-vs-todos do filtro principal:
    - gerar_excel_validacao_area(area, municipio, ano, dados) -> BytesIO
      (ficha de 1 área, 1 município: Campo, Código/Fonte, Valor Atual)
    - gerar_zip_validacao_area_todos(area, ano, lista_municipios_dados) -> BytesIO
      (1 ficha de validação da área por município, .zip com todos)

Zero dependência externa: nada de wkhtmltopdf/reportlab. O HTML abre
em qualquer navegador e pode ser impresso em PDF por ali (Ctrl+P →
Salvar como PDF) se algum dia precisar do formato PDF de fato.
"""

from io import BytesIO
import zipfile

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from flask import render_template

NAVY = "16213E"
NAVY2 = "1C2D52"
GRAY = "5B6472"
ORANGE = "C1560E"  # mesma cor do "Esgotamento Sanitário" no card web (--orange) — usada para diferenciar a seção Esgoto de Água (navy) na ficha
AUTO_BG, AUTO_TXT = "E3F5EA", "2E8B57"
MANUAL_BG, MANUAL_TXT = "FBF1DB", "B8860B"
LAUD_BG, LAUD_TXT = "FDECE1", "C1440E"
EXT_BG, EXT_TXT = "ECEEF1", "6B7280"
LIGHT_BG = "F7F8FA"
LEGEND_BG = "FFFBEA"
LINE = "D9DCE1"

_HEADER_FONT = Font(name="Calibri", color="FFFFFF", bold=True, size=18)
_SUBHEADER_FONT = Font(name="Calibri", color="DBE6F5", bold=True, size=12)
_SECTION_FONT = Font(name="Calibri", color="FFFFFF", bold=True, size=11)
_SUBSECTION_FONT = Font(name="Calibri", color="FFFFFF", bold=True, size=10)
_LABEL_FONT = Font(name="Calibri", color="1B2434", size=10.5)
_VALUE_FONT = Font(name="Calibri", color="16213E", bold=True, size=10.5)
_VALUE_PENDING_FONT = Font(name="Calibri", color="A9B0BB", italic=True, size=9.5)
_TAG_AUTO_FONT = Font(name="Calibri", color=AUTO_TXT, bold=True, size=9.5)
_TAG_MANUAL_FONT = Font(name="Calibri", color=MANUAL_TXT, bold=True, size=9.5)
_TAG_LAUD_FONT = Font(name="Calibri", color=LAUD_TXT, bold=True, size=9.5)
_TAG_EXT_FONT = Font(name="Calibri", color=EXT_TXT, bold=True, size=9.5)
_LEGEND_FONT = Font(name="Calibri", color="6B5B1F", italic=True, size=9.5)
_FOOTER_FONT = Font(name="Calibri", color="6B7280", italic=True, size=8.5)
_HEADER_TABLE_FONT = Font(name="Calibri", color="FFFFFF", bold=True, size=10)

_TAG_STYLE = {
    "auto":   (AUTO_BG, _TAG_AUTO_FONT, "Automático"),
    "manual": (MANUAL_BG, _TAG_MANUAL_FONT, "Manual"),
    "laud":   (LAUD_BG, _TAG_LAUD_FONT, "Laudenise"),
    "ext":    (EXT_BG, _TAG_EXT_FONT, "Base externa"),
}

_THIN = Side(style="thin", color=LINE)
_BORDER_ALL = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_ROW_HEIGHT = 15  # 15pt = altura padrão do Excel p/ Calibri 10-11 (equivalente ao autoajuste de 1 linha)


def _altura_com_quebra(texto: str, largura_coluna: float, altura_minima: int = _ROW_HEIGHT) -> int:
    """Estima a altura de linha (em pontos) necessária para um texto com
    wrap_text=True não cortar visualmente. 1 unidade de largura de coluna
    do Excel ≈ 1 caractere visível (aproximação da própria definição de
    largura de coluna do Excel, em nº de dígitos '0' da fonte padrão) —
    testado de olho no LibreOffice/Excel com os nomes de ETE reais do
    município de Dourados (5 ETEs) antes de fechar o valor. Usamos 0.95
    em vez de 1.0 de propósito, para folgar por segurança e nunca
    subestimar a quantidade de linhas (melhor uma linha em branco a
    mais do que texto cortado de novo). ~15pt por linha de texto é a
    altura padrão de uma linha simples. Sempre retorna pelo menos
    `altura_minima`."""
    if not texto:
        return altura_minima
    caracteres_por_linha = max(1, int(largura_coluna * 0.95))
    n_linhas = -(-len(str(texto)) // caracteres_por_linha)  # ceil sem importar math
    return max(altura_minima, n_linhas * 15)

# Quantas colunas extras (D, E, F...) uma linha pode "tomar emprestado"
# da coluna C quando 1 obra isolada (já sem \n pra separar) ainda é um
# parágrafo grande demais para a largura padrão. 6 colunas extras (até
# a I) leva a largura de 36 para ~87 unidades — o suficiente pra cortar
# a altura por ~2-3x nos casos reais (mediana 148 e p90 355 caracteres
# dos 68 municípios) sem deixar a planilha visualmente torta nas outras
# linhas, que continuam A:C.
_COLUNAS_EXTRA_TEXTO_LONGO = 6
_LIMIAR_OBRA_ISOLADA_LONGA = 120  # obra sem \n com mais que isso já pede a coluna extra


def _field_row_obra(ws, row: int, campo_meta: dict, valor) -> int:
    """Como field_row (ver mais abaixo), mas para os 6 campos de texto
    livre da Laudenise (Obras/Melhorias) — que podem trazer VÁRIAS obras
    juntas no mesmo campo, separadas por \\n. Em vez de 1 célula gigante
    (o que gerava linhas de 20+ de altura — ver Alterações no RAD 2025,
    análise dos 68 municípios: mediana 148 caracteres, mas o pior caso —
    Dourados — passa de 1.200), quebra em 1 linha Excel por obra, com o
    rótulo repetido em cada uma — mesmo padrão já aprovado em
    _linhas_obra_validacao (Relatório de Validação de Investimentos).
    Quando uma obra isolada (já sem \\n) ainda é grande, a linha mescla
    C com mais _COLUNAS_EXTRA_TEXTO_LONGO colunas só ali, alargando a
    célula em vez de deixá-la só mais alta. Retorna quantas linhas do
    Excel foram usadas, para o chamador avançar `r` de acordo (nem toda
    chamada consome só 1 linha, diferente de field_row)."""
    from scripts.rad_loader import eh_explicacao, texto_explicacao

    label = campo_meta["label"]
    if eh_explicacao(valor):
        obras = [texto_explicacao(valor)]
    elif valor:
        obras = [linha.strip() for linha in str(valor).split("\n") if linha.strip()]
        if not obras:
            obras = ["a preencher"]
    else:
        obras = ["a preencher"]

    linha = row
    for obra in obras:
        lc = ws.cell(row=linha, column=2, value=label)
        lc.font = _LABEL_FONT
        lc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        eh_pendencia = obra in ("a preencher", "Não aplicável")
        vc = ws.cell(row=linha, column=3, value=obra)
        vc.font = _VALUE_PENDING_FONT if eh_pendencia else _VALUE_FONT

        col_final = 3  # só C, como field_row normal
        largura_disponivel = ws.column_dimensions["C"].width
        if len(obra) > _LIMIAR_OBRA_ISOLADA_LONGA:
            col_final = 3 + _COLUNAS_EXTRA_TEXTO_LONGO  # C:I
            ws.merge_cells(start_row=linha, start_column=3, end_row=linha, end_column=col_final)
            largura_disponivel = sum(
                ws.column_dimensions[get_column_letter(c)].width or 8.43
                for c in range(3, col_final + 1)
            )
        vc.alignment = Alignment(
            horizontal="left",
            vertical="top" if col_final > 3 else "center",
            wrap_text=True,
        )

        for col_idx in range(1, col_final + 1):
            cell = ws.cell(row=linha, column=col_idx)
            cell.border = _BORDER_ALL
            if col_idx in (1, 2):
                cell.fill = _fill(LIGHT_BG)

        ws.row_dimensions[linha].height = _altura_com_quebra(obra, largura_disponivel)
        linha += 1

    return linha - row


_CORES_AREA = {
    "agua": "2F6FB0",
    "esgoto": ORANGE,
    "contabil": "5B6472",
    "investimento": "3F6B47",
}


def _fill(cor):
    return PatternFill("solid", fgColor=cor)


def _fmt_valor_excel(valor):
    """Formata número no padrão que a ficha usa (ex.: 254.918 / 87,65 /
    14.221.927,87). Strings (Captação, ETE, Localidades) passam direto."""
    if valor is None or valor == "":
        return None
    if isinstance(valor, str):
        return valor
    try:
        numero = float(valor)
    except (TypeError, ValueError):
        return valor
    if numero == int(numero):
        return f"{int(numero):,}".replace(",", ".")
    texto = f"{numero:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return texto


def _texto_ete(nome: str, tratamento_pct, eficiencia_pct) -> str:
    """Nome da ETE + Eficiência de tratamento (DBO), quando disponível
    (aba etes_municipio, planilha de eficiência 2025 — Bruno reverteu
    em 18/09/2026 a decisão de 17/09 de omitir isso, agora que a
    maioria das ETEs tem o valor real, não mais "a preencher").
    tratamento_pct continua sem aparecer aqui: é sempre 100% pela
    regra de negócio da SANESUL, não agrega informação."""
    if eficiencia_pct is None:
        return nome
    return f"{nome} — Eficiência: {_fmt_valor_excel(eficiencia_pct)}%"


def _nome_aba_unico(nomes_usados: set, base: str) -> str:
    """Garante nome de aba <= 31 caracteres e sem repetição (Excel não
    aceita abas duplicadas nem nomes maiores que 31 caracteres)."""
    base = base[:31]
    nome = base
    sufixo = 2
    while nome in nomes_usados:
        corte = 31 - len(f" ({sufixo})")
        nome = f"{base[:corte]} ({sufixo})"
        sufixo += 1
    nomes_usados.add(nome)
    return nome


def _nome_arquivo_unico(nomes_usados: set, nome_desejado: str) -> str:
    """Evita colisão de nomes dentro do .zip (municípios com nome muito
    parecido após normalização de espaços, por exemplo)."""
    nome = nome_desejado
    sufixo = 2
    raiz, ponto, ext = nome_desejado.rpartition(".")
    while nome in nomes_usados:
        nome = f"{raiz} ({sufixo}).{ext}" if ponto else f"{nome_desejado} ({sufixo})"
        sufixo += 1
    nomes_usados.add(nome)
    return nome


# ═══════════════════════════ FICHA (RAD completo) ═══════════════════

def _escrever_ficha_municipio(ws, municipio: str, ano: int, dados: dict, linha_inicial: int = 1) -> int:
    """Escreve a ficha completa de 1 município na worksheet `ws`, a
    partir de `linha_inicial`. Retorna a próxima linha livre (para
    permitir múltiplas fichas na mesma aba, se algum dia precisar)."""
    from scripts.rad_loader import CAMPOS_RAD, SECOES, eh_explicacao, texto_explicacao

    # 3 colunas de verdade (A-C): A é decorativa, B é o rótulo, C é o
    # valor. As colunas D (respiro) e E ("Origem": Automático/Manual/...)
    # existiram entre 17/09 e hoje e foram removidas a pedido do Bruno —
    # ele não precisa mais dessa informação na ficha (o Glossário, se um
    # dia precisar, continua guardando classificacao/código de cada
    # campo em CAMPOS_RAD).
    widths = {"A": 3, "B": 58, "C": 36}  # calibrado p/ 0 labels e 0 valores
    # quebrarem linha à toa (17/09: análise dos 62 campos de CAMPOS_RAD
    # + os 68 municípios reais — maior label tem 53 chars, maior valor
    # "Não aplicável (sistema por poços)" tem 33 chars; B=58/C=36 cobre
    # os dois com folga, deixando quase toda linha numa altura só)
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    r = linha_inicial

    def merge_band(row, texto, cor, fonte, altura=22):
        ws.merge_cells(f"A{row}:C{row}")
        c = ws.cell(row=row, column=1, value=texto)
        c.font = fonte
        c.fill = _fill(cor)
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[row].height = altura

    def linha_ete_simples(row, ete: dict | None):
        """1 ETE só — mostra nome, tratamento% e eficiência% numa única
        linha (mesma aparência de um field_row comum). Quebra o texto
        (wrap_text) e ajusta a altura da linha dinamicamente — nome de
        ETE + tratamento + eficiência pode passar da largura da coluna C."""
        lc = ws.cell(row=row, column=2, value="ETE")
        lc.font = _LABEL_FONT
        if ete:
            texto_valor = _texto_ete(ete["nome"], ete["tratamento_pct"], ete["eficiencia_pct"])
        else:
            texto_valor = "Não aplicável" if dados.get("_sem_esgoto") else "a preencher"
        vc = ws.cell(row=row, column=3, value=texto_valor)
        vc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        vc.font = _VALUE_FONT if ete else _VALUE_PENDING_FONT
        for col in "ABC":
            cell = ws[f"{col}{row}"]
            cell.border = _BORDER_ALL
            if col != "C":
                cell.fill = _fill(LIGHT_BG)
        ws.row_dimensions[row].height = _altura_com_quebra(texto_valor, ws.column_dimensions["C"].width)

    def bloco_etes(row, etes: list[dict]) -> int:
        """Bloco de ETE(s) do município — 1 linha simples se só há 1 ETE;
        se houver 2+, um cabeçalho identificador seguido de 1 linha por
        ETE (numerada), cada uma com seu próprio tratamento% e
        eficiência% (tratamento é sempre 100% pela regra de negócio da
        SANESUL; eficiência ainda é '' até termos o dado por
        ETE). Cada linha quebra o texto e tem altura ajustada ao tamanho
        do nome da ETE (evita texto cortado, ver print de Dourados com
        5 ETEs). Retorna a próxima linha livre."""
        if len(etes) <= 1:
            linha_ete_simples(row, etes[0] if etes else None)
            return row + 1

        ws.merge_cells(f"A{row}:C{row}")
        titulo = ws.cell(row=row, column=1, value=f"ETEs DO MUNICÍPIO ({len(etes)})")
        titulo.font = _SUBSECTION_FONT
        titulo.fill = _fill(GRAY)
        titulo.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[row].height = 20
        row += 1

        for i, ete in enumerate(etes, start=1):
            lc = ws.cell(row=row, column=2, value=f"ETE {i} de {len(etes)}")
            lc.font = _LABEL_FONT
            texto_valor = _texto_ete(ete["nome"], ete["tratamento_pct"], ete["eficiencia_pct"])
            vc = ws.cell(row=row, column=3, value=texto_valor)
            vc.font = _VALUE_FONT
            vc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            for col in "ABC":
                cell = ws[f"{col}{row}"]
                cell.border = _BORDER_ALL
                if col != "C":
                    cell.fill = _fill(LIGHT_BG)
            ws.row_dimensions[row].height = _altura_com_quebra(texto_valor, ws.column_dimensions["C"].width)
            row += 1
        return row

    def field_row(row, campo_meta, valor):
        """Rótulo e valor quebram linha (wrap_text) e a altura da linha
        é calculada a partir do maior dos dois textos — EXCETO
        Localidades Atendidas, que fica numa linha só sem quebra (ver
        nota abaixo): célula grande demais incomodava o Bruno."""
        label = campo_meta["label"]
        if campo_meta.get("unidade") and "(" not in label:
            label += f" ({campo_meta['unidade']})"
        texto_label = label
        lc = ws.cell(row=row, column=2, value=texto_label)
        lc.font = _LABEL_FONT
        lc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        valor_fmt = _fmt_valor_excel(valor)
        eh_expl = eh_explicacao(valor)
        texto_valor = texto_explicacao(valor) if eh_expl else (valor_fmt if valor_fmt is not None else "")

        # Localidades Atendidas pode listar 5+ nomes de cidade/distrito
        # (ex.: Dourados) — em vez de quebrar linha (o que deixava a
        # célula alta e incomodava, Alterações no RAD 2025 item 2 de
        # 17/09), fica numa única linha sem wrap: o texto "vazia"
        # visualmente por cima da célula vazia à direita, sem forçar a
        # linha a crescer em altura (mesmo padrão adotado no Relatório
        # de Validação).
        eh_localidades = label.startswith("Localidades Atendidas")

        vc = ws.cell(row=row, column=3, value=texto_valor)
        if eh_localidades:
            vc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)
        else:
            vc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        vc.font = _VALUE_PENDING_FONT if (eh_expl or valor_fmt is None) else _VALUE_FONT

        for col in "ABC":
            cell = ws[f"{col}{row}"]
            cell.border = _BORDER_ALL
            if col != "C":
                cell.fill = _fill(LIGHT_BG)

        if eh_localidades:
            ws.row_dimensions[row].height = _ROW_HEIGHT
        else:
            altura = max(
                _altura_com_quebra(texto_label, ws.column_dimensions["B"].width),
                _altura_com_quebra(texto_valor, ws.column_dimensions["C"].width),
            )
            ws.row_dimensions[row].height = altura

    # ---- Cabeçalho ----
    merge_band(r, "RAD - RELATÓRIO ANUAL DE DESEMPENHO", NAVY, _HEADER_FONT, altura=30); r += 1
    merge_band(r, f"MUNICÍPIO: {municipio.upper()} — {ano}", NAVY2, _SUBHEADER_FONT, altura=22); r += 1
    r += 1

    campos_por_secao = {}
    for campo, meta in CAMPOS_RAD.items():
        campos_por_secao.setdefault(meta["secao"], []).append((campo, meta))

    # Numeração romana I-V — igual à tela/HTML do sistema (Alterações no
    # RAD 2025, 17/09): Sistemas de Água/Esgoto ficam lado a lado
    # conceitualmente e NÃO recebem número próprio (antes eram II e III
    # aqui no Excel, separados — o Bruno achou estranho ter números
    # diferentes entre o Excel e o sistema e pediu pra igualar).
    merge_band(r, "I — " + SECOES["municipio"].upper(), NAVY2, _SECTION_FONT); r += 1
    for campo, meta in campos_por_secao.get("municipio", []):
        field_row(r, meta, dados.get(campo))
        r += 1
    r += 1

    merge_band(r, SECOES["agua"].upper(), NAVY2, _SECTION_FONT); r += 1
    for campo, meta in campos_por_secao.get("agua", []):
        field_row(r, meta, dados.get(campo)); r += 1
    r += 1

    merge_band(r, SECOES["esgoto"].upper(), ORANGE, _SECTION_FONT); r += 1
    for campo, meta in campos_por_secao.get("esgoto", []):
        field_row(r, meta, dados.get(campo)); r += 1
    r += 1  # respiro entre "Sistemas de Esgoto" e o bloco de ETEs
    r = bloco_etes(r, dados.get("etes", []))
    r += 1

    # II — Indicadores Operacionais: dividido em Água / Esgoto / Geral
    # (Alterações no RAD 2025, item 8) — mesmo padrão de sub-bandas já
    # usado em Investimentos Realizados. Grupos fixos em
    # rad_loader.GRUPOS_OPERACIONAL (não dá pra usar area_validacao
    # aqui: tarifas/ticket são area_validacao=None e ligações são
    # area_validacao="investimento", mesmo sendo indicador operacional).
    from scripts.rad_loader import GRUPOS_OPERACIONAL
    campos_operacional = {c: m for c, m in campos_por_secao.get("operacional", [])}

    merge_band(r, "II — " + SECOES["operacional"].upper(), NAVY2, _SECTION_FONT); r += 1

    merge_band(r, "ÁGUA", "2F6FB0", _SUBSECTION_FONT, altura=20); r += 1
    for campo in GRUPOS_OPERACIONAL["agua"]:
        field_row(r, campos_operacional[campo], dados.get(campo)); r += 1
    r += 1

    merge_band(r, "ESGOTO", ORANGE, _SUBSECTION_FONT, altura=20); r += 1
    if GRUPOS_OPERACIONAL["esgoto"]:
        for campo in GRUPOS_OPERACIONAL["esgoto"]:
            field_row(r, campos_operacional[campo], dados.get(campo)); r += 1
    else:
        ws.cell(row=r, column=2, value="Sem indicador operacional específico cadastrado").font = _VALUE_PENDING_FONT
        for col in "ABC":
            cell = ws[f"{col}{r}"]
            cell.border = _BORDER_ALL
            if col != "C":
                cell.fill = _fill(LIGHT_BG)
        ws.row_dimensions[r].height = _ROW_HEIGHT
        r += 1
    r += 1

    merge_band(r, "GERAL (ÁGUA + ESGOTO)", GRAY, _SUBSECTION_FONT, altura=20); r += 1
    for campo in GRUPOS_OPERACIONAL["geral"]:
        field_row(r, campos_operacional[campo], dados.get(campo)); r += 1
    r += 1

    merge_band(r, "III — " + SECOES["metas"].upper(), GRAY, _SECTION_FONT); r += 1
    for campo, meta in campos_por_secao.get("metas", []):
        field_row(r, meta, dados.get(campo)); r += 1
    r += 1

    # IV — Investimentos Realizados: dividido em 3 sub-bandas (Água /
    # Esgoto / Compartilhado), cada campo já como seu próprio field_row
    # (label + valor) — em vez da lista única e sem classificação de
    # antes (Alterações no RAD 2025, item 1 do Excel completo).
    # Ligações reais faturadas/não faturadas NÃO aparecem mais aqui —
    # mudaram de seção para "operacional" (ver rad_loader.py), saem
    # junto com II — Indicadores Operacionais, ao lado das tarifas.
    merge_band(r, "IV — " + SECOES["investimentos"].upper(), NAVY2, _SECTION_FONT); r += 1

    campos_investimentos = {c: m for c, m in campos_por_secao.get("investimentos", [])}
    _GRUPOS_INVESTIMENTOS_FICHA = [
        ("ÁGUA", "2F6FB0", [
            "obras_cobertura_agua", "obras_producao_agua", "melhorias_agua",
            "invest_fonte_propria_agua", "invest_fonte_onerosa_agua",
            "invest_fonte_nao_onerosa_agua", "investimento_total_agua",
        ]),
        ("ESGOTO", ORANGE, [
            "obras_cobertura_esgoto", "obras_tratamento_esgoto", "melhorias_esgoto",
            "invest_fonte_propria_esgoto", "invest_fonte_onerosa_esgoto",
            "invest_fonte_nao_onerosa_esgoto", "investimento_total_esgoto",
        ]),
        ("ATIVOS DE USO COMPARTILHADO", GRAY, [
            "construcao_compartilhado", "veiculos_compartilhado",
            "maquinarios_compartilhado", "reforma_compartilhado",
            "obras_andamento_compartilhado", "outros_investimentos",
        ]),
    ]
    # Campos de texto livre da Laudenise (Obras/Melhorias) — únicos com
    # classificacao "laud" — passam por _field_row_obra em vez do
    # field_row genérico (ver nota na função: evita a célula gigante de
    # 1 linha só quando o texto é longo).
    _campos_texto_longo = {c for c, m in CAMPOS_RAD.items() if m["classificacao"] == "laud"}

    for titulo_grupo, cor_grupo, campos_grupo in _GRUPOS_INVESTIMENTOS_FICHA:
        merge_band(r, titulo_grupo, cor_grupo, _SUBSECTION_FONT, altura=20); r += 1
        for campo in campos_grupo:
            if campo in _campos_texto_longo:
                r += _field_row_obra(ws, r, campos_investimentos[campo], dados.get(campo))
            else:
                field_row(r, campos_investimentos[campo], dados.get(campo)); r += 1
        r += 1  # respiro entre sub-bandas

    merge_band(r, "V — " + SECOES["contratual"].upper(), GRAY, _SECTION_FONT); r += 1
    for campo, meta in campos_por_secao.get("contratual", []):
        field_row(r, meta, dados.get(campo)); r += 1
    r += 1

    ws.merge_cells(f"A{r}:C{r}")
    rodape = ws.cell(row=r, column=1, value=f"Relatório gerado pelo módulo Relatórios AGEMS do SISPLAN v2 · {municipio}/{ano}")
    rodape.font = _FOOTER_FONT
    rodape.alignment = Alignment(horizontal="center", wrap_text=False)
    ws.row_dimensions[r].height = 18
    r += 1

    ws.sheet_view.showGridLines = False
    return r

def gerar_excel_rad(municipio: str, ano: int, regional: str, dados: dict) -> BytesIO:
    """Ficha de 1 município, layout do mockup aprovado. Retorna BytesIO
    pronto para `send_file`."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"RAD_{municipio}"[:31]
    _escrever_ficha_municipio(ws, municipio, ano, dados)
    ws.freeze_panes = "A3"  # só as 2 primeiras linhas (título + município) ficam fixas

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def gerar_zip_excel_rad_todos(ano: int, lista_municipios_dados: list[tuple[str, dict]]) -> BytesIO:
    """Função 3 — todos os municípios: 1 arquivo .xlsx por município,
    todos dentro de um único .zip para download."""
    zip_buffer = BytesIO()
    nomes_usados = set()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for municipio, dados in lista_municipios_dados:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = f"RAD_{municipio}"[:31]
            _escrever_ficha_municipio(ws, municipio, ano, dados)
            ws.freeze_panes = "A3"  # só as 2 primeiras linhas (título + município) ficam fixas

            arquivo_buffer = BytesIO()
            wb.save(arquivo_buffer)

            nome_arquivo = _nome_arquivo_unico(nomes_usados, f"RAD_{municipio.replace(' ', '_')}_{ano}.xlsx")
            zf.writestr(nome_arquivo, arquivo_buffer.getvalue())

    zip_buffer.seek(0)
    return zip_buffer


# ═══════════════════════════════ HTML ═══════════════════════════════

def gerar_html_rad(municipio: str, ano: int, regional: str, dados: dict) -> str:
    """Documento .html standalone de 1 município — mesmo visual da tela."""
    return render_template(
        "relatorios_agems/rad_export.html",
        municipio=municipio, ano=ano, regional=regional, dados=dados,
    )


def gerar_zip_html_rad_todos(ano: int, lista_municipios_dados: list[tuple[str, dict]]) -> BytesIO:
    """Função 3 — todos os municípios: 1 arquivo .html por município,
    todos dentro de um único .zip para download."""
    from scripts.rad_loader import obter_regional_por_municipio

    zip_buffer = BytesIO()
    nomes_usados = set()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for municipio, dados in lista_municipios_dados:
            regional = dados.get("regional") or obter_regional_por_municipio(municipio)
            html = render_template(
                "relatorios_agems/rad_export.html",
                municipio=municipio, ano=ano, regional=regional, dados=dados,
            )
            nome_arquivo = _nome_arquivo_unico(nomes_usados, f"RAD_{municipio.replace(' ', '_')}_{ano}.html")
            zf.writestr(nome_arquivo, html.encode("utf-8"))

    zip_buffer.seek(0)
    return zip_buffer


# ═══════════════════════ RELATÓRIO DE VALIDAÇÃO ═════════════════════
# Função 2: 1 Excel por área (água/esgoto/contábil/investimentos), com
# o valor atual de cada campo — sem colunas de confirmação/observação.

def _banda_validacao(ws, row: int, texto: str, cor: str, altura: int = 22) -> None:
    """Faixa colorida (B:C) usada para segregar sub-blocos dentro do
    Relatório de Validação de Investimentos (Água / Esgoto / Ativos
    Compartilhados) — mesmo estilo visual das bandas da ficha completa
    (merge_band em _escrever_ficha_municipio), só que restrita a B:C
    porque a coluna A do relatório de validação é decorativa e, desde
    17/09, a coluna Código/Fonte foi retirada (layout ficou com só
    Campo | Valor Atual)."""
    ws.merge_cells(f"B{row}:C{row}")
    c = ws.cell(row=row, column=2, value=texto)
    c.font = _HEADER_TABLE_FONT
    c.fill = _fill(cor)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row].height = altura
    ws.cell(row=row, column=1).fill = _fill(cor)


def _linhas_obra_validacao(ws, row: int, label_categoria: str, valor) -> int:
    """Uma ou mais linhas de "Categoria | Descrição da Obra" — uma linha
    por obra descrita, separadas por quebra de linha (Alt+Enter) dentro
    do mesmo campo da base (ex.: obras_cobertura_agua). Usado só no
    Relatório de Validação de Investimentos (Alterações no RAD 2025,
    17/09 — pedido do Bruno: coluna "Descrição da Obra" separada da
    coluna de valores financeiros, e cada obra na sua própria linha,
    igual ao RAD antigo — em vez de 1 célula só com texto livre).
    Retorna a próxima linha livre."""
    from scripts.rad_loader import eh_explicacao, texto_explicacao
    if eh_explicacao(valor):
        obras = [texto_explicacao(valor)]
    elif valor:
        obras = [linha.strip() for linha in str(valor).split("\n") if linha.strip()]
        if not obras:
            obras = ["a preencher"]
    else:
        obras = ["a preencher"]

    for obra in obras:
        lc = ws.cell(row=row, column=2, value=label_categoria)
        lc.font = _LABEL_FONT
        lc.alignment = Alignment(horizontal="left", vertical="center")
        eh_pendencia = obra in ("a preencher", "Não aplicável")
        vc = ws.cell(row=row, column=3, value=obra)
        vc.font = _VALUE_PENDING_FONT if eh_pendencia else _VALUE_FONT

        # Mesma lógica de _field_row_obra (ficha completa): quando a
        # obra já separada por \n ainda é um parágrafo longo (calibrado
        # com os 68 municípios reais: até 433 caracteres numa obra só),
        # mescla C com mais colunas extra só nessa linha, alargando a
        # célula em vez de deixá-la só mais alta.
        col_final = 3
        largura_disponivel = ws.column_dimensions["C"].width
        if len(obra) > _LIMIAR_OBRA_ISOLADA_LONGA:
            col_final = 3 + _COLUNAS_EXTRA_TEXTO_LONGO
            ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=col_final)
            largura_disponivel = sum(
                ws.column_dimensions[get_column_letter(c)].width or 8.43
                for c in range(3, col_final + 1)
            )
        vc.alignment = Alignment(
            horizontal="left",
            vertical="top" if col_final > 3 else "center",
            wrap_text=True,
        )
        for col_idx in range(1, col_final + 1):
            cell = ws.cell(row=row, column=col_idx)
            cell.border = _BORDER_ALL
            if col_idx in (1, 2):
                cell.fill = _fill(LIGHT_BG)
        ws.row_dimensions[row].height = _altura_com_quebra(obra, largura_disponivel)
        row += 1
    return row


def _linha_campo_validacao(ws, row: int, meta: dict, valor) -> None:
    """1 linha (Campo | Valor Atual) do relatório de validação — mesmo
    estilo da linha genérica usada dentro do loop padrão de
    gerar_excel_validacao_area, só extraída para função para poder ser
    reaproveitada pelo layout segregado de Investimentos. A coluna
    Código/Fonte foi retirada em 17/09 (pedido do Bruno: informação
    técnica que não ajudava quem valida os números por e-mail — o
    Glossário já cobre isso para quem precisar)."""
    valor_fmt = _fmt_valor_excel(valor)
    from scripts.rad_loader import eh_explicacao, texto_explicacao
    eh_expl = eh_explicacao(valor)
    texto_valor = texto_explicacao(valor) if eh_expl else (valor_fmt if valor_fmt is not None else "")
    label = meta["label"]
    if meta.get("unidade") and "(" not in label:
        label += f" ({meta['unidade']})"
    ws.cell(row=row, column=2, value=label).font = _LABEL_FONT
    vc = ws.cell(row=row, column=3, value=texto_valor)
    vc.font = _VALUE_PENDING_FONT if (eh_expl or valor_fmt is None) else _VALUE_FONT
    vc.alignment = Alignment(horizontal="center", vertical="center")
    for col in "ABC":
        cell = ws[f"{col}{row}"]
        cell.border = _BORDER_ALL
        if col != "C":
            cell.fill = _fill(LIGHT_BG)
    ws.row_dimensions[row].height = _ROW_HEIGHT


# Grupos do Relatório de Validação de Investimentos — cada um vira uma
# banda colorida própria (Água = azul do nosso layout, Esgoto = laranja,
# Compartilhado = cinza neutro). "Contábil"/GECONT não existe mais como
# área separada (Alterações no RAD 2025, item 5) — a fonte de
# investimento em água + total entram direto na banda ÁGUA.
_GRUPOS_INVESTIMENTO = [
    ("ÁGUA", "2F6FB0", [
        "obras_cobertura_agua", "obras_producao_agua", "melhorias_agua",
        "invest_fonte_propria_agua", "invest_fonte_onerosa_agua",
        "invest_fonte_nao_onerosa_agua", "investimento_total_agua",
    ]),
    ("ESGOTO", "C1560E", [
        "obras_cobertura_esgoto", "obras_tratamento_esgoto", "melhorias_esgoto",
        "invest_fonte_propria_esgoto", "invest_fonte_onerosa_esgoto",
        "invest_fonte_nao_onerosa_esgoto", "investimento_total_esgoto",
    ]),
    ("ATIVOS DE USO COMPARTILHADO", "5B6472", [
        "construcao_compartilhado", "veiculos_compartilhado",
        "maquinarios_compartilhado", "reforma_compartilhado",
        "obras_andamento_compartilhado", "outros_investimentos",
    ]),
]

# Grupos do Relatório de Validação de Água e de Esgoto (Alterações no RAD
# 2025, item 1 da função "Exportar relatório de validação") — cada um
# segrega por assunto (Sistemas / Indicadores Operacionais / Situação
# Atual), em vez da lista única e misturada de antes. Cor de cada banda
# segue a mesma paleta da ficha completa (água=azul, esgoto=laranja,
# indicadores/metas=navy/cinza neutros).
_GRUPOS_AGUA = [
    ("SISTEMAS DE ÁGUA", "2F6FB0", [
        "pop_atendida_agua_hab", "pop_atendida_agua_pct", "captacao", "eta",
        "pocos", "extensao_rede_agua_km", "reservacao_m3",
    ]),
    ("INDICADORES OPERACIONAIS", "1C2D52", [
        "volume_produzido_m3ano", "indice_perdas_distribuicao_pct",
        "indice_hidrometracao_pct", "indice_macromedicao_pct",
        "consumo_medio_economia",
    ]),
    ("SITUAÇÃO ATUAL / METAS", "5B6472", [
        "cobertura_agua_atual_pct", "perdas_lig_dia_atual", "iqa",
    ]),
]
_GRUPOS_ESGOTO = [
    ("SISTEMAS DE ESGOTO", "C1560E", [
        "pop_atendida_esgoto_hab", "pop_atendida_esgoto_pct",
        "tratamento_esgoto_pct", "extensao_rede_esgoto_km",
    ]),
    ("SITUAÇÃO ATUAL / METAS", "5B6472", [
        "cobertura_esgoto_atual_pct", "tratamento_esgoto_atual_pct",
        "eficiencia_tratamento",
    ]),
]
_GRUPOS_POR_AREA = {
    "agua": _GRUPOS_AGUA,
    "esgoto": _GRUPOS_ESGOTO,
    "investimento": _GRUPOS_INVESTIMENTO,
}


def gerar_excel_validacao_area(area: str, municipio: str, ano: int, dados: dict) -> BytesIO:
    """Ficha de validação de 1 área, 1 município — cada campo em uma
    linha (rótulo, código/fonte, valor atual). Pensado para ser
    enviado por e-mail à área responsável conferir os números."""
    from scripts.rad_loader import AREAS_VALIDACAO, campos_por_area, CAMPOS_RAD, eh_explicacao, texto_explicacao

    nome_area = AREAS_VALIDACAO[area]
    cor_area = _CORES_AREA[area]
    campos = campos_por_area(area)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Validacao_{nome_area}_{municipio}"[:31]

    widths = {"A": 3, "B": 58, "C": 36}  # mesma calibração da ficha completa
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.sheet_view.showGridLines = False

    r = 1
    ws.merge_cells(f"A{r}:C{r}")
    c = ws.cell(row=r, column=1, value=f"RELATÓRIO DE VALIDAÇÃO — {nome_area.upper()}")
    c.font = _HEADER_FONT; c.fill = _fill(cor_area)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[r].height = 30
    r += 1

    ws.merge_cells(f"A{r}:C{r}")
    c = ws.cell(row=r, column=1, value=f"MUNICÍPIO: {municipio.upper()} — ANO {ano}")
    c.font = _SUBHEADER_FONT; c.fill = _fill(NAVY2)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[r].height = 22
    r += 2

    # Bloco "CONTEXTO DO MUNICÍPIO" (população/regional/localidades)
    # removido em 17/09 (pedido do Bruno, item 8) — Município e Ano já
    # aparecem nos cabeçalhos acima; o resto não ajudava a validação.
    linha_contexto_fim = r

    if area == "investimento":
        # Água/Esgoto: 2 sub-tabelas por setor (Alterações no RAD 2025,
        # 17/09 — pedido do Bruno): "Descrição da Obra" (texto, 1 linha
        # por obra) separada de "Valor Atual" (financeiro) — antes as
        # duas coisas ficavam misturadas na mesma coluna genérica.
        _CAMPOS_OBRAS = {
            "agua":   [("Cobertura", "obras_cobertura_agua"), ("Produção", "obras_producao_agua"), ("Melhorias Operacionais", "melhorias_agua")],
            "esgoto": [("Cobertura", "obras_cobertura_esgoto"), ("Tratamento", "obras_tratamento_esgoto"), ("Melhorias Operacionais", "melhorias_esgoto")],
        }
        _CAMPOS_FONTES = {
            "agua":   ["invest_fonte_propria_agua", "invest_fonte_onerosa_agua", "invest_fonte_nao_onerosa_agua", "investimento_total_agua"],
            "esgoto": ["invest_fonte_propria_esgoto", "invest_fonte_onerosa_esgoto", "invest_fonte_nao_onerosa_esgoto", "investimento_total_esgoto"],
        }
        for titulo_setor, cor_setor, chave_setor in [("ÁGUA", "2F6FB0", "agua"), ("ESGOTO", "C1560E", "esgoto")]:
            _banda_validacao(ws, r, titulo_setor, cor_setor)
            r += 1

            ws.cell(row=r, column=1).fill = _fill(LIGHT_BG)
            for i, h in enumerate(["Categoria", "Descrição da Obra"], start=2):
                c = ws.cell(row=r, column=i, value=h)
                c.font = Font(name="Calibri", color=GRAY, bold=True, size=9.5)
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.fill = _fill(LIGHT_BG)
            ws.row_dimensions[r].height = 18
            r += 1
            for label_categoria, campo in _CAMPOS_OBRAS[chave_setor]:
                r = _linhas_obra_validacao(ws, r, label_categoria, dados.get(campo))
            r += 1  # respiro entre obras e fontes

            ws.cell(row=r, column=1).fill = _fill(LIGHT_BG)
            for i, h in enumerate(["Campo", "Valor Atual"], start=2):
                c = ws.cell(row=r, column=i, value=h)
                c.font = Font(name="Calibri", color=GRAY, bold=True, size=9.5)
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.fill = _fill(LIGHT_BG)
            ws.row_dimensions[r].height = 18
            r += 1
            for campo in _CAMPOS_FONTES[chave_setor]:
                _linha_campo_validacao(ws, r, CAMPOS_RAD[campo], dados.get(campo))
                r += 1
            r += 1  # respiro entre setores

        # Ativos de Uso Compartilhado — só financeiro, segue igual.
        titulo_grupo, cor_grupo, campos_grupo = next(
            g for g in _GRUPOS_INVESTIMENTO if g[0] == "ATIVOS DE USO COMPARTILHADO"
        )
        _banda_validacao(ws, r, titulo_grupo, cor_grupo)
        r += 1
        ws.cell(row=r, column=1).fill = _fill(LIGHT_BG)
        for i, h in enumerate(["Campo", "Valor Atual"], start=2):
            c = ws.cell(row=r, column=i, value=h)
            c.font = Font(name="Calibri", color=GRAY, bold=True, size=9.5)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.fill = _fill(LIGHT_BG)
        ws.row_dimensions[r].height = 18
        r += 1
        for campo in campos_grupo:
            _linha_campo_validacao(ws, r, CAMPOS_RAD[campo], dados.get(campo))
            r += 1
        r += 1

        last_row = r - 2

    elif area in _GRUPOS_POR_AREA:
        # Layout segregado em bandas coloridas por bloco temático, em vez
        # de uma tabela única genérica (Alterações no RAD 2025: item 1 da
        # função "Exportar relatório de validação" pede isso para Água e
        # Esgoto; a segregação de Investimentos já existia). Sem Table()
        # do openpyxl aqui de propósito: uma Tabela do Excel exige 1 único
        # cabeçalho no topo do intervalo, incompatível com bandas coloridas
        # intercaladas.
        for titulo_grupo, cor_grupo, campos_grupo in _GRUPOS_POR_AREA[area]:
            _banda_validacao(ws, r, titulo_grupo, cor_grupo)
            r += 1

            ws.cell(row=r, column=1).fill = _fill(LIGHT_BG)
            for i, h in enumerate(["Campo", "Valor Atual"], start=2):
                c = ws.cell(row=r, column=i, value=h)
                c.font = Font(name="Calibri", color=GRAY, bold=True, size=9.5)
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.fill = _fill(LIGHT_BG)
            ws.row_dimensions[r].height = 18
            r += 1

            for campo in campos_grupo:
                _linha_campo_validacao(ws, r, CAMPOS_RAD[campo], dados.get(campo))
                r += 1
            r += 1  # respiro entre blocos

        # Área Esgoto: bloco de ETE(s) do município entra por último,
        # como sub-banda própria — mesma lógica de sempre (1 linha se só
        # há 1 ETE; senão, 1 linha por ETE, numeradas).
        if area == "esgoto":
            etes = dados.get("etes", [])
            _banda_validacao(ws, r, "ETE(S) DO MUNICÍPIO", GRAY)
            r += 1
            if len(etes) <= 1:
                ete = etes[0] if etes else None
                ws.cell(row=r, column=2, value="ETE").font = _LABEL_FONT
                if ete:
                    texto_valor = _texto_ete(ete["nome"], ete["tratamento_pct"], ete["eficiencia_pct"])
                else:
                    texto_valor = "Não aplicável" if dados.get("_sem_esgoto") else "a preencher"
                vc = ws.cell(row=r, column=3, value=texto_valor)
                vc.font = _VALUE_FONT if ete else _VALUE_PENDING_FONT
                vc.alignment = Alignment(horizontal="center", vertical="center")
                for col in "ABC":
                    cell = ws[f"{col}{r}"]
                    cell.border = _BORDER_ALL
                    if col != "C":
                        cell.fill = _fill(LIGHT_BG)
                ws.row_dimensions[r].height = _ROW_HEIGHT
                r += 1
            else:
                for i, ete in enumerate(etes, start=1):
                    ws.cell(row=r, column=2, value=f"ETE {i} de {len(etes)}").font = _LABEL_FONT
                    vc = ws.cell(row=r, column=3, value=_texto_ete(ete["nome"], ete["tratamento_pct"], ete["eficiencia_pct"]))
                    vc.font = _VALUE_FONT
                    vc.alignment = Alignment(horizontal="center", vertical="center")
                    for col in "ABC":
                        cell = ws[f"{col}{r}"]
                        cell.border = _BORDER_ALL
                        if col != "C":
                            cell.fill = _fill(LIGHT_BG)
                    ws.row_dimensions[r].height = _ROW_HEIGHT
                    r += 1
            r += 1  # respiro após o bloco de ETEs

        last_row = r - 2  # última linha de dado (descontando o respiro)
    else:
        # Cabeçalho da tabela de campos da área. Coluna A é só a faixa
        # decorativa (sempre em branco) — headers[0]="" tornava o Table()
        # do openpyxl inválido (toda coluna de uma Tabela do Excel exige um
        # nome de cabeçalho não vazio e único), e era isso que corrompia o
        # arquivo ("Reparos em ... Tabela de parte de /xl/tables/table1.xml").
        # A tabela de verdade cobre só B:C; A recebe cor/borda à parte.
        ws.cell(row=r, column=1).fill = _fill(cor_area)
        headers = ["Campo", "Valor Atual"]
        for i, h in enumerate(headers, start=2):
            c = ws.cell(row=r, column=i, value=h)
            c.font = _HEADER_TABLE_FONT
            c.fill = _fill(cor_area)
            c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[r].height = 22
        linha_cabecalho_tabela = r
        r += 1

        for campo, meta in campos:
            _linha_campo_validacao(ws, r, meta, dados.get(campo))
            r += 1

        # "ete" não é mais um campo escalar em CAMPOS_RAD (ver rad_loader.py)
        # — na área Esgoto, a validação da(s) ETE(s) entra à parte, como
        # linha(s) extra na mesma tabela (mesmas 3 colunas: rótulo/valor),
        # já com tratamento% e eficiência% individuais de cada ETE.
        if area == "esgoto":
            etes = dados.get("etes", [])
            if len(etes) <= 1:
                ete = etes[0] if etes else None
                ws.cell(row=r, column=2, value="ETE").font = _LABEL_FONT
                if ete:
                    texto_valor = _texto_ete(ete["nome"], ete["tratamento_pct"], ete["eficiencia_pct"])
                else:
                    texto_valor = "Não aplicável" if dados.get("_sem_esgoto") else "a preencher"
                vc = ws.cell(row=r, column=3, value=texto_valor)
                vc.font = _VALUE_FONT if ete else _VALUE_PENDING_FONT
                vc.alignment = Alignment(horizontal="center", vertical="center")
                for col in "ABC":
                    cell = ws[f"{col}{r}"]
                    cell.border = _BORDER_ALL
                    if col != "C":
                        cell.fill = _fill(LIGHT_BG)
                ws.row_dimensions[r].height = _ROW_HEIGHT
                r += 1
            else:
                for i, ete in enumerate(etes, start=1):
                    ws.cell(row=r, column=2, value=f"ETE {i} de {len(etes)}").font = _LABEL_FONT
                    vc = ws.cell(row=r, column=3, value=_texto_ete(ete["nome"], ete["tratamento_pct"], ete["eficiencia_pct"]))
                    vc.font = _VALUE_FONT
                    vc.alignment = Alignment(horizontal="center", vertical="center")
                    for col in "ABC":
                        cell = ws[f"{col}{r}"]
                        cell.border = _BORDER_ALL
                        if col != "C":
                            cell.fill = _fill(LIGHT_BG)
                    ws.row_dimensions[r].height = _ROW_HEIGHT
                    r += 1

        last_row = r - 1
        tabela = Table(displayName=f"Validacao{area.capitalize()}", ref=f"B{linha_cabecalho_tabela}:C{last_row}")
        tabela.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(tabela)

    r += 1
    ws.merge_cells(f"A{r}:C{r}")
    rodape = ws.cell(row=r, column=1, value=f"Relatório de Validação — {nome_area} · {municipio}/{ano} · Gerado pelo módulo Relatórios AGEMS do SISPLAN v2")
    rodape.font = _FOOTER_FONT
    rodape.alignment = Alignment(horizontal="center", wrap_text=False)

    # No layout segregado em bandas (Água/Esgoto/Investimentos) não há 1
    # único cabeçalho de tabela para travar — trava logo abaixo do
    # cabeçalho de contexto; no layout genérico antigo (hoje sem
    # nenhuma área o utilizando, mantido por segurança), trava logo
    # abaixo do cabeçalho da tabela, como antes.
    ws.freeze_panes = f"A{linha_contexto_fim}" if area in _GRUPOS_POR_AREA else f"A{linha_cabecalho_tabela + 1}"

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def gerar_zip_validacao_area_todos(area: str, ano: int, lista_municipios_dados: list[tuple[str, dict]]) -> BytesIO:
    """Quando o filtro de Município = 'Todos os Municípios': gera 1
    ficha de validação da área por município (mesmo formato de
    gerar_excel_validacao_area), todas dentro de um único .zip."""
    from scripts.rad_loader import AREAS_VALIDACAO

    nome_area = AREAS_VALIDACAO[area]
    zip_buffer = BytesIO()
    nomes_usados = set()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for municipio, dados in lista_municipios_dados:
            arquivo_buffer = gerar_excel_validacao_area(area, municipio, ano, dados)
            nome_arquivo = _nome_arquivo_unico(
                nomes_usados,
                f"Validacao_{nome_area}_{municipio.replace(' ', '_')}_{ano}.xlsx",
            )
            zf.writestr(nome_arquivo, arquivo_buffer.getvalue())

    zip_buffer.seek(0)
    return zip_buffer
