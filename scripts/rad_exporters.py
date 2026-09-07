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
_ROW_HEIGHT = 20  # altura fixa (sem wrap) — todas as linhas de campo iguais


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

_CORES_AREA = {
    "agua": "2F6FB0",
    "esgoto": "16305C",
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
    from scripts.rad_loader import CAMPOS_RAD, SECOES

    # 4 colunas de verdade agora (A-D): a antiga coluna D era um
    # espaçador decorativo que nunca recebia conteúdo — removida. Tag
    # (Automático/Manual/...) passa a ser a coluna D.
    widths = {"A": 3, "B": 40, "C": 22, "D": 16}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    r = linha_inicial

    def merge_band(row, texto, cor, fonte, altura=22):
        ws.merge_cells(f"A{row}:D{row}")
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
        lc = ws.cell(row=row, column=2, value="ETE — nome / tratamento / eficiência")
        lc.font = _LABEL_FONT
        if ete:
            trat = f"{_fmt_valor_excel(ete['tratamento_pct'])}%" if ete["tratamento_pct"] is not None else "trat. a preencher"
            efic = f"{_fmt_valor_excel(ete['eficiencia_pct'])}%" if ete["eficiencia_pct"] is not None else "efic. a preencher"
            texto_valor = f"{ete['nome']} — {trat} — {efic}"
        else:
            texto_valor = "a preencher"
        vc = ws.cell(row=row, column=3, value=texto_valor)
        vc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        vc.font = _VALUE_FONT if ete else _VALUE_PENDING_FONT
        _, tag_font, tag_label = _TAG_STYLE["manual"]
        tc = ws.cell(row=row, column=4, value=tag_label)
        tc.font = tag_font
        tc.alignment = Alignment(horizontal="center", vertical="center")
        for col in "ABCD":
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
        SANESUL; eficiência ainda é 'a preencher' até termos o dado por
        ETE). Cada linha quebra o texto e tem altura ajustada ao tamanho
        do nome da ETE (evita texto cortado, ver print de Dourados com
        5 ETEs). Retorna a próxima linha livre."""
        if len(etes) <= 1:
            linha_ete_simples(row, etes[0] if etes else None)
            return row + 1

        ws.merge_cells(f"A{row}:D{row}")
        titulo = ws.cell(row=row, column=1, value=f"ETEs DO MUNICÍPIO ({len(etes)}) — tratamento e eficiência individuais")
        titulo.font = _SUBSECTION_FONT
        titulo.fill = _fill(GRAY)
        titulo.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[row].height = 20
        row += 1

        for i, ete in enumerate(etes, start=1):
            lc = ws.cell(row=row, column=2, value=f"ETE {i} de {len(etes)}")
            lc.font = _LABEL_FONT
            trat = f"{_fmt_valor_excel(ete['tratamento_pct'])}%" if ete["tratamento_pct"] is not None else "trat. a preencher"
            efic = f"{_fmt_valor_excel(ete['eficiencia_pct'])}%" if ete["eficiencia_pct"] is not None else "efic. a preencher"
            texto_valor = f"{ete['nome']} — {trat} — {efic}"
            vc = ws.cell(row=row, column=3, value=texto_valor)
            vc.font = _VALUE_FONT
            vc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            _, tag_font, tag_label = _TAG_STYLE["manual"]
            tc = ws.cell(row=row, column=4, value=tag_label)
            tc.font = tag_font
            tc.alignment = Alignment(horizontal="center", vertical="center")
            for col in "ABCD":
                cell = ws[f"{col}{row}"]
                cell.border = _BORDER_ALL
                if col != "C":
                    cell.fill = _fill(LIGHT_BG)
            ws.row_dimensions[row].height = _altura_com_quebra(texto_valor, ws.column_dimensions["C"].width)
            row += 1
        return row

    def field_row(row, campo_meta, valor):
        """Rótulo E valor sempre quebram linha (wrap_text) e a altura da
        linha é calculada a partir do maior dos dois textos — antes só
        o valor podia quebrar, e olhe lá; rótulo longo (ex.: "Índice de
        Perdas na Distribuição (%) — fórmula INF1/67/9642/71/78", 65
        caracteres) simplesmente cortava visualmente contra a borda da
        coluna de valor ao lado, porque o rótulo nunca quebrava."""
        label = campo_meta["label"]
        if campo_meta.get("unidade") and "(" not in label:
            label += f" ({campo_meta['unidade']})"
        texto_label = label
        lc = ws.cell(row=row, column=2, value=texto_label)
        lc.font = _LABEL_FONT
        lc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        valor_fmt = _fmt_valor_excel(valor)
        texto_valor = valor_fmt if valor_fmt is not None else "a preencher"

        # Localidades Atendidas pode listar 5+ nomes de cidade/distrito
        # (ex.: Dourados) — quebrando só na coluna C (22 de largura) a
        # linha ficava alta demais (Alterações no RAD 2025, item 5 do
        # Excel completo / item 2 do relatório de validação). Mescla
        # C:D só para esse campo (usa a largura de C+D juntas) — texto
        # cabe em bem menos linhas, ficando bem mais baixo.
        eh_localidades = label.startswith("Localidades Atendidas")
        if eh_localidades:
            ws.merge_cells(f"C{row}:D{row}")
            largura_valor = ws.column_dimensions["C"].width + ws.column_dimensions["D"].width
        else:
            largura_valor = ws.column_dimensions["C"].width

        vc = ws.cell(row=row, column=3, value=texto_valor)
        vc.alignment = Alignment(horizontal="left" if eh_localidades else "center", vertical="center", wrap_text=True)
        vc.font = _VALUE_FONT if valor_fmt is not None else _VALUE_PENDING_FONT

        if eh_localidades:
            # Sem coluna D disponível para a tag (mesclada com C) —
            # a tag "Manual" já está implícita: única linha do RAD que
            # não tem código INF/IND, fica óbvio ao olhar as vizinhas.
            ws.cell(row=row, column=4).fill = _fill(LIGHT_BG)
        else:
            _, tag_font, tag_label = _TAG_STYLE[campo_meta["classificacao"]]
            tc = ws.cell(row=row, column=4, value=tag_label)
            tc.font = tag_font
            tc.alignment = Alignment(horizontal="center", vertical="center")

        for col in "ABCD":
            cell = ws[f"{col}{row}"]
            cell.border = _BORDER_ALL
            if col != "C":
                cell.fill = _fill(LIGHT_BG)

        altura = max(
            _altura_com_quebra(texto_label, ws.column_dimensions["B"].width),
            _altura_com_quebra(texto_valor, largura_valor),
        )
        ws.row_dimensions[row].height = altura

    # ---- Cabeçalho ----
    merge_band(r, "RAD - RELATÓRIO ANUAL DE DESEMPENHO", NAVY, _HEADER_FONT, altura=30); r += 1
    merge_band(r, f"MUNICÍPIO: {municipio.upper()} — {ano}", NAVY2, _SUBHEADER_FONT, altura=22); r += 1
    r += 1

    ws.cell(row=r, column=2, value="Município").font = _LABEL_FONT
    ws.cell(row=r, column=2).fill = _fill(LIGHT_BG)
    ws.merge_cells(f"C{r}:D{r}")
    mc = ws.cell(row=r, column=3, value=municipio)
    mc.font = _VALUE_FONT
    mc.alignment = Alignment(horizontal="center")
    for col in "ABCD":
        ws[f"{col}{r}"].border = _BORDER_ALL
    ws.row_dimensions[r].height = 20
    r += 2

    campos_por_secao = {}
    for campo, meta in CAMPOS_RAD.items():
        campos_por_secao.setdefault(meta["secao"], []).append((campo, meta))

    merge_band(r, "I — " + SECOES["municipio"].upper(), NAVY2, _SECTION_FONT); r += 1
    for campo, meta in campos_por_secao.get("municipio", []):
        # field_row já calcula a altura certa sozinho (rótulo E valor
        # quebram e a linha usa o maior dos dois) — Localidades Atendidas
        # (que pode listar 5+ localidades, ex. Dourados) não precisa mais
        # de tratamento especial aqui.
        field_row(r, meta, dados.get(campo))
        r += 1
    r += 1

    merge_band(r, "II — " + SECOES["agua"].upper(), NAVY2, _SECTION_FONT); r += 1
    for campo, meta in campos_por_secao.get("agua", []):
        field_row(r, meta, dados.get(campo)); r += 1
    r += 1

    merge_band(r, "III — " + SECOES["esgoto"].upper(), ORANGE, _SECTION_FONT); r += 1
    for campo, meta in campos_por_secao.get("esgoto", []):
        field_row(r, meta, dados.get(campo)); r += 1
    r += 1  # respiro entre "III — Sistemas de Esgoto" e o bloco de ETEs (Alterações no RAD 2025, item 2)
    r = bloco_etes(r, dados.get("etes", []))
    r += 1

    merge_band(r, "IV — " + SECOES["operacional"].upper(), NAVY2, _SECTION_FONT); r += 1
    for campo, meta in campos_por_secao.get("operacional", []):
        field_row(r, meta, dados.get(campo)); r += 1
    r += 1

    merge_band(r, "V — " + SECOES["metas"].upper(), GRAY, _SECTION_FONT); r += 1
    for campo, meta in campos_por_secao.get("metas", []):
        field_row(r, meta, dados.get(campo)); r += 1
    r += 1

    # VI — Investimentos Realizados: dividido em 3 sub-bandas (Água /
    # Esgoto / Compartilhado), cada campo já como seu próprio field_row
    # (label + valor + tag) — em vez da lista única e sem classificação
    # de antes (Alterações no RAD 2025, item 1 do Excel completo).
    # Ligações reais faturadas/não faturadas NÃO aparecem mais aqui —
    # mudaram de seção para "operacional" (ver rad_loader.py), saem
    # junto com IV — Indicadores Operacionais, ao lado das tarifas.
    merge_band(r, "VI — " + SECOES["investimentos"].upper(), NAVY2, _SECTION_FONT); r += 1

    campos_investimentos = {c: m for c, m in campos_por_secao.get("investimentos", [])}
    _GRUPOS_INVESTIMENTOS_FICHA = [
        ("ÁGUA", "2F6FB0", [
            "obras_agua", "melhorias_agua",
            "invest_fonte_propria_agua", "invest_fonte_onerosa_agua",
            "invest_fonte_nao_onerosa_agua", "investimento_total_agua",
        ]),
        ("ESGOTO", ORANGE, [
            "obras_esgoto", "melhorias_esgoto",
            "invest_fonte_propria_esgoto", "invest_fonte_onerosa_esgoto",
            "invest_fonte_nao_onerosa_esgoto", "investimento_total_esgoto",
        ]),
        ("ATIVOS DE USO COMPARTILHADO", GRAY, [
            "construcao_compartilhado", "veiculos_compartilhado",
            "maquinarios_compartilhado", "reforma_compartilhado",
            "obras_andamento_compartilhado", "outros_investimentos",
        ]),
    ]
    for titulo_grupo, cor_grupo, campos_grupo in _GRUPOS_INVESTIMENTOS_FICHA:
        merge_band(r, titulo_grupo, cor_grupo, _SUBSECTION_FONT, altura=20); r += 1
        for campo in campos_grupo:
            field_row(r, campos_investimentos[campo], dados.get(campo)); r += 1
        r += 1  # respiro entre sub-bandas

    merge_band(r, "VII — " + SECOES["contratual"].upper(), GRAY, _SECTION_FONT); r += 1
    for campo, meta in campos_por_secao.get("contratual", []):
        field_row(r, meta, dados.get(campo)); r += 1
    r += 1

    ws.merge_cells(f"A{r}:D{r}")
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
    ws.freeze_panes = "A6"

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
            ws.freeze_panes = "A6"

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
    """Faixa colorida (B:D) usada para segregar sub-blocos dentro do
    Relatório de Validação de Investimentos (Água / Esgoto / Ativos
    Compartilhados) — mesmo estilo visual das bandas da ficha completa
    (merge_band em _escrever_ficha_municipio), só que restrita a B:D
    porque a coluna A do relatório de validação é decorativa."""
    ws.merge_cells(f"B{row}:D{row}")
    c = ws.cell(row=row, column=2, value=texto)
    c.font = _HEADER_TABLE_FONT
    c.fill = _fill(cor)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row].height = altura
    ws.cell(row=row, column=1).fill = _fill(cor)


def _linha_campo_validacao(ws, row: int, meta: dict, valor) -> None:
    """1 linha (Campo | Código/Fonte | Valor Atual) do relatório de
    validação — mesmo estilo da linha genérica usada dentro do loop
    padrão de gerar_excel_validacao_area, só extraída para função para
    poder ser reaproveitada pelo layout segregado de Investimentos."""
    valor_fmt = _fmt_valor_excel(valor)
    label = meta["label"]
    if meta.get("unidade") and "(" not in label:
        label += f" ({meta['unidade']})"
    ws.cell(row=row, column=2, value=label).font = _LABEL_FONT
    ws.cell(row=row, column=3, value=meta.get("codigo", "")).font = _LABEL_FONT
    vc = ws.cell(row=row, column=4, value=valor_fmt if valor_fmt is not None else "a preencher")
    vc.font = _VALUE_FONT if valor_fmt is not None else _VALUE_PENDING_FONT
    vc.alignment = Alignment(horizontal="center", vertical="center")
    for col in "ABCD":
        cell = ws[f"{col}{row}"]
        cell.border = _BORDER_ALL
        if col != "D":
            cell.fill = _fill(LIGHT_BG)
    ws.row_dimensions[row].height = _ROW_HEIGHT


# Grupos do Relatório de Validação de Investimentos — cada um vira uma
# banda colorida própria (Água = azul do nosso layout, Esgoto = laranja,
# Compartilhado = cinza neutro). "Contábil"/GECONT não existe mais como
# área separada (Alterações no RAD 2025, item 5) — a fonte de
# investimento em água + total entram direto na banda ÁGUA.
_GRUPOS_INVESTIMENTO = [
    ("ÁGUA", "2F6FB0", [
        "obras_agua", "melhorias_agua",
        "invest_fonte_propria_agua", "invest_fonte_onerosa_agua",
        "invest_fonte_nao_onerosa_agua", "investimento_total_agua",
    ]),
    ("ESGOTO", "C1560E", [
        "obras_esgoto", "melhorias_esgoto",
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
    from scripts.rad_loader import AREAS_VALIDACAO, campos_por_area, CAMPOS_CONTEXTO, CAMPOS_RAD

    nome_area = AREAS_VALIDACAO[area]
    cor_area = _CORES_AREA[area]
    campos = campos_por_area(area)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Validacao_{nome_area}_{municipio}"[:31]

    widths = {"A": 3, "B": 40, "C": 24, "D": 20}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.sheet_view.showGridLines = False

    r = 1
    ws.merge_cells(f"A{r}:D{r}")
    c = ws.cell(row=r, column=1, value=f"RELATÓRIO DE VALIDAÇÃO — {nome_area.upper()}")
    c.font = _HEADER_FONT; c.fill = _fill(cor_area)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[r].height = 30
    r += 1

    ws.merge_cells(f"A{r}:D{r}")
    c = ws.cell(row=r, column=1, value=f"MUNICÍPIO: {municipio.upper()} — ANO {ano}")
    c.font = _SUBHEADER_FONT; c.fill = _fill(NAVY2)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[r].height = 22
    r += 1

    ws.merge_cells(f"A{r}:D{r}")
    instr = ws.cell(row=r, column=1, value="Conferir os valores atuais abaixo com a área responsável. Devolver eventuais divergências à equipe de Regulação.")
    instr.font = _LEGEND_FONT; instr.fill = _fill(LEGEND_BG)
    instr.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[r].height = 30
    r += 2

    # Cabeçalho de contexto (Informações do Município) — sempre presente,
    # para quem valida saber a que município os dados se referem.
    ws.merge_cells(f"A{r}:D{r}")
    ctx_title = ws.cell(row=r, column=1, value="CONTEXTO DO MUNICÍPIO")
    ctx_title.font = _SUBSECTION_FONT; ctx_title.fill = _fill(GRAY)
    ctx_title.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[r].height = 20
    r += 1
    for campo in CAMPOS_CONTEXTO:
        meta = CAMPOS_RAD[campo]
        valor = _fmt_valor_excel(dados.get(campo))
        texto_valor = valor if valor is not None else "—"
        lc = ws.cell(row=r, column=2, value=meta["label"])
        lc.font = _LABEL_FONT
        # Localidades Atendidas é texto longo (lista de 1 a 9+ localidades)
        # — mescla C:D (mais largura = menos linhas de quebra = linha bem
        # mais baixa, ver Alterações no RAD 2025 item 2), alinhado à
        # esquerda com quebra de linha.
        if campo == "localidades_atendidas":
            ws.merge_cells(f"C{r}:D{r}")
            largura_valor = ws.column_dimensions["C"].width + ws.column_dimensions["D"].width
            vc = ws.cell(row=r, column=3, value=texto_valor)
            vc.font = _VALUE_FONT
            vc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            ws.row_dimensions[r].height = _altura_com_quebra(texto_valor, largura_valor)
        else:
            vc = ws.cell(row=r, column=3, value=texto_valor)
            vc.font = _VALUE_FONT
            vc.alignment = Alignment(horizontal="center")
            ws.row_dimensions[r].height = _ROW_HEIGHT
        for col in "ABCD":
            ws[f"{col}{r}"].border = _BORDER_ALL
            ws[f"{col}{r}"].fill = _fill(LIGHT_BG)
        r += 1
    r += 1
    linha_contexto_fim = r

    if area in _GRUPOS_POR_AREA:
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
            for i, h in enumerate(["Campo", "Código / Fonte", "Valor Atual"], start=2):
                c = ws.cell(row=r, column=i, value=h)
                c.font = _LABEL_FONT
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
            _banda_validacao(ws, r, "ETE(S) DO MUNICÍPIO — TRATAMENTO E EFICIÊNCIA", GRAY)
            r += 1
            if len(etes) <= 1:
                ete = etes[0] if etes else None
                ws.cell(row=r, column=2, value="ETE — nome / tratamento / eficiência").font = _LABEL_FONT
                ws.cell(row=r, column=3, value="1 ou mais por município").font = _LABEL_FONT
                if ete:
                    trat = f"{_fmt_valor_excel(ete['tratamento_pct'])}%" if ete["tratamento_pct"] is not None else "trat. a preencher"
                    efic = f"{_fmt_valor_excel(ete['eficiencia_pct'])}%" if ete["eficiencia_pct"] is not None else "efic. a preencher"
                    texto_valor = f"{ete['nome']} — {trat} — {efic}"
                else:
                    texto_valor = "a preencher"
                vc = ws.cell(row=r, column=4, value=texto_valor)
                vc.font = _VALUE_FONT if ete else _VALUE_PENDING_FONT
                vc.alignment = Alignment(horizontal="center", vertical="center")
                for col in "ABCD":
                    cell = ws[f"{col}{r}"]
                    cell.border = _BORDER_ALL
                    if col != "D":
                        cell.fill = _fill(LIGHT_BG)
                ws.row_dimensions[r].height = _ROW_HEIGHT
                r += 1
            else:
                for i, ete in enumerate(etes, start=1):
                    ws.cell(row=r, column=2, value=f"ETE {i} de {len(etes)}").font = _LABEL_FONT
                    ws.cell(row=r, column=3, value="1 ou mais por município").font = _LABEL_FONT
                    trat = f"{_fmt_valor_excel(ete['tratamento_pct'])}%" if ete["tratamento_pct"] is not None else "trat. a preencher"
                    efic = f"{_fmt_valor_excel(ete['eficiencia_pct'])}%" if ete["eficiencia_pct"] is not None else "efic. a preencher"
                    vc = ws.cell(row=r, column=4, value=f"{ete['nome']} — {trat} — {efic}")
                    vc.font = _VALUE_FONT
                    vc.alignment = Alignment(horizontal="center", vertical="center")
                    for col in "ABCD":
                        cell = ws[f"{col}{r}"]
                        cell.border = _BORDER_ALL
                        if col != "D":
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
        # A tabela de verdade cobre só B:D; A recebe cor/borda à parte.
        ws.cell(row=r, column=1).fill = _fill(cor_area)
        headers = ["Campo", "Código / Fonte", "Valor Atual"]
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
        # linha(s) extra na mesma tabela (mesmas 4 colunas: rótulo/código/valor),
        # já com tratamento% e eficiência% individuais de cada ETE.
        if area == "esgoto":
            etes = dados.get("etes", [])
            if len(etes) <= 1:
                ete = etes[0] if etes else None
                ws.cell(row=r, column=2, value="ETE — nome / tratamento / eficiência").font = _LABEL_FONT
                ws.cell(row=r, column=3, value="1 ou mais por município").font = _LABEL_FONT
                if ete:
                    trat = f"{_fmt_valor_excel(ete['tratamento_pct'])}%" if ete["tratamento_pct"] is not None else "trat. a preencher"
                    efic = f"{_fmt_valor_excel(ete['eficiencia_pct'])}%" if ete["eficiencia_pct"] is not None else "efic. a preencher"
                    texto_valor = f"{ete['nome']} — {trat} — {efic}"
                else:
                    texto_valor = "a preencher"
                vc = ws.cell(row=r, column=4, value=texto_valor)
                vc.font = _VALUE_FONT if ete else _VALUE_PENDING_FONT
                vc.alignment = Alignment(horizontal="center", vertical="center")
                for col in "ABCD":
                    cell = ws[f"{col}{r}"]
                    cell.border = _BORDER_ALL
                    if col != "D":
                        cell.fill = _fill(LIGHT_BG)
                ws.row_dimensions[r].height = _ROW_HEIGHT
                r += 1
            else:
                for i, ete in enumerate(etes, start=1):
                    ws.cell(row=r, column=2, value=f"ETE {i} de {len(etes)}").font = _LABEL_FONT
                    ws.cell(row=r, column=3, value="1 ou mais por município").font = _LABEL_FONT
                    trat = f"{_fmt_valor_excel(ete['tratamento_pct'])}%" if ete["tratamento_pct"] is not None else "trat. a preencher"
                    efic = f"{_fmt_valor_excel(ete['eficiencia_pct'])}%" if ete["eficiencia_pct"] is not None else "efic. a preencher"
                    vc = ws.cell(row=r, column=4, value=f"{ete['nome']} — {trat} — {efic}")
                    vc.font = _VALUE_FONT
                    vc.alignment = Alignment(horizontal="center", vertical="center")
                    for col in "ABCD":
                        cell = ws[f"{col}{r}"]
                        cell.border = _BORDER_ALL
                        if col != "D":
                            cell.fill = _fill(LIGHT_BG)
                    ws.row_dimensions[r].height = _ROW_HEIGHT
                    r += 1

        last_row = r - 1
        tabela = Table(displayName=f"Validacao{area.capitalize()}", ref=f"B{linha_cabecalho_tabela}:D{last_row}")
        tabela.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(tabela)

    r += 1
    ws.merge_cells(f"A{r}:D{r}")
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
