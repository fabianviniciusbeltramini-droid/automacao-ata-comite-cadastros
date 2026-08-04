"""
Gerador de Ata - Comitê de Cadastros
=====================================
Como usar:
  1. Preencha o arquivo 'itens_reuniao.xlsx' com os produtos de cada comprador
  2. Execute: python gerar_ata.py
  3. A ata será gerada como 'Ata_Comite_Cadastros.docx'
"""

import sys, subprocess, os, zipfile
from datetime import datetime

try:
    import openpyxl
except ImportError:
    print("Instalando openpyxl...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl", "--user"])
    import openpyxl


# ── Leitura do Excel ──────────────────────────────────────────────────────────
def ler_excel(caminho_excel):
    wb = openpyxl.load_workbook(caminho_excel)
    ws = wb.active
    compradores = {}
    ordem = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[0]:
            continue
        comprador    = str(row[0]).strip()
        cod_barras   = str(row[1]).strip() if row[1] else "Sem Código"
        nome_produto = str(row[2]).strip() if row[2] else ""
        if not nome_produto:
            continue
        if comprador not in compradores:
            compradores[comprador] = []
            ordem.append(comprador)
        compradores[comprador].append({"cod": cod_barras, "nome": nome_produto})
    return [(c, compradores[c]) for c in ordem]


# ── Helpers XML ───────────────────────────────────────────────────────────────
def esc(s):
    return (str(s)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;"))

_cb_id = [0]

def run(texto, bold=False, color="000000", size=20, font="Arial"):
    b = "<w:b/>" if bold else ""
    return (
        f'<w:r><w:rPr>{b}'
        f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}"/>'
        f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
        f'<w:color w:val="{color}"/>'
        f'</w:rPr>'
        f'<w:t xml:space="preserve">{esc(texto)}</w:t></w:r>'
    )

def run_cb():
    """Checkbox interativo clicável no Word."""
    _cb_id[0] += 1
    uid = _cb_id[0]
    return (
        f'<w:sdt>'
        f'<w:sdtPr><w:id w:val="{uid}"/>'
        f'<w14:checkbox>'
        f'<w14:checked w14:val="0"/>'
        f'<w14:checkedState w14:val="2612" w14:font="MS Gothic"/>'
        f'<w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/>'
        f'</w14:checkbox>'
        f'</w:sdtPr>'
        f'<w:sdtContent>'
        f'<w:r><w:rPr>'
        f'<w:rFonts w:ascii="MS Gothic" w:hAnsi="MS Gothic" w:cs="MS Gothic"/>'
        f'<w:sz w:val="20"/><w:szCs w:val="20"/>'
        f'</w:rPr><w:t>&#x2610;</w:t></w:r>'
        f'</w:sdtContent>'
        f'</w:sdt>'
    )

def para(runs_xml, align="left", before=0, after=60,
         shading=None, border_bottom=None, border_top=None, indent_left=0):
    jc = {"left":"left","center":"center","right":"right"}.get(align,"left")
    pPr = f'<w:jc w:val="{jc}"/>'
    pPr += f'<w:spacing w:before="{before}" w:after="{after}"/>'
    if indent_left:
        pPr += f'<w:ind w:left="{indent_left}"/>'
    if shading:
        pPr += f'<w:shd w:val="clear" w:color="auto" w:fill="{shading}"/>'
    borders = ""
    if border_bottom:
        borders += f'<w:bottom w:val="single" w:sz="{border_bottom[1]}" w:space="1" w:color="{border_bottom[0]}"/>'
    if border_top:
        borders += f'<w:top w:val="single" w:sz="{border_top[1]}" w:space="1" w:color="{border_top[0]}"/>'
    if borders:
        pPr += f'<w:pBdr>{borders}</w:pBdr>'
    return f'<w:p><w:pPr>{pPr}</w:pPr>{"".join(runs_xml)}</w:p>'

def heading(runs_xml, nivel=1, collapsed=True,
            before=0, after=0, shading=None, indent_left=0,
            border_bottom=None, border_top=None, line=None):
    """
    Parágrafo com estilo Heading (Ttulo1/2/3) + w15:collapsed no próprio heading.
    Esse é o segredo para abrir recolhido no Word!
    """
    style = f"Ttulo{nivel}"
    pPr = f'<w:pStyle w:val="{style}"/>'
    if line:
        pPr += f'<w:spacing w:before="{before}" w:after="{after}" w:line="{line}" w:lineRule="auto"/>'
    else:
        pPr += f'<w:spacing w:before="{before}" w:after="{after}"/>'
    if indent_left:
        pPr += f'<w:ind w:left="{indent_left}"/>'
    if shading:
        pPr += f'<w:shd w:val="clear" w:color="auto" w:fill="{shading}"/>'
    borders = ""
    if border_bottom:
        borders += f'<w:bottom w:val="single" w:sz="{border_bottom[1]}" w:space="1" w:color="{border_bottom[0]}"/>'
    if border_top:
        borders += f'<w:top w:val="single" w:sz="{border_top[1]}" w:space="1" w:color="{border_top[0]}"/>'
    if borders:
        pPr += f'<w:pBdr>{borders}</w:pBdr>'
    # Suprime espaço extra que o Word adiciona automaticamente entre headings
    pPr += '<w:contextualSpacing/>'
    # w15:collapsed no próprio heading = abre recolhido!
    if collapsed:
        pPr += '<w15:collapsed/>'
    return f'<w:p><w:pPr>{pPr}</w:pPr>{"".join(runs_xml)}</w:p>'


# ── Monta o XML do documento ──────────────────────────────────────────────────
def build_document_xml(compradores_itens, data_reuniao):
    body = []

    # ── Título principal (sem collapsed — fica sempre visível) ────────────────
    body.append(para(
        [run("ATA DA REUNIÃO – Comitê de Cadastros", bold=True, color="1F3864", size=32)],
        align="center", before=0, after=120, border_bottom=("1F3864", 8)
    ))

    def info(label, valor):
        return para([run(label, bold=True, size=20), run(valor, size=20)], before=40, after=40)

    body.append(info("Reunião: ", "Comitê de Cadastros"))
    body.append(info("Data: ", data_reuniao))
    body.append(info("Local: ", "Sala Rede Top"))
    body.append(info("Horário de início: ", "8h00"))

    # Abertura — heading nível 1, sem collapsed (visível)
    # ── Caixa de participantes ───────────────────────────────────────────────
    membros = "Douglas, Edgar, Eduardo, Fabian, Jardel, Zamboni, Joel, Erica"
    participantes = ", ".join([c for c, _ in compradores_itens])

    def caixa_participantes():
        def para_caixa(runs_xml, before=20, after=20):
            pPr = f'<w:jc w:val="right"/><w:spacing w:before="{before}" w:after="{after}"/><w:ind w:left="2000" w:right="200"/>'
            return f'<w:p><w:pPr>{pPr}</w:pPr>{"".join(runs_xml)}</w:p>'

        return [
            para_caixa([run("Integrantes do Comitê, Presentes:", bold=True, size=18)], before=80, after=10),
            para_caixa([run(membros, size=18)], before=0, after=40),
            para_caixa([run("Participantes / Apresentações:", bold=True, size=18)], before=20, after=10),
            para_caixa([run(participantes, size=18)], before=0, after=80),
        ]

    for p in caixa_participantes():
        body.append(p)

    body.append(para([run("", size=20)], before=0, after=80))

    body.append(heading(
        [run("1. Abertura", bold=True, color="1F3864", size=24)],
        nivel=1, collapsed=False, before=120, after=80
    ))
    body.append(para([run(
        "As 08h00, o Comitê de Cadastros iniciou os trabalhos. Foi confirmada a lista de "
        "participantes e apresentado o objetivo da reunião, que consistiu na análise e "
        "apresentação de itens para cadastro, inclusão e substituições de produtos.", size=20
    )], before=0, after=120))

    # Itens Apresentados — heading nível 1, sem collapsed (visível)
    body.append(heading(
        [run("2. Itens Apresentados", bold=True, color="1F3864", size=24)],
        nivel=1, collapsed=False, before=0, after=0, line=480
    ))

    # ── Compradores ───────────────────────────────────────────────────────────
    for comprador, itens in compradores_itens:

        # Heading nível 1 — COMPRADOR — collapsed=True (recolhido ao abrir)
        body.append(heading(
            [run(f"Itens apresentados por {comprador}:", bold=True, color="000000", size=22)],
            nivel=1, collapsed=True,
            before=0, after=80,
            indent_left=200
        ))

        for item in itens:
            titulo = (f"{item['cod']} - {item['nome']}"
                      if item['cod'] != "Sem Código"
                      else f"Sem Código - {item['nome']}")

            # Heading nível 2 — ITEM — collapsed=True (recolhido ao abrir)
            # "Reprovado" fica na mesma linha do nome do produto
            body.append(heading(
                [run(titulo + " | ", size=20),
                 run("Reprovado ", bold=True, size=20),
                 run_cb()],
                nivel=2, collapsed=True,
                before=80, after=10,
                border_bottom=("DDDDDD", 2)
            ))

            def linha_lojas(tipo, opcoes):
                rs = [run(tipo + " ", bold=True, size=20)]
                for i, op in enumerate(opcoes):
                    rs.append(run_cb())
                    rs.append(run(op + ("   " if i < len(opcoes)-1 else ""), size=20))
                return para(rs, before=0, after=20, indent_left=360)

            body.append(linha_lojas("Lojas Grandes",   ["Premium", "Tradicionais", "Populares"]))
            body.append(linha_lojas("Lojas Médias ",   ["Premium", "Tradicionais", "Populares"]))
            body.append(linha_lojas("Lojas Compactas", ["Premium", "Tradicionais"]))

            body.append(para(
                [run("Bandeira Top ", bold=True, size=20), run_cb(),
                 run("   Bandeira Preceiro ", bold=True, size=20), run_cb()],
                before=20, after=60, indent_left=360
            ))

        # Encaminhamentos específicos — heading nível 2, collapsed=True
        body.append(heading(
            [run("Encaminhamentos Específicos:", bold=True, color="000000", size=20)],
            nivel=2, collapsed=True,
            before=320, after=160
        ))
        body.append(para([run("", size=20)], before=0, after=360))

    # Parágrafo vazio entre último comprador e seção 3
    body.append(para([run("", size=20)], before=0, after=160))

    # ── Rodapé ────────────────────────────────────────────────────────────────
    body.append(heading(
        [run("3. Desenvolvimento das Discussões", bold=True, color="1F3864", size=24)],
        nivel=1, collapsed=True, before=360, after=80, line=360
    ))
    body.append(para([run(
        "Foram discutidos os itens apresentados, qualificando cada proposta de novo produto, "
        "substituição ou bloqueio. Os membros fizeram considerações sobre adequação, possíveis "
        "impactos e continuidade no cadastro dos itens.", size=20
    )], before=0, after=120))

    body.append(heading(
        [run("4. Deliberações e Decisões", bold=True, color="1F3864", size=24)],
        nivel=1, collapsed=True, before=0, after=160
    ))
    body.append(para([run("Encaminhamentos Gerais:", size=20)], before=0, after=40, indent_left=360))
    body.append(para([run("", size=20)], before=0, after=160))

    body.append(heading(
        [run("5. Encerramento", bold=True, color="1F3864", size=24)],
        nivel=1, collapsed=False, before=0, after=80
    ))
    body.append(para([run(
        "Nada mais havendo a tratar, a reunião do Comitê de Cadastros foi oficialmente encerrada. "
        "Esta ata foi lavrada para registro e será arquivada conforme o procedimento interno de "
        "documentação de reuniões.", size=20
    )], before=0, after=0))

    # Namespace w15 incluído no documento — essencial para w15:collapsed funcionar
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
        'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
        'mc:Ignorable="w14 w15">'
        '<w:body>'
        + "".join(body) +
        '<w:sectPr>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" w:header="709" w:footer="709" w:gutter="0"/>'
        '</w:sectPr>'
        '</w:body></w:document>'
    )


# ── Estilos: Ttulo1, Ttulo2, Ttulo3 com outlineLvl ───────────────────────────
STYLES_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
          xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">

  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
      <w:sz w:val="20"/>
    </w:rPr>
  </w:style>

  <!-- Ttulo1 = Heading 1 (comprador) -->
  <w:style w:type="paragraph" w:styleId="Ttulo1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>
      <w:spacing w:before="0" w:after="80"/>
      <w:contextualSpacing/>
      <w:outlineLvl w:val="0"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
      <w:sz w:val="24"/>
    </w:rPr>
  </w:style>

  <!-- Ttulo2 = Heading 2 (item) -->
  <w:style w:type="paragraph" w:styleId="Ttulo2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>
      <w:outlineLvl w:val="1"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
      <w:sz w:val="20"/>
    </w:rPr>
  </w:style>

  <!-- Ttulo3 = Heading 3 (conteúdo filho do item) -->
  <w:style w:type="paragraph" w:styleId="Ttulo3">
    <w:name w:val="heading 3"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>
      <w:outlineLvl w:val="2"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
      <w:sz w:val="20"/>
    </w:rPr>
  </w:style>

</w:styles>'''

SETTINGS_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="851"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
  <w:hideSpellingErrors/>
  <w:hideGrammaticalErrors/>
</w:settings>'''

CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>'''

RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

WORD_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>'''


def gerar_docx(compradores_itens, caminho_saida, data_reuniao):
    _cb_id[0] = 0
    doc_xml = build_document_xml(compradores_itens, data_reuniao)
    with zipfile.ZipFile(caminho_saida, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("_rels/.rels", RELS)
        zf.writestr("word/_rels/document.xml.rels", WORD_RELS)
        zf.writestr("word/document.xml", doc_xml.encode("utf-8"))
        zf.writestr("word/styles.xml", STYLES_XML.encode("utf-8"))
        zf.writestr("word/settings.xml", SETTINGS_XML.encode("utf-8"))
    return True


# ── Execução principal ────────────────────────────────────────────────────────
def main():
    PASTA_ENTRADA = r"X:\Central de Negocios\Cadastro\Fabian\Automação - Ata Comitê de Cadastro\Base Dados Completa - Geração da Ata"
    SAIDA_BASE = r"X:\Central de Negocios\Cadastro\Fabian\Automação - Ata Comitê de Cadastro\Ata Final"

    if not os.path.isdir(PASTA_ENTRADA):
        print(f"❌ Pasta de entrada não encontrada:\n   {PASTA_ENTRADA}")
        sys.exit(1)

    arquivos_xlsx = [f for f in os.listdir(PASTA_ENTRADA) if f.endswith(".xlsx") and not f.startswith("~")]

    if not arquivos_xlsx:
        print("Erro: nenhum arquivo .xlsx encontrado na pasta.")
        sys.exit(1)

    if len(arquivos_xlsx) > 1:
        print("Mais de um arquivo .xlsx encontrado:")
        for i, f in enumerate(arquivos_xlsx, 1):
            print(f"  {i}. {f}")
        while True:
            escolha = input("Digite o número do arquivo: ").strip()
            if escolha.isdigit() and 1 <= int(escolha) <= len(arquivos_xlsx):
                EXCEL = arquivos_xlsx[int(escolha) - 1]
                break
            print("Opção inválida.")
    else:
        EXCEL = arquivos_xlsx[0]
        print(f"Planilha encontrada: '{EXCEL}'")

    print(f"Lendo planilha '{EXCEL}'...")
    compradores_itens = ler_excel(os.path.join(PASTA_ENTRADA, EXCEL))

    if not compradores_itens:
        print("Nenhum dado encontrado na planilha.")
        sys.exit(1)

    total_itens = sum(len(itens) for _, itens in compradores_itens)
    print(f"  → {len(compradores_itens)} compradores | {total_itens} itens no total")

    meses = ["janeiro","fevereiro","março","abril","maio","junho",
             "julho","agosto","setembro","outubro","novembro","dezembro"]

    while True:
        entrada = input("Qual a data da reunião? (DD/MM/AAAA) ou Enter para hoje: ").strip()
        if entrada == "":
            hoje = datetime.today()
            data_str = f"{hoje.day} de {meses[hoje.month-1]} de {hoje.year}"
            break
        try:
            data = datetime.strptime(entrada, "%d/%m/%Y")
            data_str = f"{data.day} de {meses[data.month-1]} de {data.year}"
            break
        except ValueError:
            print("  Formato inválido. Use DD/MM/AAAA (ex: 04/06/2026)")

    nome_arquivo = data.strftime("%d %m %Y") + " - Ata.docx" if entrada != "" else hoje.strftime("%d %m %Y") + " - Ata.docx"
    SAIDA = os.path.join(SAIDA_BASE, nome_arquivo)

    print(f"Gerando ata para {data_str}...")
    ok = gerar_docx(compradores_itens, SAIDA, data_str)

    if ok:
        print(f"\n✅ Ata gerada com sucesso: '{SAIDA}'")
    else:
        print("\n❌ Falha ao gerar a ata.")
        sys.exit(1)

if __name__ == "__main__":
    main()
