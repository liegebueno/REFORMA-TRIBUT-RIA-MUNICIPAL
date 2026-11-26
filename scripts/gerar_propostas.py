import copy, sys
from docx import Document
from docx.shared import Pt, Cm, Emu, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_COLOR_INDEX
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement, parse_xml

TIMBRE = "timbre.png"
FONT = "Arial"


def set_border(p, sides=("top", "left", "bottom", "right"), sz=6):
    pPr = p._p.get_or_add_pPr()
    b = OxmlElement("w:pBdr")
    for s in sides:
        e = OxmlElement(f"w:{s}")
        e.set(qn("w:val"), "single"); e.set(qn("w:sz"), str(sz)); e.set(qn("w:space"), "4"); e.set(qn("w:color"), "000000")
        b.append(e)
    pPr.append(b)


def run(p, text, b=False, i=False, u=False, size=12, hl=False):
    r = p.add_run(text)
    r.font.name = FONT; r._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    r.font.size = Pt(size); r.bold = b; r.italic = i; r.underline = u
    if hl: r.font.highlight_color = WD_COLOR_INDEX.YELLOW
    return r


def para(doc, parts, align="j", indent=True, after=12, line=1.5, left=0, keep=False, size=12):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.alignment = {"j": WD_ALIGN_PARAGRAPH.JUSTIFY, "c": WD_ALIGN_PARAGRAPH.CENTER, "r": WD_ALIGN_PARAGRAPH.RIGHT, "l": WD_ALIGN_PARAGRAPH.LEFT}[align]
    pf.space_after = Pt(after); pf.space_before = Pt(0); pf.line_spacing = line
    if indent: pf.first_line_indent = Cm(1.25)
    if left: pf.left_indent = Cm(left)
    pf.keep_with_next = keep
    if isinstance(parts, str): parts = [(parts, {})]
    for t, kw in parts:
        run(p, t, size=kw.pop("size", size), **kw)
    return p


def heading(doc, text):
    p = para(doc, [(text, {"b": True})], align="l", indent=False, after=12, keep=True)
    p.paragraph_format.space_before = Pt(6)
    set_border(p, sides=("bottom",), sz=8)
    return p


def sub(doc, text):
    return para(doc, [(text, {"b": True, "i": True})], align="j", indent=False, after=6, keep=True)


def quote(doc, lines):
    for k, t in enumerate(lines):
        para(doc, [(t, {"i": True, "size": 11})], indent=False, left=2.0, line=1.0, after=12 if k == len(lines) - 1 else 2)


def item(doc, lead, text_parts):
    parts = [(lead + " ", {"b": True})] + text_parts
    return para(doc, parts)


def add_background(section):
    hdr = section.header; hdr.is_linked_to_previous = False
    p = hdr.paragraphs[0] if hdr.paragraphs else hdr.add_paragraph()
    r = p.add_run()
    r.add_picture(TIMBRE, width=Cm(21.0), height=Cm(29.7))
    inline = r._element.find(".//" + qn("wp:inline"))
    graphic = inline.find(qn("a:graphic"))
    docPr = inline.find(qn("wp:docPr"))
    cx, cy = Cm(21.0), Cm(29.7)
    anchor = parse_xml(
        f'<wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        f'distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="0" behindDoc="1" locked="1" layoutInCell="1" allowOverlap="1">'
        f'<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{int(cx)}" cy="{int(cy)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
        f'<wp:docPr id="{docPr.get("id")}" name="Timbre Monteiro"/><wp:cNvGraphicFramePr/></wp:anchor>')
    anchor.append(graphic)
    inline.getparent().replace(inline, anchor)
    section.footer.is_linked_to_previous = False


def new_doc():
    doc = Document()
    st = doc.styles["Normal"]; st.font.name = FONT; st.font.size = Pt(12)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    s = doc.sections[0]
    s.page_width, s.page_height = Cm(21.0), Cm(29.7)
    s.top_margin, s.bottom_margin = Emu(1985 * 635), Emu(1701 * 635)
    s.left_margin, s.right_margin = Emu(2835 * 635), Emu(1134 * 635)
    s.header_distance, s.footer_distance = Emu(709 * 635), Emu(567 * 635)
    add_background(s)
    return doc


def address_box(doc, linhas, assunto):
    ps = []
    for t in linhas:
        ps.append(para(doc, t if isinstance(t, list) else [(t, {"b": True})], align="j", indent=False, after=0, keep=True))
    ps.append(para(doc, [("ASSUNTO: ", {"b": True}), (assunto, {"i": True})], align="j", indent=False, after=0))
    for p in ps: set_border(p)
    para(doc, "", indent=False, after=6)


def signature(doc):
    para(doc, "Atenciosamente,", align="c", indent=False, after=48)
    for t in ["MONTEIRO E MONTEIRO ADVOGADOS ASSOCIADOS", "BRUNO ROMERO PEDROSA MONTEIRO", "OAB/PE Nº 11.338"]:
        para(doc, [(t, {"b": True})], align="c", indent=False, after=0, line=1.15)


def annexes(doc, docs):
    p = doc.add_paragraph(); p.add_run().add_break(WD_BREAK.PAGE)
    para(doc, [("RELAÇÃO DE DOCUMENTOS ANEXOS", {"b": True})], align="c", indent=False, after=12)
    t = doc.add_table(rows=len(docs), cols=2); t.alignment = WD_TABLE_ALIGNMENT.CENTER
    widths = [Cm(2.6), Cm(12.4)]
    for k, txt in enumerate(docs):
        cells = t.rows[k].cells
        for c, w in zip(cells, widths):
            c.width = w
            if k % 2 == 1:
                tcPr = c._tc.get_or_add_tcPr(); sh = OxmlElement("w:shd")
                sh.set(qn("w:val"), "clear"); sh.set(qn("w:color"), "auto"); sh.set(qn("w:fill"), "F2F2F2"); tcPr.append(sh)
        a = cells[0].paragraphs[0]; run(a, f"DOC. {k+1:02d}", b=True); a.paragraph_format.space_after = Pt(4)
        bpar = cells[1].paragraphs[0]; run(bpar, txt); bpar.paragraph_format.space_after = Pt(4)


# ---------------------------------------------------------------- textos comuns
REFORMA = [
    "A Emenda Constitucional nº 132/2023, regulamentada pela Lei Complementar nº 214/2025, promoveu a mais profunda reforma do Sistema Tributário Nacional desde 1988. O Imposto Sobre Serviços (ISS), de competência municipal, e o ICMS estadual serão gradativamente extintos e substituídos pelo Imposto sobre Bens e Serviços (IBS), de competência compartilhada entre Estados e Municípios, ao lado da Contribuição sobre Bens e Serviços (CBS) e do Imposto Seletivo (IS).",
    "A tributação passa a observar o princípio do destino — arrecada-se onde o bem ou serviço é consumido — e a arrecadação, a fiscalização e a distribuição do IBS deixam de ser realizadas diretamente pelo Município, passando a um órgão nacional, o Comitê Gestor do IBS. Instaura-se, ademais, dupla transição: a dos tributos (2026 a 2033) e a da própria repartição federativa da receita, que se estenderá até 2078.",
]
RISCOS_MUN = "Esse novo arranjo, se não enfrentado tempestivamente, expõe o Município a riscos concretos: perda de arrecadação na migração para o critério de destino; dependência de repasses calculados e distribuídos por órgão central, sem que o ente disponha de instrumentos próprios para conferir sua exatidão; defasagem da legislação tributária municipal e dos cadastros mobiliário e imobiliário; e a necessária reconversão da estrutura fiscal, que deixa de arrecadar o ISS para auditar os repasses do IBS."
FPM = "Soma-se a isso o impacto sobre as transferências constitucionais recebidas da União. O Fundo de Participação dos Municípios (FPM), historicamente composto por parcela do Imposto de Renda e do IPI, teve sua base afetada pela extinção do IPI, tendo a Emenda Constitucional nº 132/2023 incluído o Imposto Seletivo na repartição do fundo — alteração cujos efeitos financeiros exigem acompanhamento técnico permanente, sob pena de o Município vir a receber a menor."

ESCOPO = [
    ("a) Diagnóstico e simulação de impacto financeiro:", "levantamento da receita própria e das transferências do Município e simulação dos efeitos da transição do ISS para o IBS sobre a arrecadação;"),
    ("b) Adequação jurídica:", "revisão e atualização da legislação tributária municipal e da Lei Orgânica, incluída a definição da alíquota própria do IBS, com elaboração das minutas de projetos de lei e de regulamentação;"),
    ("c) Adequação operacional:", "saneamento dos cadastros mobiliário e imobiliário e apoio à integração aos sistemas nacionais da reforma;"),
    ("d) Capacitação:", "formação de servidores, procuradores e gestores para atuação no novo modelo de tributação sobre o consumo;"),
    ("e) Monitoramento e auditoria dos repasses:", "acompanhamento continuado dos repasses do IBS e das transferências constitucionais durante o período de transição, com vistas à identificação e à recuperação de eventuais diferenças em favor do erário."),
]

ART74 = [
    "“Art. 74. É inexigível a licitação quando inviável a competição, em especial nos casos de:",
    "III - contratação dos seguintes serviços técnicos especializados de natureza predominantemente intelectual com profissionais ou empresas de notória especialização (...):",
    "c) assessorias ou consultorias técnicas e auditorias financeiras ou tributárias;",
    "(...)",
    "e) patrocínio ou defesa de causas judiciais ou administrativas;",
    "§ 3º Para fins do disposto no inciso III (...), considera-se de notória especialização o profissional ou a empresa cujo conceito no campo de sua especialidade, decorrente de desempenho anterior, estudos, experiência, publicações, organização, aparelhamento, equipe técnica ou outros requisitos relacionados com suas atividades, permita inferir que o seu trabalho é essencial e reconhecidamente adequado à plena satisfação do objeto do contrato.”",
]

ANEXOS = [
    "RELAÇÃO DE ATUAÇÃO PERANTE ENTES PÚBLICOS E ASSOCIAÇÕES MUNICIPALISTAS;",
    "ATESTADOS DE CAPACIDADE TÉCNICA;",
    "RECOMENDAÇÃO Nº 36/2016 DO CNMP;",
    "MANIFESTAÇÃO DA AGU SOBRE CONTRATAÇÃO DE SERVIÇOS TÉCNICOS ESPECIALIZADOS;",
    "PRECEDENTE STF;",
    "PRECEDENTE STJ;",
    "ACÓRDÃO NA ÍNTEGRA DO TEMA 309 DA REPERCUSSÃO GERAL (STF);",
    "CERTIDÕES DE REGULARIDADE FISCAL, TRABALHISTA E PREVIDENCIÁRIA.",
]


def corpo_comum(doc, ente, consorcio=False):
    # 3. notória especialização
    heading(doc, "3. DA NOTÓRIA ESPECIALIZAÇÃO DO PROPONENTE")
    para(doc, "A matéria de que ora se trata — adequação ao novo Sistema Tributário Nacional e proteção da receita pública diante da Reforma Tributária — encerra elevada complexidade técnica, exigindo conhecimento especializado e integrado das áreas jurídica, contábil e econômica, o que transcende a rotina habitual das Procuradorias Municipais.")
    para(doc, [("O escritório ", {}), ("MONTEIRO E MONTEIRO ADVOGADOS ASSOCIADOS", {"b": True}), (" reúne os requisitos necessários à contratação por inexigibilidade:", {})])
    item(doc, "a) Volume e Abrangência:", [("patrocínio de mais de 1.000 (mil) demandas em favor de Municípios em todo o território nacional, atuando também na defesa dos interesses de Municípios filiados a diversas Associações Municipalistas (", {}), ("DOC. 01", {"b": True}), (");", {})])
    item(doc, "b) Corpo Técnico:", [("equipe técnica multidisciplinar — advogados, contadores e economistas — especializada em Direito Tributário, Financeiro e Constitucional aplicado à Gestão Pública;", {})])
    item(doc, "c) Capacidade Técnica Atestada:", [("capacidade técnica já atestada por entes públicos e entidades municipalistas, conforme documentação anexa (", {}), ("DOC. 02", {"b": True}), (").", {})])

    # 4. fundamentos
    heading(doc, "4. DOS FUNDAMENTOS JURÍDICOS DA INEXIGIBILIDADE DE LICITAÇÃO")
    para(doc, [("A contratação direta mediante ", {}), ("Inexigibilidade de Licitação", {"b": True}), (" ampara-se no regime da Lei nº 14.133/2021, no Estatuto da Advocacia e na jurisprudência do Supremo Tribunal Federal:", {})])
    sub(doc, "4.1. Enquadramento na Lei nº 14.133/2021")
    para(doc, [("O objeto contratual subsume-se às hipóteses do ", {}), ("artigo 74, inciso III, alíneas “c” e “e”, c/c § 3º, da Lei nº 14.133/2021", {"b": True}), (":", {})])
    quote(doc, ART74)
    sub(doc, "4.2. Natureza Singular e Técnica dos Serviços de Advocacia (Lei nº 8.906/1994)")
    para(doc, [("O ", {}), ("artigo 3º-A do Estatuto da OAB", {"b": True}), (" (incluído pela Lei nº 14.039/2020) estabelece a singularidade aplicável ao labor advocatício:", {})])
    quote(doc, ["Art. 3º-A. Os serviços profissionais de advogado são, por sua natureza, técnicos e singulares, quando comprovada sua notória especialização, nos termos da lei."])
    sub(doc, "4.3. Jurisprudência do Supremo Tribunal Federal (Tema 309 da Repercussão Geral)")
    para(doc, [("O Plenário do Supremo Tribunal Federal, ao julgar os REs nº 656.558/SP e 610.523/SP (", {}), ("Tema 309", {"b": True}), ("), reconheceu a licitude da contratação direta de serviços advocatícios por inexigibilidade, ante a inviabilidade de competição por critérios meramente objetivos:", {})])
    quote(doc, ["“Há determinados serviços que demandam primor técnico diferenciado, detido por pequena ou individualizada parcela de pessoas, as quais imprimem neles características diferenciadas e pessoais. Trata-se de serviços cuja especialização requer aporte subjetivo, o denominado ‘toque do especialista’, distinto de um para outro, o qual os qualifica como singular” (Rel. Min. Dias Toffoli)."])
    para(doc, "No mesmo julgamento, firmou-se que a remuneração deve ser compatível com a responsabilidade profissional exigida pelo caso e observar o valor médio praticado pela contratada em situações similares anteriores.")
    sub(doc, "4.4. Respaldo do CNMP, da AGU e dos Tribunais Superiores")
    para(doc, [("A possibilidade de os Entes Públicos contratarem, pela via da inexigibilidade, serviços técnicos especializados e singulares encontra entendimento favorável do ", {}), ("Conselho Nacional do Ministério Público (Recomendação nº 36/2016)", {"b": True}), (" (", {}), ("DOC. 03", {"b": True}), (") e da ", {}), ("Advocacia-Geral da União", {"b": True}), (" (", {}), ("DOC. 04", {"b": True}), ("), sendo também uníssona a jurisprudência do Supremo Tribunal Federal e do Superior Tribunal de Justiça nesse sentido (", {}), ("DOCs. 05, 06 e 07", {"b": True}), (").", {})])


def proposta_financeira(doc, consorcio=False, fee=None):
    heading(doc, "5. DA PROPOSTA FINANCEIRA")
    alvo = "de cada Município consorciado" if consorcio else "do Município"
    para(doc, f"Em consonância com as balizas fixadas no Tema 309/STF — que estabelece como parâmetro de justificativa do preço a compatibilidade com a responsabilidade do encargo e o valor médio cobrado pelo prestador em casos análogos —, propõe-se a fixação dos honorários de forma variável e proporcional ao porte econômico {alvo}, tomando-se como base de referência o valor do Fundo de Participação dos Municípios (FPM), nas seguintes modalidades, que podem ser adotadas isolada ou cumulativamente:")
    item(doc, "a) Honorário fixo por projeto:", [("correspondente a ", {}), (fee[0] if fee else "[__]%", {"b": True, "hl": not fee}), (" do valor anual do FPM do Município, referente ao escopo efetivamente contratado (diagnóstico e simulação de impacto, adequação jurídica e adequação operacional), passível de pagamento parcelado conforme cronograma de entregas;", {})])
    item(doc, "b) Honorário mensal:", [("correspondente a ", {}), (fee[1] if fee else "[__]%", {"b": True, "hl": not fee}), (" do valor do repasse mensal do FPM do Município, destinado ao acompanhamento continuado e à auditoria permanente dos repasses recebidos durante o período de transição, com identificação de eventuais diferenças em favor do erário;", {})])
    item(doc, "c) Honorário por ações de capacitação e treinamento:", [("no valor de ", {}), (fee[2] if fee else "R$ [________] ([valor por extenso])", {"b": True, "hl": not fee}), (", por turma ou na modalidade in loco, voltado à formação de servidores, procuradores e gestores.", {})])
    if consorcio:
        para(doc, "Os honorários serão apurados individualmente para cada Município consorciado, conforme o respectivo valor do FPM, podendo ser custeados mediante contrato de rateio, quanto aos serviços de natureza coletiva, e/ou por instrumento próprio de cada ente, quanto às entregas individualizadas, na forma do modelo de contratação acima detalhado.")


def finais(doc, consorcio=False):
    heading(doc, "6. CONSIDERAÇÕES FINAIS")
    para(doc, [("O proponente atende integralmente a todas as exigências legais de habilitação jurídica, qualificação técnica e regularidade fiscal, previdenciária e trabalhista perante os órgãos públicos competentes (", {}), ("DOC. 08", {"b": True}), (").", {})])
    para(doc, "Vem assim propor os serviços com base na documentação ora anexada, e em conformidade com a Lei Federal nº 14.133/2021, para que se digne Vossa " + ("Senhoria" if consorcio else "Excelência") + " determinar a abertura do procedimento de inexigibilidade, observados os ditames legais pertinentes.")
    para(doc, "Sem mais para o momento, colocamo-nos à inteira disposição para os esclarecimentos que se fizerem necessários e renovamos os protestos de elevada estima e consideração.")
    signature(doc)
    annexes(doc, ANEXOS)


# ---------------------------------------------------------------- MUNICÍPIO
FEE = ("0,5% (meio por cento)", "0,4% (quatro décimos por cento)", "R$ 15.000,00 (quinze mil reais)")

def municipio(out, mun, uf, prefeito, data):
    doc = new_doc()
    para(doc, f"Recife/PE, {data}", align="r", indent=False, after=18)
    address_box(doc, [f"À PREFEITURA MUNICIPAL DE {mun.upper()} - {uf}", f"EXMO.(A) SENHOR(A) PREFEITO(A) {prefeito.upper()}"],
                "PROPOSTA DE PRESTAÇÃO DE SERVIÇOS TÉCNICOS ESPECIALIZADOS – ASSESSORIA E CONSULTORIA TÉCNICA MULTIDISCIPLINAR PARA ADEQUAÇÃO À REFORMA TRIBUTÁRIA (EC Nº 132/2023 E LC Nº 214/2025) – CONTRATAÇÃO DIRETA POR INEXIGIBILIDADE DE LICITAÇÃO (ART. 74, III, DA LEI Nº 14.133/2021).")
    para(doc, [("Excelentíssimo(a) Senhor(a) Prefeito(a),", {"b": True})], indent=True, after=12)
    para(doc, [("Cumprimentando-o(a) cordialmente, o escritório ", {}), ("MONTEIRO E MONTEIRO ADVOGADOS ASSOCIADOS", {"b": True}),
               (" vem, por meio desta, submeter à apreciação de Vossa Excelência proposta de prestação de serviços técnicos especializados de ", {}),
               ("assessoria e consultoria multidisciplinar para adequação do Município à Reforma Tributária", {"b": True}),
               (", com vistas à proteção e ao incremento de suas receitas diante do novo modelo constitucional de tributação sobre o consumo.", {})])
    heading(doc, "1. DA REFORMA TRIBUTÁRIA E DOS SEUS IMPACTOS SOBRE O MUNICÍPIO")
    for t in REFORMA: para(doc, t)
    para(doc, RISCOS_MUN); para(doc, FPM)
    heading(doc, "2. DO OBJETO E DO ESCOPO DOS SERVIÇOS")
    para(doc, "É nesse cenário que se propõe a prestação de assessoria e consultoria técnica de natureza multidisciplinar, compreendendo, conforme o interesse da Administração:")
    for lead, txt in ESCOPO: item(doc, lead, [(txt, {})])
    corpo_comum(doc, "Município")
    proposta_financeira(doc, fee=FEE)
    finais(doc)
    doc.save(out)


# ---------------------------------------------------------------- CONSÓRCIO
def consorcio(out, data):
    doc = new_doc()
    para(doc, f"Recife/PE, {data}", align="r", indent=False, after=18)
    address_box(doc, [[("AO CONSÓRCIO ", {"b": True}), ("[NOME DO CONSÓRCIO] – [UF]", {"b": True, "hl": True})],
                      [("A/C DO EXMO.(A) SR.(A) PRESIDENTE ", {"b": True}), ("[NOME DO PRESIDENTE]", {"b": True, "hl": True})]],
                "PROPOSTA DE PRESTAÇÃO DE SERVIÇOS TÉCNICOS ESPECIALIZADOS – ASSESSORIA E CONSULTORIA TÉCNICA MULTIDISCIPLINAR, EM BENEFÍCIO DOS MUNICÍPIOS CONSORCIADOS, PARA ADEQUAÇÃO À REFORMA TRIBUTÁRIA (EC Nº 132/2023 E LC Nº 214/2025) – CONTRATAÇÃO DIRETA POR INEXIGIBILIDADE DE LICITAÇÃO (ART. 74, III, DA LEI Nº 14.133/2021).")
    para(doc, [("Senhor(a) Presidente,", {"b": True})], after=12)
    para(doc, [("Cumprimentando-o(a) cordialmente, o escritório ", {}), ("MONTEIRO E MONTEIRO ADVOGADOS ASSOCIADOS", {"b": True}),
               (" vem, por meio desta, trazer a Vossa Senhoria informações relevantes acerca dos impactos da Reforma Tributária sobre as finanças dos Municípios consorciados e da imprescindível necessidade de adequação dos entes, com vistas à proteção e ao incremento de suas receitas diante do novo modelo constitucional de tributação sobre o consumo — tarefa em que o Consórcio, por sua vocação regional, pode atuar de forma coordenada e economicamente vantajosa.", {})])
    heading(doc, "1. DA REFORMA TRIBUTÁRIA E DOS SEUS IMPACTOS SOBRE OS MUNICÍPIOS")
    for t in REFORMA: para(doc, t)
    para(doc, "A migração para o novo modelo expõe os Municípios a riscos concretos de perda de arrecadação e de dependência de repasses calculados por órgão central, além de impor a atualização da legislação tributária, o saneamento de cadastros e a reconversão da estrutura fiscal — que deixa de arrecadar o ISS para auditar os repasses do IBS. Soma-se o impacto sobre as transferências constitucionais da União, notadamente o Fundo de Participação dos Municípios (FPM), cuja base foi afetada pela extinção do IPI e cuja evolução exige acompanhamento técnico permanente.")
    heading(doc, "2. DO OBJETO, DO ESCOPO E DA FORMA DE CONTRATAÇÃO PELO CONSÓRCIO")
    para(doc, "Propõe-se a prestação de assessoria e consultoria técnica multidisciplinar, em benefício dos Municípios consorciados, compreendendo:")
    esc = list(ESCOPO)
    esc[3] = ("d) Capacitação regional:", "formação de servidores, procuradores e gestores dos Municípios consorciados para atuação no novo modelo de tributação sobre o consumo;")
    esc[4] = ("e) Monitoramento e auditoria dos repasses:", "implantação de monitoramento continuado e auditoria dos repasses do IBS e das transferências constitucionais, com identificação e recuperação de eventuais diferenças.")
    for lead, txt in esc: item(doc, lead, [(txt, {})])
    para(doc, "Os serviços podem ser contratados de forma centralizada pelo Consórcio, em duas configurações compatíveis com a legislação:")
    item(doc, "(i)", [("mediante contratação direta pelo próprio Consórcio, com custeio por contrato de rateio entre os consorciados, quanto aos serviços de natureza coletiva; e", {})])
    item(doc, "(ii)", [("mediante contratação/licitação compartilhada, em que o Consórcio conduz e gerencia o procedimento e cada Município ratifica os termos e formaliza o respectivo instrumento contratual, com seus quantitativos e dotação orçamentária própria, modelo já reconhecido pelos Tribunais de Contas.", {})])
    corpo_comum(doc, "Consórcio", consorcio=True)
    proposta_financeira(doc, consorcio=True)
    finais(doc, consorcio=True)
    doc.save(out)


if __name__ == "__main__":
    municipio(sys.argv[1], "Rosário do Ivaí", "PR", "Anizio Cesar Lino Silva", "26 de novembro de 2026")
    consorcio(sys.argv[2], "26 de novembro de 2026")
    print("ok")
