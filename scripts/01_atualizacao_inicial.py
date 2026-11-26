import json, sys, re
from copy import copy
import openpyxl
from openpyxl.chart import BarChart, Reference

SRC, RAW, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
C, L, J, T = ("Consultoria / adequação tributária", "Adequação legislativa (Código Tributário)",
              "Assessoria jurídica tributária", "Capacitação / treinamento")

SEL = {
 0:(C,"Diagnóstico e modernização da administração tributária (subsídios à legislação e capacitação)"),
 2:(L,"Consultoria em administração tributária e revisão da legislação do ITBI"),
 9:(J,"Consultoria e assessoria jurídica para a pasta tributária"),
 18:(C,"Programa de modernização da administração tributária (reestruturação legal + 151h de capacitação)"),
 19:(L,"Revisão e adequação do Código Tributário Municipal à Reforma Tributária"),
 24:(J,"Assessoria jurídica e consultoria em gestão tributária municipal"),
 25:(J,"Assessoria jurídica tributária — revisão da sistemática dos tributos"),
 26:(J,"Consultoria, auditoria tributária e suporte jurídico"),
 30:(C,"Assessoria e consultoria técnica tributária e de gestão pública"),
 31:(L,"Anteprojeto do novo Código Tributário Municipal"),
 33:(L,"Revisão da LC 07/2013 e minuta de novo Código Tributário"),
 34:(C,"Consultoria e apoio técnico em gestão tributária municipal"),
 35:(J,"Assessoria jurídica tributária, consultoria e auditoria fiscal"),
 37:(C,"Consultoria e assessoria técnica à Secretaria de Tributação"),
 38:(C,"Apoio técnico em matérias tributárias e previdenciárias complexas"),
 40:(J,"Serviços advocatícios em Direito Tributário (consultivo e judicial)"),
 42:(J,"Consultoria e assessoria em Direito Público e Tributário (setor de tributos)"),
 44:(C,"Consultoria e assessoria fiscal e tributária"),
 45:(J,"Assessoria e consultoria jurídico-tributária"),
 52:(C,"Consultoria tributária para otimização da arrecadação própria"),
 54:(C,"Assessoria e consultoria tributária (planejamento e acompanhamento)"),
 56:(L,"Revisão, atualização e consolidação do Código Tributário"),
 59:(L,"Atualização e modernização da legislação tributária"),
 63:(L,"Atualização do Código Tributário à EC 132/2023 e LC 214/2025"),
 64:(L,"Assessoria na revisão do Código Tributário"),
 74:(L,"Revisão e consolidação do novo Código Tributário Municipal"),
 75:(L,"Elaboração do novo Código Tributário Municipal"),
 77:(C,"Consultoria tributária, auditoria fiscal e assessoria econômico-financeira"),
 78:(L,"Código de Defesa do Contribuinte e revisão do Código Tributário"),
 79:(J,"Assessoria e consultoria jurídica tributária"),
 80:(L,"Assessoria tributária para elaboração do novo Código Tributário"),
 83:(J,"Assessoria jurídica nas áreas tributária e de planejamento urbano"),
 84:(C,"Assessoria técnica tributária (lançamento, arrecadação e fiscalização)"),
 91:(J,"Assessoria e consultoria jurídica administrativa tributária"),
 95:(C,"Consultoria, auditoria e capacitação em governança tributária e fiscal"),
 99:(C,"Assessoria e consultoria tributária"),
 101:(C,"Consultoria tributária com treinamento de servidores"),
 104:(C,"Consultoria tributária — revisão da planta genérica de valores"),
 108:(L,"Assessoria tributária à Câmara na reforma do Código Tributário"),
 120:(C,"Apoio técnico para adequação à Reforma Tributária do Consumo"),
 122:(C,"Assessoria tributária consultiva continuada"),
 124:(C,"Assessoria e consultoria contábil e tributária"),
 136:(C,"Consultoria tributária com treinamentos e capacitação"),
 137:(C,"Assessoria técnica em gestão tributária municipal"),
 142:(C,"Consultoria tributária continuada — modernização e revisão normativa"),
 147:(J,"Serviços jurídicos em matéria fiscal e tributária"),
 149:(J,"Consultoria e assessoria jurídica tributária (esfera administrativa)"),
 152:(J,"Assessoria em direito público e tributário p/ incremento da arrecadação"),
 153:(C,"Consultoria e assessoria tributária — fortalecimento da administração"),
 156:(C,"Assessoria tributária — apuração de ISSQN não recolhido"),
 157:(C,"Consultoria e assessoramento estratégico à Administração Tributária"),
 165:(L,"Revisão e modernização da legislação tributária municipal"),
 186:(C,"Assessoria e consultoria tributária"),
 187:(C,"Assessoria e suporte técnico em matéria tributária"),
 190:(C,"Assessoria tributária e financeira"),
 191:(C,"Implementação da Reforma Tributária e NFAg (LC 214/2025)"),
 192:(L,"Novo Código Tributário Municipal (sociedade de advogados)"),
 193:(C,"Modernização da gestão tributária municipal"),
 203:(C,"Assessoria e consultoria técnica em gestão tributária"),
 205:(J,"Assessoria jurídica tributária, consultoria e auditoria fiscal"),
 248:(J,"Consultoria e assessoria jurídica tributária (esfera administrativa)"),
 251:(C,"Consultoria e assessoria tributária com ênfase no ITR"),
 252:(J,"Assessoria jurídica e auditoria fiscal tributária"),
 266:(J,"Assessoria, auditoria e consultoria jurídico-tributária"),
 269:(C,"Consultoria, auditoria e capacitação em governança tributária e fiscal"),
 284:(C,"Adequações à Reforma Tributária do Consumo e NFAg"),
 291:(L,"Anteprojeto do novo Código Tributário (EC 132/2023 e LC 214/2025)"),
 300:(C,"Consultoria e assessoria tributária"),
 301:(C,"Consultoria em gestão e planejamento tributário"),
 # capacitação
 4:(T,"Curso on-line: Reforma Tributária do Consumo nas retenções (IBS/CBS)"),
 5:(T,"Curso on-line: Reforma Tributária do Consumo nas retenções (IBS/CBS)"),
 7:(T,"Seminário: processo legislativo e impactos da Reforma Tributária"),
 12:(T,"Curso on-line: Reforma Tributária do Consumo nas retenções (IBS/CBS)"),
 16:(T,"Curso: Reforma Tributária do Consumo nas retenções (IBS/CBS)"),
 17:(T,"Curso on-line: Reforma Tributária do Consumo nas retenções (IBS/CBS)"),
 22:(T,"Congresso FENAFIM/CONCAAT — Reforma Tributária (IBS/CBS)"),
 23:(T,"Curso presencial: Reforma Tributária e reequilíbrio de contratos"),
 32:(T,"Curso: Reforma Tributária nos pequenos municípios"),
 43:(T,"Curso: Reforma Tributária do Consumo nas retenções (IBS/CBS)"),
 46:(T,"Curso presencial: Reforma Tributária do Consumo nas retenções"),
 49:(T,"Curso: Reforma Tributária do Consumo nas retenções (IBS/CBS)"),
 50:(T,"Capacitação: Reforma Tributária na contabilidade, licitações e contratos"),
 57:(T,"Evento sobre Reforma Tributária (6 servidores)"),
 65:(T,"Capacitação in company: Reforma Tributária na contabilidade pública"),
 68:(T,"Congresso FENAFIM/CONCAAT — Reforma Tributária (7 servidores do fisco)"),
 73:(T,"Curso: redação de leis, processo legislativo e Reforma Tributária"),
 100:(T,"Congresso FENAFIM/CONCAAT — Reforma Tributária (5 servidores)"),
 115:(T,"Congresso FENAFIM/CONCAAT — Reforma Tributária"),
 123:(T,"Curso: Contratos públicos na transição da Reforma Tributária"),
 127:(T,"Curso: Reforma Tributária nos contratos administrativos (reequilíbrio)"),
 129:(T,"Encontro de vereadores: arrecadação municipal e os novos tributos"),
 130:(T,"Encontro de vereadores: arrecadação municipal e os novos tributos"),
 134:(T,"Seminário: impacto da Reforma Tributária e da IA na gestão pública"),
 135:(T,"Seminário: impacto da Reforma Tributária e da IA na gestão pública"),
 144:(T,"Curso: Impactos da Reforma Tributária nas contratações públicas"),
 148:(T,"Curso: retenções tributárias e Reforma Tributária (EC 132/23) — 6 inscrições"),
 155:(T,"Congresso FENAFIM/CONCAAT — servidores da Fazenda"),
 159:(T,"Curso: planilha de preços (IN 05/2017) e impactos da Reforma Tributária"),
 170:(T,"Curso: Reforma Tributária do Consumo nas retenções"),
 173:(T,"Curso: licitações com IA e Reforma Tributária (4 inscrições)"),
 175:(T,"Capacitação: licitações, IA e efeitos da Reforma Tributária"),
 178:(T,"Curso: fiscalização de contratos e impactos da Reforma Tributária"),
 180:(T,"Curso: Impactos da Reforma Tributária nos contratos (29 inscrições)"),
 181:(T,"Curso: novas regras da Reforma Tributária (EC 132/2023)"),
 182:(T,"Capacitação on-line: Reforma Tributária e gestão fazendária (30 servidores)"),
 183:(T,"Curso: fiscalização de contratos e impactos da Reforma Tributária"),
 185:(T,"Curso: fiscalização de contratos e impactos da Reforma Tributária"),
 188:(T,"Seminário: administração municipal, Reforma Tributária e TC"),
 198:(T,"Curso: Reforma Tributária além do IBS e CBS"),
 199:(T,"Curso: Reforma Tributária além do IBS e CBS"),
 201:(T,"Curso: Reforma Tributária na gestão e fiscalização das contratações"),
 202:(T,"Curso on-line: Reforma Tributária na Administração Pública"),
 220:(T,"Curso: licitações com IA e Reforma Tributária"),
 221:(T,"Curso: Contratos públicos na transição da Reforma Tributária"),
 224:(T,"Curso: retenção na fonte de tributos e IBS/CBS"),
 241:(T,"Curso: Reforma Tributária para a Administração Pública"),
 242:(T,"Capacitação sobre Reforma Tributária"),
 243:(T,"Formação avançada em Reforma Tributária (40 servidores da SEFIN)"),
 244:(T,"Curso: Reforma Tributária do Consumo nas retenções (IBS/CBS)"),
 246:(T,"Congresso Baiano de Direito Financeiro — impactos da Reforma Tributária"),
 247:(T,"Capacitação: planilhas de custos e efeitos da Reforma Tributária"),
 260:(T,"Curso: Reforma Tributária aplicada aos contratos administrativos"),
 261:(T,"4º Fórum Mineiro de Reforma Tributária"),
 262:(T,"Workshop Legislativo/Executivo: licitações e Reforma Tributária (13 inscrições)"),
 263:(T,"Treinamento avançado: Reforma Tributária, fiscalização e legislação tributária"),
 264:(T,"Congresso Baiano de Direito Financeiro — impactos da Reforma Tributária"),
 265:(T,"Curso: Impacto da Reforma Tributária para os municípios (IBS/CBS)"),
 282:(T,"Curso: fiscalização de contratos e impactos da Reforma Tributária"),
 283:(T,"Curso: licitações com IA e Reforma Tributária (4 inscrições)"),
 285:(T,"Curso on-line: contratos administrativos e Reforma Tributária"),
 286:(T,"Seminário Nacional: Reforma Tributária nas contratações públicas"),
 287:(T,"Congresso Baiano de Direito Financeiro — impactos da Reforma Tributária"),
 288:(T,"Seminário: contratos administrativos e Reforma Tributária"),
 289:(T,"Curso: reequilíbrio econômico-financeiro — da Reforma Tributária à planilha"),
 290:(T,"Seminário Reforma Tributária pelo Brasil"),
 295:(T,"Congresso FENAFIM/CONCAAT — Reforma Tributária"),
 297:(T,"Curso: Impactos da Reforma Tributária nos contratos administrativos"),
 298:(T,"Curso FAMURS: Reforma Tributária além do IBS e CBS"),
 299:(T,"Seminário Nacional — Reforma Tributária (on-line)"),
}

def ente(org):
    o = org.upper()
    if "CAMARA" in o: return "Câmara Municipal"
    if any(k in o for k in ("AUTONOMO", "INSTITUTO", "FUNDACAO", "FACULDADE")): return "Autarquia/fundação municipal"
    if "EMPRESA" in o: return "Empresa pública municipal"
    return "Prefeitura"

def modal(m):
    return {"Pregão - Eletrônico": "Pregão Eletrônico", "Concorrência - Eletrônica": "Concorrência Eletrônica",
            "Pregão - Presencial": "Pregão Presencial"}.get(m, m)

def valor(x):
    for v in (x["hom"], x["est"]):
        if v is not None and 10 < v < 1e9: return round(float(v), 2)
    return None

raw = json.load(open(RAW, encoding="utf-8"))
raw.sort(key=lambda x: x["pub"])

wb = openpyxl.load_workbook(SRC)
ws = wb["Levantamento PNCP"]
last_old = ws.max_row  # 97
existing = {ws.cell(r, 10).value for r in range(5, last_old + 1)}
n_old = last_old - 4

new = []
for i in sorted(SEL, key=lambda k: (raw[k]["pub"], raw[k]["mun"])):
    x = raw[i]
    assert x["canc"] is False, x["c"]
    if x["c"] in existing: continue
    cat, obj = SEL[i]
    m = re.match(r"(\d{14})-\d+-(\d+)/(\d{4})", x["c"])
    link = f"https://pncp.gov.br/app/editais/{m.group(1)}/{m.group(3)}/{int(m.group(2))}"
    new.append([ente(x["org"]), x["mun"], x["uf"], obj, cat, modal(x["mod"]), x["pub"], valor(x), x["c"], link])

tmpl = [ws.cell(last_old, c) for c in range(1, 12)]
r = last_old
for k, row in enumerate(new, start=1):
    r += 1
    vals = [float(n_old + k)] + row[:-1] + ["Abrir no PNCP"]
    for c, v in enumerate(vals, start=1):
        cell = ws.cell(r, c, v)
        cell._style = copy(tmpl[c - 1]._style)
    ws.cell(r, 11).hyperlink = row[-1]
    ws.cell(r, 11).style  # keep blue font copied
    if ws.row_dimensions[last_old].height:
        ws.row_dimensions[r].height = ws.row_dimensions[last_old].height
last = r
ws.auto_filter.ref = f"A4:K{last}"

ws["A2"] = (f"Fonte: Portal Nacional de Contratações Públicas (PNCP) — pncp.gov.br. Amostra consolidada a partir de buscas por "
            f"'reforma tributária', 'consultoria/assessoria tributária', 'Código Tributário' e 'capacitação'. Consulta inicial em 21/09/2026 "
            f"(itens 1–{n_old}); atualização em 26/11/2026 (itens {n_old+1}–{n_old+len(new)}: publicações de 21/09 a 26/11/2026, esfera municipal). "
            f"Nos itens da atualização, o valor é o total homologado no PNCP ou, sem homologação, o estimado; itens 1–{n_old} sem valor — abra o link.")
ws.row_dimensions[2].height = max(ws.row_dimensions[2].height or 15, 45)

# ---------------- Resumo ----------------
rs = wb["Resumo"]
st = {k: [copy(rs.cell(row, c)._style) for c in range(1, 4)] for k, row in
      {"sec": 12, "hdr": 13, "item": 14, "tot": 19, "vsec": 40, "vhdr": 41, "vitem": 42, "vnote": 48}.items()}
vrows = [[rs.cell(row, c).value for c in (1, 2)] for row in range(42, 48)]
note = rs.cell(48, 1).value

for row in rs.iter_rows(min_row=6, max_row=9, min_col=2, max_col=3):
    for cell in row:
        cell.value = cell.value.replace("$97", f"${last}")

for row in range(12, rs.max_row + 1):
    for c in range(1, 4):
        rs.cell(row, c).value = None
        rs.cell(row, c)._style = copy(rs.cell(3, 1)._style)

L_ = "'Levantamento PNCP'!"
def put(row, key, vals):
    for c, v in enumerate(vals, start=1):
        cell = rs.cell(row, c, v); cell._style = copy(st[key][c - 1])

def table(start, title, col, items):
    put(start, "sec", [title, None, None]); put(start + 1, "hdr", ["Item", "Qtde", "Valor (R$)"])
    r0 = start + 2
    for k, it in enumerate(items):
        rr = r0 + k
        put(rr, "item", [it, f"=COUNTIF({L_}${col}$5:${col}${last},A{rr})",
                         f"=SUMIF({L_}${col}$5:${col}${last},A{rr},{L_}$I$5:$I${last})"])
    rt = r0 + len(items)
    put(rt, "tot", ["Total", f"=SUM(B{r0}:B{rt-1})", f"=SUM(C{r0}:C{rt-1})"])
    return r0, rt

all_rows = [[ws.cell(rr, c).value for c in range(1, 11)] for rr in range(5, last + 1)]
order = ["Inexigibilidade", "Dispensa", "Pregão Eletrônico", "Pregão Presencial", "Concorrência Eletrônica", "Credenciamento", "Não informado"]
mods = [m for m in order if any(x[6] == m for x in all_rows)] + sorted({x[6] for x in all_rows} - set(order))
ufs = sorted({x[3] for x in all_rows})

m0, mt = table(12, "Por modalidade de contratação", "G", mods)
u0, ut = table(mt + 2, "Por UF", "D", ufs)
v = ut + 3
put(v, "vsec", ["Valores contratados", None, None]); put(v + 1, "vhdr", ["Indicador", "Valor", None])
for k, (a, f) in enumerate(vrows):
    put(v + 2 + k, "vitem", [a, f.replace("$97", f"${last}"), None])
put(v + 8, "vnote", [note, None, None])

# Leitura rápida
cnt = len(all_rows)
rs["D5"] = f"• Amostra de {cnt} contratações municipais no PNCP ({n_old} da consulta inicial + {len(new)} da atualização de 26/11/2026; universo real é maior)."
rs["D8"] = "• Forte pulverização geográfica (" + ", ".join(ufs) + ")."
rs["D11"] = "• Cautela: o TCE-PR já suspendeu cautelarmente contrato de consultoria tributária (Cascavel/PR). Atualização exclui recuperação de créditos (IRRF, PASEP, CFEM, FPM, previdenciários), VAF/ICMS, software e cadastro imobiliário."

# charts
rs._charts = []
def bar(title, r0, r1, anchor, w, h):
    ch = BarChart(); ch.type = "col"; ch.title = title; ch.legend = None
    ch.add_data(Reference(rs, min_col=2, min_row=r0, max_row=r1), titles_from_data=False)
    ch.set_categories(Reference(rs, min_col=1, min_row=r0, max_row=r1))
    ch.series[0].graphicalProperties.solidFill = "4F81BD"
    ch.roundedCorners = False; ch.title.overlay = False
    ch.x_axis.delete = False; ch.y_axis.delete = False
    ch.width, ch.height = w, h
    rs.add_chart(ch, anchor)
bar("Contratações por categoria", 6, 9, "D13", 11.7, 6.4)
bar("Contratações por UF", u0, ut - 1, f"D{u0}", 16, 7.5)

wb.save(OUT)
print(json.dumps({"novos": len(new), "ultima_linha": last, "mods": mods, "ufs": ufs,
                  "com_valor": sum(1 for x in new if x[7] is not None)}, ensure_ascii=False))
