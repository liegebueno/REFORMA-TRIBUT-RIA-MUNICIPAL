import json, sys, unicodedata
from copy import copy
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.comments import Comment
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList

F = sys.argv[1]
wb = openpyxl.load_workbook(F)
ws = wb["Levantamento PNCP"]
LAST = ws.max_row
while ws.cell(LAST, 1).value is None: LAST -= 1

# ---------- coluna de população ----------
n = lambda s: "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn").upper().strip()
E = json.load(open("prop/entes.json", encoding="utf-8"))["items"]
POP = {(n(e["ente"]), e["uf"]): e["populacao"] for e in E if e["esfera"] == "M"}
h = ws["L4"]; h.value = "População (IBGE)"; h._style = copy(ws["J4"]._style)
miss = []
for r in range(5, LAST + 1):
    c = ws.cell(r, 12); c._style = copy(ws.cell(r, 8)._style)
    p = POP.get((n(ws.cell(r, 3).value), ws.cell(r, 4).value))
    c.value = p; c.number_format = "#,##0"
    if p is None: miss.append(ws.cell(r, 3).value)
ws.column_dimensions["L"].width = 14
ws.auto_filter.ref = f"A4:L{LAST}"
ws["L4"].comment = Comment("População estimada usada pelo Tesouro Nacional (SICONFI, exercício 2026), consultada em 26/11/2026.", "Levantamento")

# ---------- aba Informativo ----------
if "Informativo de Valores" in wb.sheetnames: del wb["Informativo de Valores"]
inf = wb.create_sheet("Informativo de Valores", 1)
inf.sheet_view.showGridLines = False
L = "'Levantamento PNCP'!"
R = lambda col: f"{L}${col}$5:${col}${LAST}"
Fc, Ic, Dc, Gc, Lc = R("F"), R("I"), R("D"), R("G"), R("L")
CAP = '"Capacitação / treinamento"'
VAL = f"({Ic}>=100)"

NAVY = "1F3864"; LIGHT = "D9E2F3"; GREY = "F2F2F2"
ar = Font(name="Arial", size=10); arb = Font(name="Arial", size=10, bold=True)
blue = Font(name="Arial", size=10, color="0000FF")
hf = Font(name="Arial", size=10, bold=True, color="FFFFFF")
thin = Side(style="thin", color="BFBFBF"); bd = Border(top=thin, bottom=thin, left=thin, right=thin)
BRL = '"R$" #,##0;-"R$" #,##0;"-"'
INT = "#,##0"

def title(r, text, sub=None):
    inf.cell(r, 2, text).font = Font(name="Arial", size=12, bold=True, color=NAVY)
    inf.cell(r, 2).border = Border(bottom=Side(style="medium", color=NAVY))
    for c in range(3, 10): inf.cell(r, c).border = Border(bottom=Side(style="medium", color=NAVY))
    if sub:
        inf.cell(r + 1, 2, sub).font = Font(name="Arial", size=9, italic=True, color="595959")
        return r + 2
    return r + 1

def header(r, cols):
    for j, t in enumerate(cols):
        c = inf.cell(r, 2 + j, t); c.font = hf; c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); c.border = bd
    inf.row_dimensions[r].height = 30

def row(r, vals, fmts, bold=False, shade=False):
    for j, (v, f) in enumerate(zip(vals, fmts)):
        c = inf.cell(r, 2 + j)
        if isinstance(v, tuple) and v[0] == "arr":
            ref = c.coordinate; c.value = ArrayFormula(ref, v[1])
        else:
            c.value = v
        c.font = arb if bold else ar; c.border = bd; c.number_format = f
        c.alignment = Alignment(horizontal="left" if j == 0 else "right", vertical="center")
        if shade: c.fill = PatternFill("solid", fgColor=GREY)

def stats(cond):
    """estatísticas sobre valores >= R$ 100 que atendem a condição (expressão booleana em matriz)"""
    m = f"{cond}*{VAL}"
    return [f"=SUMPRODUCT({m})",
            ("arr", f"=IFERROR(MIN(IF({m},{Ic})),\"-\")"),
            ("arr", f"=IFERROR(QUARTILE(IF({m},{Ic}),1),\"-\")"),
            ("arr", f"=IFERROR(MEDIAN(IF({m},{Ic})),\"-\")"),
            ("arr", f"=IFERROR(QUARTILE(IF({m},{Ic}),3),\"-\")"),
            ("arr", f"=IFERROR(MAX(IF({m},{Ic})),\"-\")"),
            ("arr", f"=IFERROR(AVERAGE(IF({m},{Ic})),\"-\")")]

STAT_H = ["Contratos com valor", "Mínimo", "1º quartil", "Mediana", "3º quartil", "Máximo", "Média"]
STAT_F = [INT, BRL, BRL, BRL, BRL, BRL, BRL]

inf.column_dimensions["A"].width = 2
inf.column_dimensions["B"].width = 44
for col in "CDEFGHI": inf.column_dimensions[col].width = 15

# cabeçalho
inf["B2"] = "Informativo de Valores — Reforma Tributária nos Municípios"
inf["B2"].font = Font(name="Arial", size=16, bold=True, color=NAVY)
inf["B3"] = ("Uso interno · Monteiro e Monteiro Advogados Associados · Base: contratações municipais publicadas no PNCP (2024 a 26/11/2026) "
             "cujo objeto cita expressamente a Reforma Tributária. Valores = homologado ou, sem homologação, estimado. "
             "As estatísticas desconsideram contratos simbólicos de êxito (abaixo de R$ 100).")
inf["B3"].font = Font(name="Arial", size=9, italic=True, color="595959"); inf.merge_cells("B3:I3")
inf["B3"].alignment = Alignment(wrap_text=True, vertical="top"); inf.row_dimensions[3].height = 38

# 1. destaques
r = title(5, "1. Destaques")
kp = [("Contratações no levantamento", f"=COUNTA({R('J')})", INT),
      ("Municípios distintos", f"=SUMPRODUCT(1/COUNTIFS({Dc},{Dc},{R('C')},{R('C')}))", INT),
      ("Valor total contratado", f"=SUM({Ic})", BRL),
      ("Ticket mediano — serviços técnicos", ("arr", f"=MEDIAN(IF(({Fc}<>{CAP})*{VAL},{Ic}))"), BRL),
      ("Ticket mediano — capacitação", ("arr", f"=MEDIAN(IF(({Fc}={CAP})*{VAL},{Ic}))"), BRL),
      ("Serviços técnicos contratados por inexigibilidade", f"=COUNTIFS({Fc},\"<>\"&{CAP},{Gc},\"Inexigibilidade\")/COUNTIF({Fc},\"<>\"&{CAP})", "0.0%")]
for k, (lab, f, fmt) in enumerate(kp):
    rr = r + k
    c = inf.cell(rr, 2, lab); c.font = ar; c.border = bd; c.fill = PatternFill("solid", fgColor=LIGHT)
    v = inf.cell(rr, 3)
    if isinstance(f, tuple): v.value = ArrayFormula(v.coordinate, f[1])
    else: v.value = f
    v.font = Font(name="Arial", size=11, bold=True); v.number_format = fmt; v.border = bd; v.alignment = Alignment(horizontal="right")
r = r + len(kp) + 1

# 2. por categoria
r = title(r, "2. Faixa de preço por tipo de contrato", "Metade dos contratos fica entre o 1º e o 3º quartil; é a faixa de referência para propostas.")
header(r, ["Tipo de contrato"] + STAT_H); r += 1
cats = ["Adequação legislativa (Código Tributário)", "Consultoria / adequação tributária", "Assessoria jurídica tributária", "Capacitação / treinamento"]
for k, cat in enumerate(cats):
    row(r, [cat] + stats(f"({Fc}=\"{cat}\")"), ["@"] + STAT_F, shade=k % 2 == 1); r += 1
row(r, ["Serviços técnicos (exceto capacitação)"] + stats(f"({Fc}<>{CAP})"), ["@"] + STAT_F, bold=True); r += 2

# 3. por porte
r = title(r, "3. Serviços técnicos por porte do município", "População IBGE (coluna L da aba Levantamento PNCP). Exclui capacitação.")
header(r, ["Porte (habitantes)"] + STAT_H); r += 1
bands = [("Até 5 mil", 0, 5000), ("5 a 10 mil", 5000, 10188), ("10 a 20 mil", 10188, 20000), ("20 a 50 mil", 20000, 50000),
         ("50 a 100 mil", 50000, 100000), ("Acima de 100 mil", 100000, 100000000)]
porte0 = r
for k, (lab, lo, hi) in enumerate(bands):
    row(r, [lab] + stats(f"({Fc}<>{CAP})*({Lc}>{lo})*({Lc}<={hi})"), ["@"] + STAT_F, shade=k % 2 == 1); r += 1
porte1 = r - 1
row(r, ["Até 10 mil (coeficiente FPM até 0,6)"] + stats(f"({Fc}<>{CAP})*({Lc}>0)*({Lc}<=10188)"), ["@"] + STAT_F, bold=True)
ate10 = r; r += 2

# 4. modalidade
r = title(r, "4. Serviços técnicos por modalidade de contratação")
header(r, ["Modalidade", "Contratos", "% dos serviços", "Mediana", "Média"]); r += 1
for k, m in enumerate(["Inexigibilidade", "Dispensa", "Pregão Eletrônico", "Concorrência Eletrônica"]):
    cond = f"({Fc}<>{CAP})*({Gc}=\"{m}\")"
    row(r, [m, f"=SUMPRODUCT({cond})", f"=IFERROR(SUMPRODUCT({cond})/SUMPRODUCT(({Fc}<>{CAP})*1),0)",
            ("arr", f"=IFERROR(MEDIAN(IF({cond}*{VAL},{Ic})),\"-\")"), ("arr", f"=IFERROR(AVERAGE(IF({cond}*{VAL},{Ic})),\"-\")")],
        ["@", INT, "0.0%", BRL, BRL], shade=k % 2 == 1); r += 1
r += 1

# 5. PR x Brasil
r = title(r, "5. Paraná x Brasil — serviços técnicos")
header(r, ["Recorte"] + STAT_H); r += 1
row(r, ["Paraná"] + stats(f"({Fc}<>{CAP})*({Dc}=\"PR\")"), ["@"] + STAT_F); r += 1
row(r, ["Paraná — municípios até 10 mil hab."] + stats(f"({Fc}<>{CAP})*({Dc}=\"PR\")*({Lc}>0)*({Lc}<=10188)"), ["@"] + STAT_F, shade=True)
pr10 = r; r += 1
row(r, ["Brasil"] + stats(f"({Fc}<>{CAP})"), ["@"] + STAT_F, bold=True); r += 2

# 6. caso Rosário do Ivaí
r = title(r, "6. Caso prático — Rosário do Ivaí/PR (proposta de 26/11/2026)", "Células em azul são premissas editáveis; o restante é calculado.")
header(r, ["Item", "Valor", "Observação"]); inf.merge_cells(start_row=r, start_column=4, end_row=r, end_column=9); r += 1

def line(lab, val, fmt, obs="", inp=False, bold=False):
    global r
    a = inf.cell(r, 2, lab); a.font = arb if bold else ar; a.border = bd
    b = inf.cell(r, 3, val); b.font = blue if inp else (arb if bold else ar); b.number_format = fmt; b.border = bd; b.alignment = Alignment(horizontal="right")
    if inp: b.fill = PatternFill("solid", fgColor="FFF2CC")
    o = inf.cell(r, 4, obs); o.font = Font(name="Arial", size=9, color="595959"); inf.merge_cells(start_row=r, start_column=4, end_row=r, end_column=9)
    o.alignment = Alignment(wrap_text=True, vertical="center")
    if len(str(obs)) > 70: inf.row_dimensions[r].height = 26
    r += 1
    return b.coordinate

pop = line("População (IBGE)", 5484, INT, "SICONFI 2026 — faixa de coeficiente 0,6 do FPM.", inp=True)
fpm = line("FPM anual (out/2025 a set/2026)", 16161880.74, BRL, "Tesouro Transparente — Transferências Constitucionais para Municípios (repasses líquidos de FUNDEB), consultado em 26/11/2026.", inp=True)
fpmm = line("FPM médio mensal", f"={fpm}/12", BRL, "Varia de cerca de R$ 0,97 mi a R$ 1,78 mi conforme o mês.")
p1 = line("Honorário fixo por projeto (% do FPM anual)", 0.005, "0.0%", "Percentual da proposta.", inp=True)
p2 = line("Honorário mensal (% do repasse mensal do FPM)", 0.004, "0.0%", "Percentual da proposta.", inp=True)
p3 = line("Capacitação por turma", 15000, BRL, "Valor da proposta.", inp=True)
nt = line("Turmas de capacitação no 1º ano", 1, INT, "Premissa para o total do 1º ano.", inp=True)
v1 = line("Honorário fixo estimado", f"={fpm}*{p1}", BRL, "Projeto: diagnóstico, adequação jurídica e operacional.", bold=True)
v2 = line("Honorário mensal estimado", f"={fpmm}*{p2}", BRL, "Monitoramento e auditoria dos repasses.", bold=True)
v3 = line("Total estimado no 1º ano", f"={v1}+12*{v2}+{nt}*{p3}", BRL, "Fixo + 12 mensalidades + turmas.", bold=True)
line("Total do 1º ano como % do FPM anual", f"={v3}/{fpm}", "0.00%", "")
line("Referência: mediana de serviços em municípios até 10 mil hab.", f"=F{ate10}", BRL, "Seção 3 deste informativo.")
line("Referência: mediana de serviços no PR até 10 mil hab.", f"=F{pr10}", BRL, "Seção 5 deste informativo (amostra pequena).")
line("Referência: mediana de capacitação em turma (contratos ≥ R$ 10 mil)", ArrayFormula("C0", f"=MEDIAN(IF(({Fc}={CAP})*({Ic}>=10000),{Ic}))"), BRL, "Exclui inscrições individuais.")
# corrige referência da fórmula matricial acima
last_ref = f"C{r-1}"; inf[last_ref].value = ArrayFormula(last_ref, f"=MEDIAN(IF(({Fc}={CAP})*({Ic}>=10000),{Ic}))")
r += 1

# 7. notas
r = title(r, "7. Como ler e usar")
notas = [
    "Tema 309/STF: a justificativa de preço em inexigibilidade deve considerar o valor médio cobrado pelo próprio escritório em casos similares; estes números de mercado complementam, não substituem, essa referência.",
    "Use a faixa entre o 1º e o 3º quartil como intervalo defensável perante Tribunais de Contas; valores acima do 3º quartil pedem escopo e entregas bem descritos.",
    "Amostras pequenas (poucos contratos no recorte) oscilam muito: confira a coluna 'Contratos com valor' antes de comparar.",
    "Contratos por êxito aparecem no PNCP com valor simbólico (R$ 0,20 a R$ 1) e ficam fora das estatísticas.",
    f"As fórmulas leem as linhas 5 a {LAST} da aba Levantamento PNCP; ao incluir novas contratações, estenda esse intervalo (ou peça a atualização).",
]
for t in notas:
    c = inf.cell(r, 2, "• " + t); c.font = ar; c.alignment = Alignment(wrap_text=True, vertical="top")
    inf.merge_cells(start_row=r, start_column=2, end_row=r, end_column=9); inf.row_dimensions[r].height = 28; r += 1

# gráfico: mediana por porte
ch = BarChart(); ch.type = "bar"; ch.style = 10
ch.title = "Mediana de serviços técnicos por porte (R$)"; ch.legend = None
ch.add_data(Reference(inf, min_col=6, min_row=porte0, max_row=porte1), titles_from_data=False)
ch.set_categories(Reference(inf, min_col=2, min_row=porte0, max_row=porte1))
ch.series[0].graphicalProperties.solidFill = NAVY
ch.dataLabels = DataLabelList(); ch.dataLabels.showVal = True; ch.dataLabels.showSerName = False; ch.dataLabels.showCatName = False; ch.dataLabels.showLegendKey = False; ch.dataLabels.showPercent = False; ch.dataLabels.numFmt = '"R$" #,##0'
ch.y_axis.majorGridlines = None
ch.y_axis.numFmt = '"R$" #,##0'; ch.x_axis.scaling.orientation = "maxMin"
ch.roundedCorners = False; ch.title.overlay = False; ch.x_axis.delete = False; ch.y_axis.delete = False
ch.width, ch.height = 16, 7.5
inf.add_chart(ch, "K5")

inf.freeze_panes = "A4"
inf.sheet_properties.tabColor = NAVY
inf.page_setup.orientation = "landscape"; inf.page_setup.fitToWidth = 1; inf.page_setup.fitToHeight = 0
inf.sheet_properties.pageSetUpPr.fitToPage = True
wb.active = 1
wb.calculation.fullCalcOnLoad = True
wb.save(F)
print("ok", LAST, "sem pop:", miss)
