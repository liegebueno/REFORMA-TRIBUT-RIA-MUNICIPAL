import json, re, sys, csv
from copy import copy
from collections import Counter
import openpyxl
from openpyxl.chart import BarChart, Reference
from rf import is_rf

F = sys.argv[1]
D = json.load(open("descs.json", encoding="utf-8"))
O = json.load(open("objetos_full.json", encoding="utf-8"))
for c, v in O.items():
    if "o" in v:
        D[c] = (v["o"] or "") + " " + (v["i"] or "")
    else:
        D[c] = ""  # removido do PNCP: não verificável

wb = openpyxl.load_workbook(F)
ws = wb["Levantamento PNCP"]
old_last = ws.max_row
tmpl = [copy(ws.cell(old_last, c)._style) for c in range(1, 12)]
rows, out = [], []
for r in range(5, old_last + 1):
    vals = [ws.cell(r, c).value for c in range(1, 12)]
    link = ws.cell(r, 11).hyperlink.target if ws.cell(r, 11).hyperlink else None
    (rows if is_rf(D.get(vals[9], "")) else out).append((vals, link))
for r in range(5, old_last + 1):
    for c in range(1, 12):
        cell = ws.cell(r, c); cell.value = None; cell.hyperlink = None
for i, (vals, link) in enumerate(rows):
    r = 5 + i
    vals[0] = float(i + 1)
    for c, v in enumerate(vals, 1):
        cell = ws.cell(r, c, v); cell._style = copy(tmpl[c - 1])
    if link: ws.cell(r, 11).hyperlink = link
blank = copy(ws.cell(old_last + 5, 1)._style)
for r in range(5 + len(rows), old_last + 1):
    for c in range(1, 12): ws.cell(r, c)._style = copy(blank)
last = 4 + len(rows)
ws.auto_filter.ref = f"A4:K{last}"
ws["A2"] = ("Fonte: Portal Nacional de Contratações Públicas (PNCP) — pncp.gov.br. Contratações de prefeituras, câmaras e autarquias municipais "
            "publicadas de 2024 a 26/11/2026 cujo objeto no PNCP cita expressamente a Reforma Tributária (reforma tributária, EC 132/2023, LC 214/2025, "
            "IBS/CBS, IVA dual, Comitê Gestor, imposto seletivo, NFAg). Contratações tributárias genéricas, sem menção à reforma, foram excluídas em 26/11/2026. "
            "Base consolidada da consulta inicial (21/09/2026), da atualização de 26/11/2026 e de varredura completa do PNCP na mesma data. "
            "Valor (R$) = total homologado ou, sem homologação, o estimado (inclui valores simbólicos de contratos por êxito); em branco = não informado/zerado no PNCP.")

# ---------- Resumo ----------
rs = wb["Resumo"]
lab = {rs.cell(i, 1).value: i for i in range(1, rs.max_row + 1) if rs.cell(i, 1).value}
rm = lab["Por modalidade de contratação"]; rv = lab["Valores contratados"]
tot_rows = [i for i in range(1, rs.max_row + 1) if rs.cell(i, 1).value == "Total"]
st = {k: [copy(rs.cell(row, c)._style) for c in range(1, 4)] for k, row in
      {"sec": rm, "hdr": rm + 1, "item": rm + 2, "tot": tot_rows[1], "vsec": rv, "vhdr": rv + 1, "vitem": rv + 2, "vnote": rv + 8}.items()}
vrows = [[rs.cell(i, 1).value, rs.cell(i, 2).value] for i in range(rv + 2, rv + 8)]
note = rs.cell(rv + 8, 1).value
bl = copy(rs.cell(3, 1)._style)
for row in rs.iter_rows(min_row=6, max_row=9, min_col=2, max_col=3):
    for cell in row: cell.value = cell.value.replace(f"${old_last}", f"${last}")
for i in range(rm, rs.max_row + 1):
    for c in range(1, 4):
        rs.cell(i, c).value = None; rs.cell(i, c)._style = copy(bl)
Lr = "'Levantamento PNCP'!"
def put(row, key, vals):
    for c, v in enumerate(vals, 1):
        cell = rs.cell(row, c, v); cell._style = copy(st[key][c - 1])
def table(start, title, col, items):
    put(start, "sec", [title, None, None]); put(start + 1, "hdr", ["Item", "Qtde", "Valor (R$)"])
    r0 = start + 2
    for k, it in enumerate(items):
        rr = r0 + k
        put(rr, "item", [it, f"=COUNTIF({Lr}${col}$5:${col}${last},A{rr})", f"=SUMIF({Lr}${col}$5:${col}${last},A{rr},{Lr}$I$5:$I${last})"])
    rt = r0 + len(items)
    put(rt, "tot", ["Total", f"=SUM(B{r0}:B{rt-1})", f"=SUM(C{r0}:C{rt-1})"])
    return r0, rt
data = [v for v, _ in rows]
cm = Counter(x[6] for x in data)
mods = [m for m, _ in cm.most_common() if m != "Não informado"] + (["Não informado"] if "Não informado" in cm else [])
ufs = sorted({x[3] for x in data})
m0, mt = table(rm, "Por modalidade de contratação", "G", mods)
u0, ut = table(mt + 2, "Por UF", "D", ufs)
v = ut + 3
put(v, "vsec", ["Valores contratados", None, None]); put(v + 1, "vhdr", ["Indicador", "Valor", None])
for k, (a, f) in enumerate(vrows):
    put(v + 2 + k, "vitem", [a, f.replace(f"${old_last}", f"${last}"), None])
put(v + 8, "vnote", [note, None, None])
rs["D5"] = f"• {len(data)} contratações municipais no PNCP desde 2024, todas com menção expressa à Reforma Tributária no objeto (varredura de 26/11/2026)."
rs["D7"] = "• Objetos: novo Código Tributário à luz da EC 132/LC 214, consultoria de transição, assessoria jurídica e, sobretudo, capacitação."
rs["D8"] = f"• Pulverização geográfica: {len(ufs)} UFs (" + ", ".join(ufs) + ")."
rs._charts = []
def bar(title, r0, r1, anchor, w, h):
    ch = BarChart(); ch.type = "col"; ch.title = title; ch.legend = None
    ch.add_data(Reference(rs, min_col=2, min_row=r0, max_row=r1), titles_from_data=False)
    ch.set_categories(Reference(rs, min_col=1, min_row=r0, max_row=r1))
    ch.series[0].graphicalProperties.solidFill = "4F81BD"
    ch.roundedCorners = False; ch.title.overlay = False; ch.x_axis.delete = False; ch.y_axis.delete = False
    ch.width, ch.height = w, h; rs.add_chart(ch, anchor)
bar("Contratações por categoria", 6, 9, f"D{rm+1}", 11.7, 6.4)
bar("Contratações por UF", u0, ut - 1, f"D{u0}", 18, 8)

# ---------- Comparativo UF ----------
cu = wb["Comparativo UF"]
for row in cu.iter_rows():
    for cell in row:
        if isinstance(cell.value, str) and cell.value.startswith("="):
            cell.value = cell.value.replace(f"${old_last}", f"${last}")

wb.calculation.fullCalcOnLoad = True
wb.save(F)

SH = {"Capacitação / treinamento": "Capacitação", "Consultoria / adequação tributária": "Consultoria",
      "Adequação legislativa (Código Tributário)": "Código Tributário", "Assessoria jurídica tributária": "Assessoria jurídica"}
with open("dash/contratacoes.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f); w.writerow(["item", "uf", "municipio", "ente", "categoria", "servico", "modalidade", "data", "ano", "valor", "controle_pncp"])
    for x in data:
        w.writerow([int(x[0]), x[3], x[2], x[1], SH[x[5]], "nao" if x[5] == "Capacitação / treinamento" else "sim", x[6], x[7], x[7][:4], "" if x[8] is None else x[8], x[9]])
with open("excluidos.txt", "w", encoding="utf-8") as f:
    for vals, _ in out: print(vals[2], vals[3], "|", vals[5], "|", vals[4][:110], file=f)
print(json.dumps({"mantidos": len(data), "excluidos": len(out), "cat": Counter(x[5] for x in data), "excl_cat": Counter(v[5] for v, _ in out),
                  "ultima": last, "ufs": len(ufs), "mods": mods}, ensure_ascii=False))
