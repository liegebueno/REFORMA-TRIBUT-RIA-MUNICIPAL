import json, re, sys, unicodedata
from copy import copy
from collections import Counter
import openpyxl
from openpyxl.chart import BarChart, Reference

F = sys.argv[1]
C_, L_, J_, T_ = ("Consultoria / adequação tributária", "Adequação legislativa (Código Tributário)",
                  "Assessoria jurídica tributária", "Capacitação / treinamento")
norm = lambda s: "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn").upper()

novos = json.load(open("novos2.json", encoding="utf-8"))
drop2 = json.load(open("drop2.json", encoding="utf-8"))

DROP = {("Sagres","2026-11-18"),("Lucas do Rio Verde","2026-10-09"),("Tianguá","2026-05-27"),("Cezarina","2026-05-08"),
        ("Urubici","2025-10-09"),("Campo Formoso","2026-07-31"),("Nazaré da Mata","2026-03-05"),("Angra dos Reis","2026-02-03"),
        ("Abdon Batista","2026-01-30"),("Parauapebas","2025-08-29"),("Sabáudia","2024-08-12"),("Cerejeiras","2026-01-28"),
        ("Novo Brasil","2026-04-08"),("Nova Ramada","2026-08-26")}
TOT = {("Pinheiral","2026-07-22"),("Luiziana","2026-06-30"),("Luiziana","2026-01-20"),("Luiziana","2024-01-23"),
       ("Ubaíra","2026-04-08"),("Ubaíra","2026-01-28"),("Cacaulândia","2025-10-15"),("Meridiano","2026-08-14"),
       ("Eunápolis","2026-01-30"),("Porangatu","2024-11-22"),("Itaara","2026-05-22"),("Palmares do Sul","2026-05-07"),
       ("São Gabriel","2025-08-14"),("Nova Hartz","2025-08-07"),("Nova Santa Rita","2025-07-25")}
RESC = {("Teixeirópolis","2026-10-27"):C_,("Mirandópolis","2026-10-20"):C_,("Arapongas","2026-10-09"):T_,("Castro","2026-08-31"):T_,
        ("Lindoeste","2026-08-10"):C_,("Santa Cruz da Baixa Verde","2026-06-16"):J_,("Pinheiral","2026-04-17"):T_,
        ("Restinga","2026-03-06"):T_,("Bento Gonçalves","2025-11-25"):T_,("Mundo Novo","2025-11-12"):T_,("Jacarezinho","2025-10-20"):T_,
        ("Borebi","2025-08-25"):C_,("Tijucas","2025-04-10"):L_,("Engenheiro Beltrão","2026-07-24"):L_,("Engenheiro Beltrão","2025-12-10"):L_,
        ("Aragarças","2026-03-10"):C_,("Novo Progresso","2025-11-13"):C_,("Euclides da Cunha","2026-01-28"):L_,("Porto Alegre","2025-10-21"):T_}
LOGI = re.compile(r"CAFE DA MANHA|COFFEE|LANCHE|SALGADO|HOSPEDAGEM|PASSAGE|\bVAN\b|VEICULO|PORTARIA MDR|LOCACAO DO SALAO")

def fix(s):
    if "Ã" in s:
        try: s = s.encode("latin1").decode("utf-8")
        except Exception: pass
    return s.replace("&#8220;", "“").replace("&#8221;", "”").replace("&quot;", '"').replace("¿", '"')

sel = []
for x in novos:
    k = (x["mun"], x["pub"])
    if k in DROP or LOGI.search(norm(x["desc"])): continue
    if k in TOT: x["cat"] = T_
    sel.append(x)
for _, x in drop2:
    k = (x["mun"], x["pub"])
    if k in RESC: sel.append(dict(x, cat=RESC[k]))
sel.sort(key=lambda x: (x["pub"], x["mun"]))

def objeto(desc):
    s = " ".join(fix(desc).split())
    s = re.sub(r"^\[[^\]]+\]\s*-\s*", "", s)
    s = re.sub(r"^(\d[\d./ -]*-\s*)", "", s)
    letters = [c for c in s if c.isalpha()]
    if letters and sum(c.isupper() for c in letters) / len(letters) > 0.6:
        s = s.lower(); s = s[:1].upper() + s[1:]
    return s if len(s) <= 160 else s[:157].rstrip(" ,;-") + "…"

def ente(org):
    o = norm(org)
    if "CAMARA" in o: return "Câmara Municipal"
    if any(k in o for k in ("AUTONOMO", "INSTITUTO", "FUNDACAO", "FACULDADE", "AUTARQUIA")): return "Autarquia/fundação municipal"
    if "EMPRESA" in o: return "Empresa pública municipal"
    return "Prefeitura"

MOD = {"Pregão - Eletrônico": "Pregão Eletrônico", "Concorrência - Eletrônica": "Concorrência Eletrônica",
       "Pregão - Presencial": "Pregão Presencial", "Concorrência - Presencial": "Concorrência Presencial"}
def valor(h, e):
    for v in (h, e):
        if v is not None and 0 < v < 1e9: return round(float(v), 2)

wb = openpyxl.load_workbook(F)
ws = wb["Levantamento PNCP"]
last_old = ws.max_row
have = {ws.cell(r, 10).value for r in range(5, last_old + 1)}
n_old = last_old - 4
tmpl = [ws.cell(last_old, c) for c in range(1, 12)]
r = last_old; added = []
for x in sel:
    if x["c"] in have: continue
    have.add(x["c"])
    m = re.match(r"(\d{14})-\d+-(\d+)/(\d{4})", x["c"])
    r += 1
    vals = [float(r - 4), ente(x["org"]), fix(x["mun"]), x["uf"], objeto(x["desc"]), x["cat"], MOD.get(x["mod"], x["mod"]),
            x["pub"], valor(x["hom"], x["est"]), x["c"], "Abrir no PNCP"]
    for c, v in enumerate(vals, 1):
        cell = ws.cell(r, c, v); cell._style = copy(tmpl[c - 1]._style)
    ws.cell(r, 11).hyperlink = f"https://pncp.gov.br/app/editais/{m.group(1)}/{m.group(3)}/{int(m.group(2))}"
    added.append(x)
last = r
ws.auto_filter.ref = f"A4:K{last}"
ws["A2"] = ("Fonte: Portal Nacional de Contratações Públicas (PNCP) — pncp.gov.br. Itens 1–93: consulta inicial (21/09/2026). "
            f"Itens 94–{n_old}: atualização de 26/11/2026 (publicações de 21/09 a 26/11/2026, triagem manual). "
            f"Itens {n_old+1}–{last-4}: varredura completa do PNCP em 26/11/2026 (publicações desde 01/01/2024, esfera municipal), "
            "restrita a objetos que citam expressamente a Reforma Tributária (EC 132/2023, LC 214/2025, IBS/CBS) ou o Código/legislação tributária municipal; "
            "classificação automática por palavras-chave e objeto = descrição do PNCP resumida. Exclui software, cadastro, serviços bancários, recuperação de créditos e VAF/ICMS. "
            "Valor (R$) = total homologado ou, sem homologação, o estimado (inclui valores simbólicos de contratos por êxito); em branco = não informado/zerado no PNCP.")
ws.row_dimensions[2].height = 75

# ---------- Resumo ----------
rs = wb["Resumo"]
lab = {rs.cell(i, 1).value: i for i in range(1, rs.max_row + 1) if rs.cell(i, 1).value}
rm = lab["Por modalidade de contratação"]; rv = lab["Valores contratados"]
st = {"sec": rm, "hdr": rm + 1, "item": rm + 2, "tot": lab["Total"] if False else None}
tot_rows = [i for i in range(1, rs.max_row + 1) if rs.cell(i, 1).value == "Total"]
st = {k: [copy(rs.cell(row, c)._style) for c in range(1, 4)] for k, row in
      {"sec": rm, "hdr": rm + 1, "item": rm + 2, "tot": tot_rows[1], "vsec": rv, "vhdr": rv + 1, "vitem": rv + 2, "vnote": rv + 8}.items()}
vrows = [[rs.cell(i, 1).value, rs.cell(i, 2).value] for i in range(rv + 2, rv + 8)]
note = rs.cell(rv + 8, 1).value
blank = copy(rs.cell(3, 1)._style)
for row in rs.iter_rows(min_row=6, max_row=9, min_col=2, max_col=3):
    for cell in row: cell.value = re.sub(r"\$(\d+)\b(?=[,)])", lambda m: f"${last}" if int(m.group(1)) > 9 else m.group(0), cell.value)
for i in range(rm, rs.max_row + 1):
    for c in range(1, 4):
        rs.cell(i, c).value = None; rs.cell(i, c)._style = copy(blank)
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
rows = [[ws.cell(i, c).value for c in range(1, 11)] for i in range(5, last + 1)]
cm = Counter(x[6] for x in rows)
mods = [m for m, _ in cm.most_common() if m != "Não informado"] + (["Não informado"] if "Não informado" in cm else [])
ufs = sorted({x[3] for x in rows})
m0, mt = table(rm, "Por modalidade de contratação", "G", mods)
u0, ut = table(mt + 2, "Por UF", "D", ufs)
v = ut + 3
put(v, "vsec", ["Valores contratados", None, None]); put(v + 1, "vhdr", ["Indicador", "Valor", None])
for k, (a, f) in enumerate(vrows):
    put(v + 2 + k, "vitem", [a, f.replace(f"${last_old}", f"${last}"), None])
put(v + 8, "vnote", [note, None, None])
cnt = len(rows)
rs["D5"] = f"• Amostra de {cnt} contratações municipais no PNCP desde 2024 (varredura completa em 26/11/2026 para reforma tributária e Código Tributário; consultoria tributária genérica só nos itens 1–{n_old})."
rs["D8"] = f"• Forte pulverização geográfica: {len(ufs)} UFs (" + ", ".join(ufs) + ")."
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
wb.calculation.fullCalcOnLoad = True
wb.save(F)
print(json.dumps({"adicionados": len(added), "ultima": last, "total": cnt, "cat": Counter(x["cat"] for x in added),
                  "anos": Counter(x["pub"][:4] for x in added), "mods": mods, "nufs": len(ufs),
                  "com_valor": sum(1 for x in added if valor(x["hom"], x["est"]) is not None)}, ensure_ascii=False, default=str))
