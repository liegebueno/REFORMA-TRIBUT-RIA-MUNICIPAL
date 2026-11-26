import json, re, unicodedata, sys
import openpyxl

F = sys.argv[1]
d = json.load(open("full.json", encoding="utf-8"))
ws = openpyxl.load_workbook(F)["Levantamento PNCP"]
have = {ws.cell(r, 10).value for r in range(5, ws.max_row + 1)}

def norm(s):
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn").upper()

RF = re.compile(r"REFORMA TRIBUT|\bIBS\b|\bCBS\b|EMENDA CONSTITUCIONAL (N[O.º°]*\s*)?132|EC\s*(N[O.º°]*\s*)?132|LC\s*(N[O.º°]*\s*)?214|LEI COMPLEMENTAR (FEDERAL )?(N[O.º°]*\s*)?214|IVA[ -]DUAL|TRIBUTACAO SOBRE O CONSUMO|REFORMA DO CONSUMO")
CT = re.compile(r"CODIGO TRIBUT|LEGISLACAO TRIBUT")
EVT = re.compile(r"CURSO|INSCRI[CÇ]|CONGRESSO|SEMINARIO|WORKSHOP|FORUM|ENCONTRO|PALESTRA|OFICINA|JORNADA|CAPACITA|TREINAMENTO|WEBINAR|EVENTO|LIVRO|PUBLICAC")
SRV = re.compile(r"CONSULTORIA|ASSESSORIA|ASSESSORAMENTO|ELABORACAO|REVISAO|APOIO TECNICO|SUPORTE TECNICO|ADEQUACAO|IMPLEMENTA|IMPLANTA|DIAGNOSTICO|ESTUDO")
SOFT = re.compile(r"SOFTWARE|SISTEMA|LICEN[CS]|LOCACAO|CLOUD|NUVEM|\bERP\b|NOTA FISCAL|NFS-?E|NFA?G\b")
JUR = re.compile(r"JURIDIC|ADVOCA|ADVOGAD")
EX = re.compile(r"CADASTR|GEOPROCESS|GEORREF|CARTOGRA|AEROFOTO|BANCARI|INSTITUIC[AO]ES? FINANCEIRA|RECUPERA|COMPENSA|PASEP|CFEM|\bFPM\b|\bVAF\b|\bICMS\b|ROYALT|\bCOSIP\b|DIVIDA ATIVA|PROTESTO|PLANTA GENERICA")
C_, L_, J_, T_ = ("Consultoria / adequação tributária", "Adequação legislativa (Código Tributário)",
                  "Assessoria jurídica tributária", "Capacitação / treinamento")

out, drop = [], []
for x in d["cand"]:
    if x["c"] in have or x["canc"] or x["pub"] < "2024-01-01":
        continue
    t = norm(x["desc"])
    if RF.search(t):
        srv = SRV.search(t) and not re.search(r"CURSO|INSCRI[CÇ]|CONGRESSO|SEMINARIO|WORKSHOP|FORUM|PALESTRA|WEBINAR", t)
        if srv:
            if SOFT.search(t) or EX.search(t): drop.append(("soft/ex", x)); continue
            cat = L_ if CT.search(t) else (J_ if JUR.search(t) else C_)
        elif EVT.search(t):
            cat = T_
        else:
            drop.append(("rf-outro", x)); continue
    elif CT.search(t):
        if SOFT.search(t) or EX.search(t): drop.append(("ct-soft", x)); continue
        if EVT.search(t) and not SRV.search(t): cat = T_
        else: cat = L_
    else:
        continue
    out.append(dict(x, cat=cat))

json.dump(out, open("novos2.json", "w", encoding="utf-8"), ensure_ascii=False)
json.dump([[k, x] for k, x in drop], open("drop2.json", "w", encoding="utf-8"), ensure_ascii=False)
from collections import Counter
print(len(out), Counter(o["cat"] for o in out), Counter(k for k, _ in drop))
with open("rev_srv.txt", "w", encoding="utf-8") as f:
    for o in out:
        if o["cat"] != T_:
            print(o["cat"][:5], o["pub"], o["mun"], o["uf"], "|", " ".join(o["desc"].split())[:170], file=f)
with open("rev_drop.txt", "w", encoding="utf-8") as f:
    for k, o in drop:
        print(k, o["pub"], o["mun"], "|", " ".join(o["desc"].split())[:170], file=f)
import random; random.seed(1)
with open("rev_cap.txt", "w", encoding="utf-8") as f:
    for o in [o for o in out if o["cat"] == T_]:
        print(o["pub"], o["mun"], "|", " ".join(o["desc"].split())[:140], file=f)
