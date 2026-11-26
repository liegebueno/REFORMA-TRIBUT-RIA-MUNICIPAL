"""
Preenche a coluna "Valor (R$)" da planilha do levantamento consultando a
API pública do PNCP (https://pncp.gov.br/api/consulta) a partir do
"Nº Controle PNCP" de cada linha.

Uso (Python 3.8+, sem dependências externas):
    python preencher_valores_pncp.py "Levantamento_PNCP_-_Reforma_Tributaria_Municipios.xlsx"

Gera, na mesma pasta:
  • <nome>_com_valores.xlsx  – planilha com a coluna I preenchida
  • <nome>_relatorio.csv     – linha a linha: valor e origem (homologado/estimado/sem valor)

Regra: usa o valor total HOMOLOGADO; se ainda não houver, usa o ESTIMADO.
O arquivo é editado direto no XML, então gráficos, formatação, links e
fórmulas da aba Resumo são preservados (o Excel recalcula ao abrir).
"""
import csv
import json
import re
import sys
import time
import urllib.request
import zipfile
from pathlib import Path

API = "https://pncp.gov.br/api/consulta/v1/orgaos/{cnpj}/compras/{ano}/{seq}"
CTRL = re.compile(r"^\s*(\d{14})-\d+-(\d+)/(\d{4})\s*$")
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def consultar(cnpj, ano, seq, tentativas=3):
    url = API.format(cnpj=cnpj, ano=ano, seq=seq)
    erro = None
    for i in range(tentativas):
        try:
            req = urllib.request.Request(url, headers={"Accept": "application/json",
                                                       "User-Agent": "levantamento-pncp/1.0"})
            with urllib.request.urlopen(req, timeout=30) as r:
                return json.loads(r.read().decode("utf-8"))
        except Exception as e:
            erro = e
            time.sleep(2 * (i + 1))
    raise erro


def ler_linhas(z):
    """Retorna {linha: (municipio, nº controle)} da aba 'Levantamento PNCP' (sheet1)."""
    import xml.etree.ElementTree as ET
    ss = []
    if "xl/sharedStrings.xml" in z.namelist():
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
        for si in root.findall("m:si", NS):
            ss.append("".join(t.text or "" for t in si.iter(f"{{{NS['m']}}}t")))
    root = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
    linhas = {}
    for row in root.find("m:sheetData", NS).findall("m:row", NS):
        r = int(row.get("r"))
        vals = {}
        for c in row.findall("m:c", NS):
            col = re.match(r"[A-Z]+", c.get("r")).group(0)
            v = c.find("m:v", NS)
            if c.get("t") == "s" and v is not None:
                vals[col] = ss[int(v.text)]
            elif c.get("t") == "inlineStr":
                vals[col] = "".join(t.text or "" for t in c.iter(f"{{{NS['m']}}}t"))
            elif v is not None:
                vals[col] = v.text
        if r >= 5 and vals.get("J"):
            linhas[r] = (vals.get("C", ""), vals["J"])
    return linhas


def gravar_valor(xml, r, valor):
    pad = re.compile(rf'<c r="I{r}"( s="\d+")?(?: t="\w+")?\s*(?:/>|>.*?</c>)', re.S)
    m = pad.search(xml)
    novo = f'<c r="I{r}"{m.group(1) or ""}><v>{valor}</v></c>'
    return xml[:m.start()] + novo + xml[m.end():]


def main(caminho):
    origem = Path(caminho)
    with zipfile.ZipFile(origem) as z:
        linhas = ler_linhas(z)
        sheet = z.read("xl/worksheets/sheet1.xml").decode("utf-8")

    relatorio, cont = [], {"homologado": 0, "estimado": 0, "sem valor": 0, "falha": 0}
    for r, (mun, ctrl) in sorted(linhas.items()):
        m = CTRL.match(ctrl)
        if not m:
            relatorio.append([r, mun, ctrl, "", "falha", "nº de controle inválido"]); cont["falha"] += 1
            continue
        cnpj, seq, ano = m.group(1), int(m.group(2)), m.group(3)
        try:
            d = consultar(cnpj, ano, seq)
        except Exception as e:
            print(f"Linha {r} ({mun}): falha — {e}")
            relatorio.append([r, mun, ctrl, "", "falha", str(e)]); cont["falha"] += 1
            continue
        hom, est = d.get("valorTotalHomologado"), d.get("valorTotalEstimado")
        if hom:
            valor, origem_v = float(hom), "homologado"
        elif est:
            valor, origem_v = float(est), "estimado"
        else:
            relatorio.append([r, mun, ctrl, "", "sem valor", "PNCP não informa valor"]); cont["sem valor"] += 1
            continue
        sheet = gravar_valor(sheet, r, repr(valor))
        cont[origem_v] += 1
        relatorio.append([r, mun, ctrl, f"{valor:.2f}".replace(".", ","), origem_v, ""])
        print(f"Linha {r} ({mun}): R$ {valor:,.2f} [{origem_v}]")
        time.sleep(0.3)

    destino = origem.with_name(origem.stem + "_com_valores.xlsx")
    with zipfile.ZipFile(origem) as zin, zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            dados = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                dados = sheet.encode("utf-8")
            elif item.filename == "xl/workbook.xml":
                t = dados.decode("utf-8")
                if "fullCalcOnLoad" not in t:
                    t = re.sub(r"<calcPr\s*/>", '<calcPr fullCalcOnLoad="1"/>', t)
                    t = re.sub(r"<calcPr(?![^>]*fullCalcOnLoad)", '<calcPr fullCalcOnLoad="1"', t, count=1)
                dados = t.encode("utf-8")
            zout.writestr(item, dados)

    rel = origem.with_name(origem.stem + "_relatorio.csv")
    with open(rel, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Linha", "Município", "Nº Controle PNCP", "Valor (R$)", "Origem", "Observação"])
        w.writerows(relatorio)

    print(f"\nHomologados: {cont['homologado']} | Estimados: {cont['estimado']} | "
          f"Sem valor: {cont['sem valor']} | Falhas: {cont['falha']}")
    print(f"Planilha: {destino}\nRelatório: {rel}")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(1)
    main(sys.argv[1])
