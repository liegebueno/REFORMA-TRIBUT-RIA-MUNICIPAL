import re,unicodedata
def norm(s): return "".join(c for c in unicodedata.normalize("NFD",s) if unicodedata.category(c)!="Mn").upper()
RF=re.compile(r"REFORMA TRIBUT|REFORMA DA TRIBUTACAO|REFORMA DO CONSUMO|TRIBUTACAO (SOBRE O|DO) CONSUMO|\bIBS\b|\bCBS\b|IVA[ -]?DUAL|\bIVA\b|IMPOSTO SOBRE BENS E SERVICOS|CONTRIBUICAO SOBRE BENS E SERVICOS|IMPOSTO SELETIVO|COMITE GESTOR|EMENDA CONSTITUCIONAL\s*(N[O.º°]*\s*)?132|\bEC\s*(N[O.º°]*\s*)?132|LEI COMPLEMENTAR\s*(FEDERAL\s*)?(N[O.º°]*\s*)?214|\bLC\s*(N[O.º°]*\s*)?214|\bPLP?\s*(N[O.º°]*\s*)?68|NOVO SISTEMA TRIBUT|TRANSICAO TRIBUT|NFAG|NOTA FISCAL DE AGUA")
def is_rf(s): return bool(RF.search(norm(s or "")))
