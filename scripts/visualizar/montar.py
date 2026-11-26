import re, sys, json, html
from pathlib import Path

base = Path(sys.argv[1]); out_dir = Path(sys.argv[2])
d3 = (base / "standalone/d3.min.js").read_text(encoding="utf-8")
shim = (base / "standalone/shim.js").read_text(encoding="utf-8")

TOKENS = """
:root{
  --color-fg:#1c2430; --color-fg-muted:#5b6574; --color-bg:#ffffff; --color-panel:#f5f7fa; --color-border-line:#dfe4ea;
  --color-ok:#1f8a5b; --color-warn:#c27c0e; --color-bad:#c2410c;
  --font-anthropic-sans:"Segoe UI",system-ui,-apple-system,"Helvetica Neue",Arial,sans-serif;
  --font-anthropic-serif:Georgia,"Times New Roman",serif;
  --cds-font-size-caption:12px; --cds-font-size-body:14px; --cds-font-size-heading:18px; --cds-font-size-title:30px; --cds-font-weight-medium:600;
  --cds-gap-xs:4px; --cds-gap-sm:8px; --cds-gap-md:14px; --cds-gap-lg:22px; --cds-gap-xl:36px;
  --cds-pad-xs:4px; --cds-pad-sm:8px; --cds-pad-md:12px; --cds-pad-lg:16px; --cds-pad-xl:24px; --cds-radius:10px;
  --cds-chart-grid:#e6e9ee; --cds-chart-axis:#c9ced6; --cds-chart-muted:#b8c0cc; --cds-chart-reference:#8a94a3; --cds-text-secondary:#5b6574;
  --cds-surface-popover:#ffffff; --cds-border:#dfe4ea; --cds-shadow-popover:0 6px 20px rgba(20,30,50,.14);
  --c0:#1f4e8c; --c1:#2e8a6e; --c2:#c27c0e; --c3:#8a4fb3; --c4:#c2410c; --c5:#4c7a9e;
  color-scheme: light;
}
@media (prefers-color-scheme: dark){ :root{
  --color-fg:#e7ebf0; --color-fg-muted:#a3adba; --color-bg:#12161c; --color-panel:#1a2029; --color-border-line:#2c3440;
  --cds-chart-grid:#28303b; --cds-chart-axis:#3a4452; --cds-chart-muted:#4f5a69; --cds-text-secondary:#a3adba;
  --cds-surface-popover:#1f2630; --cds-border:#2c3440; --cds-shadow-popover:0 6px 20px rgba(0,0,0,.5);
  --c0:#6ea3e6; --c1:#5cc4a0; --c2:#e3a84a; --c3:#b88ad9; --c4:#ef7a4f; --c5:#86b1d3;
  color-scheme: dark; } }
html,body{margin:0;background:var(--color-bg);color:var(--color-fg);font-family:var(--font-anthropic-sans);font-size:14px}
#dash-root{max-width:1120px;margin:0 auto;padding:28px max(16px,env(safe-area-inset-left,0px)) 40px max(16px,env(safe-area-inset-right,0px))}
.dash-skeleton{color:transparent}
.sa-foot{margin-top:28px;padding-top:12px;border-top:1px solid var(--color-border-line);font-size:12px;color:var(--color-fg-muted)}
"""


def build(page, datasets, title, out_name):
    src = (base / page).read_text(encoding="utf-8")
    src = re.sub(r'<meta name="dash-[^"]+"[^>]*>\s*', "", src)
    m = re.search(r"<script>(.*)</script>\s*$", src, re.S)
    page_js = m.group(1); body = src[: m.start()]
    data = {k: (base / v).read_text(encoding="utf-8") for k, v in datasets.items()}
    data_js = "window.__DATASETS=" + json.dumps(data, ensure_ascii=False).replace("</", "<\\/") + ";"
    foot = ('<div class="sa-foot">Versão independente gerada em 26/11/2026 a partir do painel publicado no Claude. '
            'Dados fixos daquela data (PNCP, Tesouro Transparente e IBGE). Uso interno — Monteiro e Monteiro Advogados Associados.</div>')
    doc = f"""<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>{html.escape(title)}</title>
<style>{TOKENS}</style>
<script>{d3}</script>
</head>
<body>
<div id="dash-root">
{body}
{foot}
</div>
<script>{data_js}</script>
<script>{shim}</script>
<script>{page_js}</script>
<script>window.__dashStart();</script>
</body>
</html>
"""
    (out_dir / out_name).write_text(doc, encoding="utf-8")
    print(out_name, len(doc))


build("standalone/mercado_atual.html", {"contratacoes": "dash/contratacoes.csv", "ibge_municipios": "dash/ibge_municipios.csv"},
      "Reforma Tributária nos Municípios — Mercado", "painel_mercado.html")
build("inf/index.html", {"contratacoes": "inf/contratacoes.csv", "fpm_rosario": "inf/fpm_rosario.csv"},
      "Informativo de Valores — Reforma Tributária", "informativo_de_valores.html")
