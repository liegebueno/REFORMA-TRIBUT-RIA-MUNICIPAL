# Reforma Tributária nos Municípios — mercado e propostas

Projeto interno da Monteiro e Monteiro Advogados Associados: levantamento das contratações municipais ligadas à Reforma Tributária (EC 132/2023 e LC 214/2025) publicadas no PNCP, análise de preços praticados e propostas timbradas de assessoria.

Repositório privado. Contém dados de precificação e propostas comerciais: não tornar público.

## Conteúdo

| Pasta | O que tem |
|---|---|
| `dados/` | Planilha do levantamento (versão atual e a original de 21/09/2026) e as bases em CSV usadas nos painéis |
| `painel/` | Código das páginas publicadas no Claude (painel de mercado e Informativo de Valores) |
| `propostas/` | Propostas timbradas de Reforma Tributária (Rosário do Ivaí/PR e modelo para consórcio), minutas originais e o timbre |
| `scripts/` | Scripts Python usados para coletar, filtrar e gerar os arquivos |

## Planilha (`dados/Levantamento_PNCP_-_Reforma_Tributaria_Municipios_atualizado_2026-11-26.xlsx`)

- **Levantamento PNCP**: 1.187 contratações municipais (01/01/2024 a 26/11/2026) cujo objeto no PNCP cita expressamente a Reforma Tributária (reforma tributária, EC 132/2023, LC 214/2025, IBS/CBS, IVA dual, Comitê Gestor, imposto seletivo, NFAg). Valor = homologado ou, sem homologação, estimado. Coluna L = população IBGE (SICONFI 2026).
- **Informativo de Valores**: faixas de preço por tipo de contrato, porte do município, modalidade, Paraná x Brasil e o caso Rosário do Ivaí. Tudo por fórmula (as fórmulas leem as linhas 5 a 1191).
- **Comparativo UF**: municípios que contrataram, penetração sobre o total do IBGE, leads (municípios que só fizeram curso) e valores médios por estado.
- **Resumo**: contagens e valores por categoria, modalidade e UF.

Contratações tributárias genéricas, sem menção à reforma, foram excluídas em 26/11/2026.

## Painéis publicados

- Painel de mercado e modelo de negócio: https://claude.ai/artifact/PfAVkdWvH3YUadmW6DDiMx
- Informativo de Valores: https://claude.ai/artifact/E8g9VD6QUGf9roNXGhtKwo

Os arquivos em `painel/` são o código dessas páginas. Eles rodam dentro do tipo "Data" dos artefatos do Claude (usam a API `dash` e `d3` fornecidas pela página) e não abrem sozinhos no navegador.

## Propostas

- `PROPOSTA TIMBRADA - REFORMA TRIBUTÁRIA _ ROSÁRIO DO IVAÍ` (.docx e .pdf): honorários de 0,5% do FPM anual (projeto), 0,4% do repasse mensal do FPM (acompanhamento) e R$ 15.000 por turma de capacitação. FPM de out/2025 a set/2026: R$ 16.161.880,74 (Tesouro Transparente).
- `MODELO - PROPOSTA TIMBRADA - REFORMA TRIBUTÁRIA _ CONSÓRCIO.docx`: campos do consórcio e honorários em amarelo, a preencher.
- Pendências: assinatura e juntada dos DOCs 01 (atuação perante entes e associações) e 02 (atestados).

O timbre (`propostas/timbre/timbre.png`) foi extraído das propostas timbradas em PDF e entra como imagem de fundo do cabeçalho.

## Scripts

Escritos para Python 3.12 com `openpyxl`, `python-docx`, `pypdfium2` e `Pillow`. Os caminhos dentro deles apontam para as pastas usadas na sessão de trabalho; ajuste antes de rodar.

| Script | Função |
|---|---|
| `01_atualizacao_inicial.py` | Primeira atualização da planilha com novas contratações (set–nov/2026) |
| `02_classifica_varredura.py` | Classifica a varredura completa do PNCP por palavras-chave |
| `03_inclui_varredura.py` | Inclui os itens da varredura na planilha e refaz o Resumo |
| `04_filtra_so_reforma.py` | Mantém só contratações cujo objeto cita a reforma |
| `rf.py` | Expressão regular que define "cita a reforma" |
| `aba_informativo.py` | Cria a aba Informativo de Valores e a coluna de população |
| `gerar_propostas.py` | Gera as propostas timbradas (.docx) a partir das minutas |
| `preencher_valores_pncp.py` | Script original para preencher valores pela API de consulta do PNCP |

## Fontes

- PNCP — Portal Nacional de Contratações Públicas (API de busca e de consulta), consultado em 26/11/2026.
- Tesouro Transparente — Transferências Constitucionais para Municípios (FPM).
- SICONFI/Tesouro Nacional — população dos municípios (exercício 2026).
- IBGE — número de municípios por UF.
