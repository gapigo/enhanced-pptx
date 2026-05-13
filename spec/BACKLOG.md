# BACKLOG

Tasks em ordem de dependência. Agente pega a primeira com `status: ready`. Não pula.

---

## T01 — Setup base
- **status:** ready
- **escopo:** chore
- **descrição:** Cria pyproject.toml, instala python-pptx-ng, lxml, pytest, lê PPTX de fixture e devolve um chartSpace XML em stdout.
- **DoD:** `pytest tests/regression/test_T01_setup.py::test_can_load_chart_xml` passa.

## T02 — xml_primitives.nsmap + element_factory
- **status:** ready
- **depends:** T01
- **descrição:** Implementa NSMAP completo OOXML (a, c, p, r, etc) e funções `make_solidFill(rgb)`, `make_ln(w, rgb)`, `make_txPr(...)`. Testa contra goldens.
- **DoD:** 5 elementos produzidos batem byte-a-byte com goldens criados a partir de PPTX salvo do PowerPoint real.

## T03 — Color helpers
- **status:** ready
- **depends:** T02
- **descrição:** `Color.from_hex("#1F3864") → srgbClr`, suporte a tint/shade/alpha, conversor para schemeClr.
- **DoD:** golden tests.

## T04 — Positioner (EMU ↔ data coord)
- **status:** ready
- **depends:** T02
- **descrição:** Dado um chartSpace, lê plotArea bounding box, eixo X (categorical date) e eixo Y (linear), retorna `data_to_emu(x_date, y_value)`. Lida com slide dimensions.
- **DoD:** Erro < 1% relativo entre posição calculada e posição real medida em PNG renderizado.

## T05 — line_with_end_label v0 (XML path: cores de série)
- **status:** ready
- **depends:** T03
- **descrição:** Pinta cada série da cor especificada via `c:ser/c:spPr/a:ln/a:solidFill`. Sem label ainda.
- **DoD:** schema gate + roundtrip gate + render-OCR (cores corretas via análise de pixels).

## T06 — overlays.label_box
- **status:** ready
- **depends:** T04
- **descrição:** Cria um shape rect com fill rgb, texto branco/preto auto (contraste WCAG), posicionado em (x_emu, y_emu).
- **DoD:** golden xml do shape + render check.

## T07 — line_with_end_label v1 (overlay path: labels coloridas)
- **status:** ready
- **depends:** T05, T06
- **descrição:** Junta T05 + T06. Para cada série pintada, calcula posição do último ponto e adiciona label_box com o valor formatado.
- **DoD:** todos os 4 gates passam. Compara com `tests/fixtures/andes_vs_cdi_golden.pptx`.

## T08 — line_area_between v0 (XML path: combo line + area)
- **status:** ready
- **depends:** T05
- **descrição:** Tenta primeiro via OOXML nativo (areaChart com transparência) numa cópia separada. Documenta limitações em `_agent/memory.md`.
- **DoD:** ou funciona com nativo, ou registra impossibilidade e libera T09.

## T09 — overlays.polygon_freeform
- **status:** ready
- **depends:** T04
- **descrição:** Cria `a:custGeom` com vértices arbitrários. Suporta fill com alpha. Usado para area-between.
- **DoD:** polígono triangular e polígono de 50 vértices, ambos renderizam corretamente.

## T10 — line_area_between v1 (overlay path)
- **status:** ready
- **depends:** T08, T09
- **descrição:** Calcula vértices entre 2 linhas, detecta interseções, gera 1 ou N polígonos. Cor única ou dual (acima/abaixo).
- **DoD:** 4 gates + caso de cruzamento testado.

## T11 — overlays.column_annotation
- **status:** ready
- **depends:** T06
- **descrição:** Text box com leader line (cxnSp) apontando pra coluna específica em barChart.
- **DoD:** golden + render.

## T12 — line_stacked_combo
- **status:** ready
- **depends:** T05, T11
- **descrição:** barChart empilhado + lineChart no mesmo plotArea. Annotation lateral para colunas.
- **DoD:** 4 gates.

## T13 — Validation harness completo
- **status:** ready
- **depends:** T01
- **descrição:** Pipeline `validate(pptx_path) → ValidationReport(schema, roundtrip, win32, render_ocr)`. Win32 só roda se `os.name == 'nt'`.
- **DoD:** rodar contra os 3 gráficos do MVP e contra 3 pptx broken-on-purpose (validação detecta os 3).

## T14 — chart_detector (OCR + heurística)
- **status:** ready
- **depends:** T07, T10, T12
- **descrição:** Dado PNG/JPEG, retorna `"line_with_end_label" | "line_area_between" | "line_stacked_combo"`.
- **DoD:** 90% accuracy num set de 30 imagens (a serem coletadas pelo agente da web ou geradas).

## T15 — sql_to_chart
- **status:** ready
- **depends:** T07
- **descrição:** Conecta postgres, executa query, mapeia colunas, chama chart function. Modo sintético se DB não disponível.
- **DoD:** exemplo `examples/andes_vs_cdi.py` produz pptx final.

## T16 — README.md raiz + exemplo gif
- **status:** ready
- **depends:** T15
- **descrição:** README explica setup, exemplos, arquitetura, como agente novo continua. Inclui um GIF/screenshot do output.
- **DoD:** README revisado pelo próprio agente via checklist.
