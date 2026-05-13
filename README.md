# kinea-pptx-plus

> **Status:** lab experimental. Em construção autônoma por agente. Estado atual em `_agent/memory.md`.

Extensão de `python-pptx` (sobre o fork mantido `python-pptx-ng`) para reproduzir gráficos no nível ThinkCell — manipulação XML direta + overlay de shapes — open-source, MIT.

## Por que existe

PowerPoint é um formato terrível pra automatizar gráficos sofisticados. ThinkCell resolve, mas é proprietário e caro. Esta lib tenta reproduzir o que ele faz, com foco em casos financeiros (séries temporais, comparativos, attribution analysis).

## MVP — 3 gráficos

1. **line_with_end_label** — linhas históricas com label colorida ao final de cada série (estilo Andes vs CDI)
2. **line_area_between** — sombreamento entre 2 linhas
3. **line_stacked_combo** — linha + colunas empilhadas + annotation lateral

## Como funciona

Cada gráfico é uma função que tenta 3 estratégias em ordem:

1. **XML puro** — edita `chart1.xml` diretamente via lxml
2. **Híbrido com validação win32** — XML puro + abre o pptx no PowerPoint COM pra confirmar que não corrompeu silenciosamente
3. **Overlay de shapes** — gráfico nativo cru + shapes posicionados em cima (labels coloridas, polígonos de área, annotations)

Decisão é interna. Usuário só chama:

```python
from kinea_pptx_plus import Studio

studio = Studio(template="kinea_master.pptx")
slide = studio.new_slide(layout="chart_full")
slide.add_chart_from_sql(
    sql="SELECT date, andes_pct, cdi_pct FROM perf.fund_vs_index",
    chart_hint="line_with_end_label",
    title="Andes vs CDI",
)
studio.save("output.pptx")
```

## Para agentes (humanos ou LLM) que vão continuar

Leia nesta ordem:

1. `CLAUDE.md` — constituição com regras absolutas
2. `docs/MASTER_PLAN.md` — plano completo
3. `spec/BACKLOG.md` — tasks ordenadas, pegue a primeira com `status: ready`
4. `_agent/memory.md` — o que aprendemos até agora
5. `_agent/BLOCKED.md` — impasses pendentes (se houver)

## Setup

```bash
pip install -e .
pre-commit install
pytest -xvs tests/regression/
```

## Stack

`python-pptx-ng` · `lxml` · `pywin32` (Windows, opcional) · `Pillow` + `pytesseract` (OCR) · `pandas` · `pytest`

## License

MIT.
