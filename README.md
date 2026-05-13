# kinea-pptx-plus

> **Status:** MVP complete — 3 chart types, 79 tests, all green.
> Build on `python-pptx-ng` with direct OOXML manipulation + overlay shapes.

Extensão de `python-pptx` (sobre o fork mantido `python-pptx-ng`) para reproduzir gráficos no nível ThinkCell — manipulação XML direta + overlay de shapes — open-source, MIT.

## MVP — 3 gráficos

| Chart type | Description | Strategy |
|---|---|---|
| **line_with_end_label** | Linhas históricas com label colorida ao final de cada série (estilo Andes vs CDI) | XML + overlay label_box |
| **line_area_between** | Sombreamento entre 2 linhas com suporte a cruzamento | overlay polygon_freeform |
| **line_stacked_combo** | Linha + colunas empilhadas + annotation lateral | XML combo + overlay annotation |

## Quick start

```bash
pip install -e .
pip install -e ".[dev]"   # development dependencies

# Generate an Andes vs CDI chart
python examples/andes_vs_cdi.py

# Run tests
pytest -xvs tests/regression/
```

## API

### High-level (SQL → Chart)

```python
from kinea_pptx_plus.registry.sql_to_chart import chart_from_data

data = [("2024-Q1", 42.0, 38.0), ("2024-Q2", 55.0, 50.0)]
chart_from_data(data, chart_type="line_with_end_label",
                colors=["1F3864", "C0392B"],
                output_path="output.pptx",
                title="Andes vs CDI")
```

### Low-level (per chart function)

```python
from kinea_pptx_plus.charts.line_with_end_label import build_line_with_end_label
from kinea_pptx_plus.charts.line_area_between import build_line_area_between
from kinea_pptx_plus.charts.line_stacked_combo import build_line_stacked_combo

# line_with_end_label — series colours + end labels
build_line_with_end_label("template.pptx", colors=["1F3864", "C0392B"],
                          output_path="output.pptx")

# line_area_between — shaded area between 2 series
build_line_area_between("template.pptx", fill_color="4472C4",
                        fill_alpha=40000, output_path="output.pptx")

# line_stacked_combo — stacked bar + line + annotation
build_line_stacked_combo("template.pptx", output_path="output.pptx",
                         annotation_text="Peak", annotation_target_cat=3)
```

### Overlays (individual shapes)

```python
from kinea_pptx_plus.overlays.label_box import add_label_box
from kinea_pptx_plus.overlays.polygon_freeform import add_freeform_polygon
from kinea_pptx_plus.overlays.column_annotation import add_column_annotation

# Coloured label box with auto WCAG contrast
add_label_box(slide, x_emu=5000000, y_emu=3000000,
              cx_emu=1200000, cy_emu=400000,
              fill_hex="1F3864", text="42.5%")

# Freeform polygon (area highlight)
add_freeform_polygon(slide, vertices=[(x1,y1), (x2,y2), ...],
                     fill_color="4472C4", fill_alpha=40000)

# Annotation with leader line
add_column_annotation(slide, x_emu=5000000, y_emu=3000000,
                      cx_emu=2000000, cy_emu=500000,
                      text="Key insight", target_x=6000000, target_y=4000000)
```

## Architecture

```
src/kinea_pptx_plus/
├── charts/                    # Chart-type implementations
│   ├── line_with_end_label.py # Line chart + end labels
│   ├── line_area_between.py   # Shaded area between lines
│   └── line_stacked_combo.py  # Stacked bar + line
├── xml_primitives/            # Low-level OOXML manipulation
│   ├── nsmap.py               # All OOXML namespaces
│   ├── element_factory.py     # make_solidFill, make_ln, make_txPr
│   ├── color.py               # Color from_hex, scheme, tint/shade/alpha
│   └── tree_ops.py            # insert_sorted, replace_child, deep_clone
├── overlays/                  # Shape overlays
│   ├── label_box.py           # Coloured rect with WCAG contrast
│   ├── polygon_freeform.py    # Freeform polygon with custGeom
│   ├── column_annotation.py   # Text box + leader line connector
│   └── positioner.py          # EMU ↔ data coordinate conversion
├── validation/                # Multi-gate validation pipeline
│   ├── xml_schema.py          # Parsing gate
│   ├── win32_probe.py         # PowerPoint COM open/close
│   ├── render_check.py        # LibreOffice render + OCR
│   └── golden_compare.py      # Byte-level XML comparison
├── registry/                  # Chart detection + SQL mapping
│   ├── chart_detector.py      # OCR + heuristic classification
│   └── sql_to_chart.py        # Data/chart-from-data → PPTX
├── io.py                      # extract_chart_xml
└── _agent/                    # Agent loop, memory, BLOCKED tracking
```

## Validation gates

For every chart produced, 4 gates are available:

1. **Schema gate** — chart XML parses without error
2. **Roundtrip gate** — save → reload → save produces identical XML
3. **Win32 probe** — opens PPTX in PowerPoint via COM (Windows only)
4. **Render-OCR** — renders to PNG via LibreOffice, checks text via Tesseract

## Data: 5-year synthetic fixture

The repo includes `tests/fixtures/andes_vs_cdi_5yr.pptx` with 70 monthly data points (Jan 2020 - Oct 2025) generated by `scripts/synth_andes_vs_cdi.py`.

## Development

```bash
# Run full test suite
pytest -xvs tests/regression/

# Install with dev dependencies
pip install -e ".[dev]"
pre-commit install
```

### Task tracking

- `spec/CURRENT_TASK.md` — current task
- `spec/BACKLOG.md` — full backlog with status
- `spec/DONE/` — completed task specs
- `src/kinea_pptx_plus/_agent/memory.md` — architecture decisions & lessons

## License

MIT.
