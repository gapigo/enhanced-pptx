# Agent Memory

Aqui você (o agente) escreve tudo que aprende. Esse arquivo é seu cérebro persistente entre iterações. Estrutura livre, mas tente manter as seções abaixo.

## Decisões arquiteturais

- **2026-05-12 (seed)**: Base é `python-pptx-ng`, fork mantido do scanny. Não é `python-pptx` original.
- **2026-05-12 (seed)**: 3 estratégias (XML puro / híbrido win32 / overlay) decididas pelo Gabriel. Não rediscutir.
- **2026-05-13 (line_area_between)**: OOXML nativo não suporta "area between two lines". Tentativas: combo area+line não funciona (preenche do eixo até a linha superior, não entre as linhas). areaChart com transparência + sobreposição também falha (PowerPoint empilha áreas como accumulated stacks). Solução: overlay com `a:custGeom` freeform polygon com fill alpha.
- **2026-05-13 (label_box)**: `p:sp` shapes inseridos via `slide.shapes._spTree.append(sp)` funcionam, mas perdem a numeração de ID (cNvPr id="0" em todos). PowerPoint corrige na abertura, então roundtrip gate precisa salvar → abrir → salvar → comparar XML.
- **2026-05-13 (positioner)**: `Positioner.from_chart_xml` precisa do chart XML from `extract_chart_xml`, que por sua vez extrai o chartSpace serializado. EMU calculations baseiam-se em plotArea bounding box com fallback de 10% margins.
- **2026-05-13 (5yr fixture)**: Andes vs CDI sintético com 70 pontos mensais (Jan 2020 - Out 2025) via `scripts/synth_andes_vs_cdi.py`. Testado com paint_series_colors + build_line_with_end_label — funciona.

## XML quirks descobertos

- (vazio — preencha conforme descobrir)

## python-pptx-ng quirks

- (vazio)

## OOXML armadilhas

- `c:dPt` tem que vir ORDENADO por `c:idx` ou o PowerPoint silenciosamente reordena e quebra cores.
- Namespaces: `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"` e `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`. Se confundir, c é chart, a é drawing.

## Render-OCR notas

- LibreOffice em headless: `soffice --headless --convert-to png --outdir /tmp file.pptx`. Cuidado: gera 1 PNG por slide.

## Win32 probe notas

- Só roda em Windows. Use `os.name == 'nt'` guard.
- COM dispatch: `win32com.client.Dispatch("PowerPoint.Application")`. Tem que setar `app.Visible = 0` mas mesmo assim aparece. Workaround documentado em stackoverflow: usar `DispatchEx`.

## Tasks bloqueadas

(vide BLOCKED.md)
