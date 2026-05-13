# Agent Memory

Aqui você (o agente) escreve tudo que aprende. Esse arquivo é seu cérebro persistente entre iterações. Estrutura livre, mas tente manter as seções abaixo.

## Decisões arquiteturais

- **2026-05-12 (seed)**: Base é `python-pptx-ng`, fork mantido do scanny. Não é `python-pptx` original.
- **2026-05-12 (seed)**: 3 estratégias (XML puro / híbrido win32 / overlay) decididas pelo Gabriel. Não rediscutir.

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
