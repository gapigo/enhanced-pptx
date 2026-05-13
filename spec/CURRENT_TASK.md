# CURRENT_TASK — T02: xml_primitives.nsmap + element_factory

> Quando completar essa task, mova-a para `spec/DONE/T02.md`, escreva o próximo task ID em `spec/CURRENT_TASK.md`, e copie o conteúdo de `BACKLOG.md` para o T03.

## Goal
Implementar NSMAP completo OOXML (a, c, p, r, etc) e funções `make_solidFill(rgb)`, `make_ln(w, rgb)`, `make_txPr(...)`. Testar contra goldens.

## Steps

1. Criar `src/kinea_pptx_plus/xml_primitives/nsmap.py`:
   ```python
   """Namespace constants for OOXML chart-related XML."""

   NSMAP = {
       "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
       "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
       "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
       "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
       "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
       "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
       "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
       "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
       "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
       "v": "urn:schemas-microsoft-com:vml",
       "o": "urn:schemas-microsoft-com:office:office",
       "14": "http://schemas.microsoft.com/office/drawing/2010/main",
   }
   ```

2. Criar `src/kinea_pptx_plus/xml_primitives/__init__.py` vazio.

3. Criar `src/kinea_pptx_plus/xml_primitives/element_factory.py` com:
   - `make_solidFill(rgb_hex: str) -> etree.Element` — cria `<a:solidFill><a:srgbClr val="RRGGBB"/></a:solidFill>`
   - `make_ln(w: int, rgb_hex: str) -> etree.Element` — cria `<a:ln w="..."><a:solidFill><a:srgbClr val="RRGGBB"/></a:solidFill></a:ln>`
   - `make_txPr(body: etree.Element | None = None) -> etree.Element` — cria `<c:txPr>` com `<a:bodyPr/>`, `<a:lstStyle/>`, e `<a:p>` opcional contendo `body`

4. Criar goldens em `tests/fixtures/xml_primitives/golden_solidFill.xml`, `golden_ln.xml`, `golden_txPr.xml` — usar XML salvo de um PPTX real do PowerPoint como referência (ou gerar via python-pptx-ng e extrair).

5. Criar `tests/regression/test_T02_xml_primitives.py`:
   ```python
   def test_make_solidFill_matches_golden():
       from kinea_pptx_plus.xml_primitives.element_factory import make_solidFill
       import lxml.etree as etree
       elem = make_solidFill("1F3864")
       golden = etree.parse("tests/fixtures/xml_primitives/golden_solidFill.xml").getroot()
       assert etree.tostring(elem) == etree.tostring(golden)

   def test_make_ln_matches_golden():
       ...

   def test_make_txPr_matches_golden():
       ...
   ```

6. `pytest -xvs tests/regression/test_T02_xml_primitives.py` → green.

7. Commit: `feat(xml-primitives): T02 nsmap + element_factory (solidFill, ln, txPr)`

## Validation
- Golden tests pass: 5 elementos batem byte-a-byte com os goldens
- `pip install -e .` segue limpo

## Output esperado
- Arquivos criados: `src/kinea_pptx_plus/xml_primitives/__init__.py`, `nsmap.py`, `element_factory.py`, `tests/regression/test_T02_xml_primitives.py`, `tests/fixtures/xml_primitives/*.xml`
- 1 commit
- Atualizar `spec/BACKLOG.md`: T02 status → done
- Criar `spec/CURRENT_TASK.md` apontando para T03
