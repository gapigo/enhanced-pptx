# CURRENT_TASK — T01: Setup base

> Quando completar essa task, mova-a para `spec/DONE/T01.md`, escreva o próximo task ID em `spec/CURRENT_TASK.md`, e copie o conteúdo de `BACKLOG.md` para o T02.

## Goal
Criar a estrutura mínima do pacote, instalar deps, ler um PPTX de fixture e devolver o XML do primeiro chartSpace.

## Steps

1. Criar `pyproject.toml`:
   ```toml
   [project]
   name = "kinea-pptx-plus"
   version = "0.0.1"
   requires-python = ">=3.11"
   dependencies = [
     "python-pptx-ng>=1.0.0",
     "lxml>=5.0",
     "pandas>=2.0",
     "Pillow>=10.0",
     "pytesseract>=0.3",
   ]

   [project.optional-dependencies]
   windows = ["pywin32>=306"]
   dev = ["pytest>=8", "pytest-xdist", "black", "ruff", "mypy", "pre-commit"]
   ```

2. Criar `src/kinea_pptx_plus/__init__.py` com `__version__ = "0.0.1"`.

3. Criar `src/kinea_pptx_plus/io.py` com função:
   ```python
   def extract_chart_xml(pptx_path: str, chart_index: int = 0) -> bytes:
       """Devolve o XML cru do chart no índice especificado.
       
       Lê o pptx via python-pptx-ng, encontra todos os GraphicFrames com has_chart,
       pega o de índice <chart_index>, retorna o XML serializado do chartSpace.
       
       Example:
           >>> xml = extract_chart_xml("fixture.pptx", 0)
           >>> b"<c:chartSpace" in xml
           True
       """
   ```

4. Criar fixture `tests/fixtures/simple_line_chart.pptx` (gere via python-pptx-ng com 2 séries).

5. Criar `tests/regression/test_T01_setup.py`:
   ```python
   def test_can_load_chart_xml():
       from kinea_pptx_plus.io import extract_chart_xml
       xml = extract_chart_xml("tests/fixtures/simple_line_chart.pptx")
       assert b"<c:chartSpace" in xml
       assert b"lineChart" in xml.lower() or b"linechart" in xml.lower()
   ```

6. `pytest -xvs tests/regression/test_T01_setup.py` → green.

7. Commit: `feat(setup): T01 base package + chart xml extraction`

## Validation
- pytest passa
- `pip install -e .` funciona limpo
- `python -c "from kinea_pptx_plus.io import extract_chart_xml; print(extract_chart_xml('tests/fixtures/simple_line_chart.pptx')[:200])"` mostra XML válido

## Output esperado
- Arquivos criados (não modifique outros): pyproject.toml, src/kinea_pptx_plus/__init__.py, src/kinea_pptx_plus/io.py, tests/fixtures/simple_line_chart.pptx, tests/regression/test_T01_setup.py
- 1 commit
- Atualizar `spec/BACKLOG.md`: T01 status → done
- Criar `spec/CURRENT_TASK.md` apontando para T02
