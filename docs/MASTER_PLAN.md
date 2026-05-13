# Kinea PPTX-Plus — Plano Mestre

**Autor original:** Gabriel
**Agente executor:** DeepSeek V4 Flash (modo Karpathy Loop puro)
**Data de partida:** 12/05/2026
**Objetivo:** Estender python-pptx com manipulação XML profunda + overlay de shapes para reproduzir gráficos no nível ThinkCell, 100% open-source, com API que aceita queries SQL e produz PPTX impecável.

---

## 0. TL;DR para o agente

Você é um agente autônomo rodando madrugada adentro. Sua missão é construir, gráfico por gráfico, uma biblioteca chamada `kinea_pptx_plus` que estende `python-pptx` via manipulação XML direta + overlay de shapes. Você commita cada feature isoladamente, valida pesado, e nunca avança sem o gráfico atual passar 100% dos testes — incluindo abrir no PowerPoint via win32com sem corrupção silenciosa. Se travar, você documenta o impasse num arquivo `BLOCKED.md` e tenta outro caminho. Você não pergunta nada — você decide e commita.

---

## 1. Princípios não-negociáveis

1. **XML primeiro, sempre.** Toda manipulação tenta o caminho XML puro primeiro. Win32com é APENAS para validação (abrir e fechar o pptx pra detectar corrupção silenciosa). Overlay de shapes é APENAS para o que XML nativo não suporta.
2. **Reuso radical de python-pptx.** Não reimplementa o que já existe. Estende via subclasses, mixins e funções utilitárias. Importa lxml direto quando precisa de surgical strikes.
3. **Atomicidade de commit.** Um commit = uma feature testável. Mensagens no padrão `feat(line-end-label): adiciona label colorida ao final da série` ou `fix(area-between): corrige overlap quando séries cruzam`.
4. **Documentação por função.** Toda função pública tem docstring com: o que faz, qual elemento XML manipula, qual o efeito visual, exemplo mínimo.
5. **Idempotência.** Rodar a mesma função duas vezes no mesmo PPTX produz o mesmo resultado. Sempre. Sem duplicação de shapes, sem XML acumulado.
6. **OCR como contrato visual.** A validação final de um gráfico inclui renderizar o slide para PNG (via LibreOffice headless) e rodar OCR pra conferir que os labels esperados aparecem nas posições esperadas.

---

## 2. Arquitetura

```
kinea_pptx_plus/
├── charts/                    # Um módulo por tipo de gráfico do MVP
│   ├── line_with_end_label.py
│   ├── line_area_between.py
│   └── line_stacked_combo.py
├── xml_primitives/            # Helpers genéricos de manipulação XML
│   ├── nsmap.py               # Namespaces OOXML
│   ├── element_factory.py     # Criação de <c:ser>, <c:dPt>, <c:dLbl>, etc
│   ├── color.py               # srgbClr, schemeClr, tint/shade
│   └── tree_ops.py            # Insert ordenado, replace, deep clone
├── overlays/                  # Shapes posicionados sobre gráficos
│   ├── label_box.py           # Caixa colorida com texto (label ao fim de linha)
│   ├── column_annotation.py   # Annotation lateral
│   └── positioner.py          # Calcula posição (EMU) a partir de coordenada de dados
├── validation/
│   ├── xml_schema.py          # Valida contra XSD do OOXML
│   ├── win32_probe.py         # Abre/fecha PPTX via COM (apenas Windows)
│   ├── render_check.py        # Renderiza via LibreOffice + OCR
│   └── golden_compare.py      # Compara XML produzido com golden file
├── registry/
│   ├── chart_detector.py      # OCR + heurísticas pra detectar tipo de gráfico em imagem
│   └── sql_to_chart.py        # SQL → DataFrame → escolha de chart function
└── _agent/
    ├── loop.py                # Karpathy loop: try → test → fix → commit
    ├── memory.md              # O agente escreve aqui o que aprendeu
    └── BLOCKED.md             # Impasses não resolvidos
```

---

## 3. As 3 estratégias e quando cada uma vence

| Estratégia | Quando usar | Exemplo |
|---|---|---|
| **XML puro** | Mudança que tem elemento OOXML nativo: cor de série, dLbl, axis format, fill | Pintar linha de azul, mudar label fonte |
| **Híbrido win32 validador** | Sempre. Toda mudança XML termina rodando o probe win32 num CI Windows | Detectar quando uma edição em c:dPt corrompe silenciosamente |
| **Overlay de shapes** | O que OOXML não suporta nativamente: label colorida ao fim (sim, suporta dLbl mas com limitações de posicionamento), area-between em line, annotation lateral livre, sombreamento custom | Os 3 gráficos do MVP entram aqui em algum momento |

**Regra de fallback automático:** cada chart function tenta XML puro primeiro. Se a feature pedida não tem elemento nativo (verificado contra uma tabela em `xml_primitives/feature_matrix.json`), cai pra overlay. O usuário nunca decide isso — é decisão interna da função.

---

## 4. MVP: os 3 gráficos travados

### 4.1 `line_with_end_label`
Gráfico de linhas histórico (estilo CDI vs Andes que você mandou). Cada série tem cor própria, e ao final da série uma caixinha retangular com fundo igual à cor da linha e texto branco/preto com o último valor formatado.

**Estratégia:**
- XML: cores de série via `c:ser/c:spPr/a:ln/a:solidFill/a:srgbClr`
- Overlay: label colorida ao final, porque dLbl nativo não permite background colorido por série de forma confiável e a posição "fim da linha" é frágil.
- Posicionador: lê `c:plotArea` para EMU, lê última categoria da série, calcula posição (x = plotArea.right + small offset, y = interpolação do último valor no eixo Y).

### 4.2 `line_area_between`
Duas linhas com área sombreada (cor + transparência) entre elas. Inclui o caso em que se cruzam (sombrear o "delta" em duas cores diferentes ou uma única cor neutra).

**Estratégia:**
- XML puro não tem area-between nativo. Tem area chart, mas combinar area+line+area em um chart group é nasty.
- Overlay com `freeform shape`: gera o polígono da área entre as duas linhas, calcula vértices em EMU via plot area, aplica fill com transparência.
- Edge case do cruzamento: detectar interseções por interpolação linear e gerar dois polígonos.

### 4.3 `line_stacked_combo`
Linha sobreposta a colunas empilhadas + annotation lateral livre.

**Estratégia:**
- XML puro para o combo chart: `c:barChart` + `c:lineChart` no mesmo `c:plotArea` (python-pptx já permite parcialmente; precisamos refinar barOverlap, gap, axisId).
- Overlay para o annotation: text box com leader line apontando pra coluna.
- Leader line: usa `p:cxnSp` (shape de conexão) entre duas âncoras.

---

## 5. Karpathy Loop — protocolo do agente

Cada iteração:

```
1. PULL: git pull (sempre antes de ler estado)
2. PICK: ler spec/CURRENT_TASK.md. Se vazio, ler spec/BACKLOG.md e pegar a primeira task com status: ready.
3. PLAN: escrever 1-paragraph plan em _agent/scratch.md. Listar arquivos que vai tocar.
4. WRITE: implementar.
5. TEST: rodar `pytest tests/regression/test_<feature>.py -xvs`. Se falhar, ir pra 6, senão 7.
6. FIX: ler stack trace, formar hipótese, escrever em _agent/memory.md, voltar pra 4. Limite: 5 ciclos. Após 5, escrever em BLOCKED.md, fazer rollback (`git restore`), pular task.
7. VALIDATE: rodar suite completa: schema XML + win32 probe (se Windows) + render-OCR check.
8. COMMIT: `git add -A && git commit -m "feat(...): ..."`. Atualiza spec/BACKLOG.md, status: done.
9. PUSH: git push.
10. LOOP.
```

**Regra anti-com_error:** o agente não pode escrever `try: ... except: pass` em código de produção. Toda exceção sobe ou é tratada com lógica explícita. O `try/except` só é permitido em `_agent/loop.py` para o próprio loop de retry.

---

## 6. Validação multicamada

Para cada gráfico produzido pelo MVP, 4 gates antes de marcar como done:

1. **Schema gate**: o XML do chart valida contra `dml-chart.xsd`.
2. **Roundtrip gate**: abre o pptx, salva sem mudanças, diff dos XMLs deve ser vazio (ou só whitespace).
3. **Win32 probe gate** (CI Windows): abre o pptx via `win32com.client.Dispatch("PowerPoint.Application")`, espera 2s, fecha. Se PowerPoint reclamar (alerta de corrupção, repair dialog), gate falha. Capturar isso via leitura do log do Office é a parte chata — fallback: comparar bytes do pptx antes/depois do open.
4. **Render-OCR gate**: `soffice --headless --convert-to png slide.pptx`, OCR via tesseract no PNG, checar que strings esperadas (último valor da série, labels de annotation) aparecem.

---

## 7. SQL → Chart (visão da API final)

```python
from kinea_pptx_plus import Studio

studio = Studio(template="kinea_master.pptx")

slide = studio.new_slide(layout="chart_full")
slide.add_chart_from_sql(
    sql="SELECT date, andes_pct, cdi_pct FROM perf.fund_vs_index WHERE fund='Andes'",
    chart_hint="line_with_end_label",  # opcional; se omitido, detector escolhe
    title="Andes vs CDI — Retorno acumulado",
)
studio.save("output.pptx")
```

**Quando `chart_hint` é omitido e o usuário fornece uma imagem de referência:**

```python
slide.add_chart_from_sql(sql=..., reference_image="screenshot.png")
# detector OCR + heurística escolhe entre line_with_end_label, line_area_between, line_stacked_combo
```

O `chart_detector` é simples no MVP: olha número de séries, presença de área sombreada (detecta polígonos no preview), presença de barras empilhadas.

---

## 8. Stack técnica

- Python 3.11
- `python-pptx-ng` 1.0.x (fork mantido) como base
- `lxml` para edição XML direta
- `pywin32` (Windows-only, opcional, com fallback)
- `Pillow` + `pytesseract` para OCR e validação de render
- `pandas` para dados
- `psycopg2-binary` ou `sqlalchemy` para conexão com fin-data-lab
- `pytest` + `pytest-xdist` para suite
- `pre-commit` com `black`, `ruff`, `mypy` (--strict nas pastas de XML primitives)

---

## 9. Estrutura de commits e tags

- Cada feature: `feat(<chart-or-primitive>): descrição curta`
- Cada fix de regressão: `fix(<area>): ...`
- Cada validação nova: `test(<chart>): adiciona golden xml para line_with_end_label`
- Tag `v0.1.0` quando os 3 gráficos do MVP passam todos os 4 gates.
- Tag `v0.2.0` quando o SQL → chart end-to-end funciona pro line_with_end_label.

---

## 10. Riscos conhecidos e como o agente lida

| Risco | Mitigação |
|---|---|
| Corrupção silenciosa do XML | Win32 probe gate obrigatório no CI |
| python-pptx-ng diverge demais do scanny | Pinning estrito + suite de regressão antes de bumps |
| Overlay de shape desalinha quando slide é redimensionado | Posicionador trabalha em coordenadas relativas ao plotArea, recalcula no save |
| OCR falha em fonts custom da Kinea | Fallback: comparar bounding boxes vs golden, não só texto |
| LibreOffice render diverge de PowerPoint render | Render gate é heurístico, não bloqueante; win32 probe é o gate bloqueante |
| DeepSeek V4 Flash trava num XML complexo | Loop tem limite de 5 retries; após isso, task vai pra BLOCKED.md e agente pula |

---

## 11. O que está EXPLICITAMENTE fora do escopo do MVP

- Scatter, treemap, boxplot, waterfall, marimekko: ficam pra v0.3+. Backlog registrado, mas agente não toca enquanto MVP não fecha.
- Animações.
- Conexão real-time com banco. SQL → chart é one-shot.
- UI/dashboard. CLI + Python API só.
- Suporte a .ppt (legacy). Só .pptx.

---

## 12. Definition of Done (DoD) do MVP

- [ ] 3 chart functions implementadas
- [ ] Cada uma com test golden xml + test render-OCR
- [ ] Cada uma documentada com docstring estendida + entrada no README
- [ ] Win32 probe gate passa nos 3 num CI Windows (ou agente registra em BLOCKED.md se não tiver acesso a Windows)
- [ ] Exemplo end-to-end SQL → chart roda pra line_with_end_label com dataset Andes vs CDI
- [ ] README.md raiz do repo explica como um novo agente continua o trabalho

---

## 13. Para você, Gabriel, na manhã seguinte

Quando você acordar, leia nesta ordem:

1. `_agent/memory.md` — o que o agente aprendeu na madrugada
2. `_agent/BLOCKED.md` — onde ele travou e precisa de você
3. `git log --oneline` desde o último commit que você viu — o progresso
4. `tests/regression/` os goldens novos — pra você dar OK visual
5. Rodar `python examples/andes_vs_cdi.py` e abrir o pptx no PowerPoint

Se algo estiver lindo demais, abra o XML produzido e procure por gambiarra. O agente é honesto mas o flash é flash.
