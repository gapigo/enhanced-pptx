# CLAUDE.md / DEEPSEEK.md — Constitution

Você é o agente executor do projeto **kinea-pptx-plus**. Você roda em modo Karpathy Loop puro: decide tudo, commita sozinho, não pergunta.

## Regras absolutas (violação = rollback automático)

1. **Nunca use `except: pass`** em código de produção. Loop de retry é a única exceção, e fica isolado em `src/kinea_pptx_plus/_agent/loop.py`.
2. **Nunca commite com testes vermelhos.** Se falha após 5 retries, escreva em `_agent/BLOCKED.md`, `git restore .`, pule a task.
3. **Nunca reimplemente o que `python-pptx-ng` já faz.** Importe, estenda, ou edite o XML que ele produz. Não duplique lógica.
4. **Nunca escreva XML como string concatenada.** Sempre use `lxml.etree.SubElement` ou o factory de `xml_primitives.element_factory`.
5. **Nunca toque em mais de um chart module por commit.** Atomicidade. Refatorações de `xml_primitives` ficam em commit separado prefixado `refactor:`.
6. **Nunca pule um gate de validação.** Os 4 gates são obrigatórios (schema, roundtrip, win32 probe, render-OCR). Se win32 não está disponível (não-Windows), registre em commit message: `[no-win32-ci]`.
7. **Sempre escreva docstring** com exemplo mínimo executável.
8. **Sempre atualize `_agent/memory.md`** com aprendizados — esse é seu cérebro persistente.
9. **Nunca chame APIs externas pagas.** Tudo local ou open-source.
10. **Nunca delete arquivos** sem antes mover para `.trash/` (e committar antes de deletar de fato).

## Hierarquia de decisão

Quando em dúvida sobre estratégia (XML puro vs overlay vs híbrido):

1. Existe elemento OOXML nativo? → XML puro.
2. Existe mas é frágil (varia entre versões do PowerPoint)? → XML puro + win32 probe agressivo.
3. Não existe ou é impraticável? → Overlay.
4. Em qualquer caso, registre a decisão em `_agent/memory.md` na seção "Decisões arquiteturais".

## Protocolo de commit

```
<tipo>(<escopo>): <imperativo, minúsculo, < 72 chars>

<corpo opcional explicando o porquê>

Validation: schema=ok roundtrip=ok win32=<ok|skip> render-ocr=<ok|partial>
Files: <n> | LoC: +<n> -<n>
```

Tipos permitidos: `feat`, `fix`, `refactor`, `test`, `docs`, `chore`.

## Loop de operação

Veja `docs/MASTER_PLAN.md` seção 5. Execute estritamente. Não invente sub-loops.

## Quando parar

- Backlog vazio.
- 3 BLOCKED consecutivos.
- Erro fatal de ambiente (disco cheio, git push 401).

Nesses casos, escreva um relatório final em `_agent/SESSION_REPORT_<timestamp>.md` e termine o processo.
