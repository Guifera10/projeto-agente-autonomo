# Memória — Projeto Agente Autônomo

Atualizado em 07/10/2026. Pasta canónica:

`C:\projetos Grok Build\Acoes documentos\Projeto Agente Autonomo`

Encadeia as pastas irmãs **AgenteDOE** (extração FG/cargos a partir do PDF) e **Atualização DOE** (cruzamento e patch da SPGG). Não envia nada para a nuvem: só atualiza a planilha local `SPGG_PARA_ENVIAR_NUVEM.xlsx`.

---

## O que existe nesta pasta

| Arquivo | Função |
|---|---|
| `orquestrador_doe.py` | Ciclo: consulta API do DOE → baixa PDF → extrai FG → aplica lote na SPGG → grava `resultados/AAAA-MM-DD/` |
| `rodar_ciclo.cmd` | Wrapper da tarefa agendada (UTF-8, log em `resultados\ultimo_ciclo.log`) |
| `agendar_tarefa.ps1` | Cria/atualiza a tarefa Windows **DOE Agente Autonomo** |
| `estado_doe.json` | IDs já processados (fonte da verdade do orquestrador) |
| `MEMORIA.md` | Este arquivo |
| `resultados\` | Uma subpasta por dia extraído + log e resumo do último ciclo |

Python: `C:\Users\guilherme-medina\AppData\Local\Programs\Python\Python313\python.exe`

API: `https://doe-backend.pro.rs.gov.br`

---

## Como rodar

```
cd "C:\projetos Grok Build\Acoes documentos\Projeto Agente Autonomo"
python orquestrador_doe.py --so-detectar
python orquestrador_doe.py --so-baixar
python orquestrador_doe.py
python orquestrador_doe.py --vigia
python orquestrador_doe.py --watch --intervalo 15
python orquestrador_doe.py --sem-spgg
python orquestrador_doe.py --a-partir-de 2026-09-11
python orquestrador_doe.py --forcar
```

`--forcar` ignora `estado_doe.json` na varredura (último dia com SPGG até hoje). Extra do mesmo dia entra mesmo com a pasta do dia já gravada.

Tarefa Windows: **DOE Agente Autonomo**, **07:00** e **19:00**. Se o DOE do dia ainda não saiu, vigia a cada 10 min até gravar `resultados/AAAA-MM-DD/`; com o resultado pronto, encerra. Extra entra na sessão das 19:00.

---

## Regras fechadas

1. **Base da SPGG** = último `SPGG_PARA_ENVIAR_NUVEM.xlsx` em pasta de `resultados` **estritamente anterior** à data do DOE. Extra do mesmo dia reprocessa em cima do dia anterior, não em cima da ordinária da manhã.
2. **Um resultado por dia.** Se a API ainda não tiver o DOE de hoje, a vigia consulta de 10 em 10 min até gravar a pasta; com o resultado pronto, encerra (volta aos horários 07:00 e 19:00). Extra nova na sessão das 19:00 reprocessa a pasta do dia. Sem DOE até 19:00: grava a pasta com a SPGG anterior.
3. O resultado do ciclo vai para `resultados/AAAA-MM-DD/` (PDF, extrato FG deste ciclo, SPGG, checagem, estado, resumo).
4. O orquestrador **não refiltra** atos. Quem decide o que entra no Excel FG é o `agente_fg_comissao.py`.
5. Piso absoluto: não processar DOE anterior a **11/09/2026** (`DATA_INICIO_PADRAO`). Testes de 09/09 e 10/09 foram apagados.
6. Aviso `openpyxl` “Data Validation extension is not supported” é inócuo.
7. Se o lote incremental não aplicar nenhuma vaga (`atualizadas 0` e `novas 0`), a `CHECAGEM_atualizacao_SPGG.xlsx` **não é regravada**: permanece a última lista que teve movimentações.

---

## Pipeline

1. Varre **cada dia** desde a última pasta com SPGG em `resultados/` até hoje (inclui o último dia, para extra posterior).
2. Em cada dia, consulta o DOE **tipo 1**. Se houver extra, baixa a **última atualização** (`/pdf?nroEdicao=N`).
3. Confere tamanho e número de páginas com a API; PDF incompleto é rebaixado. Não reutiliza arquivo só porque tem mais de 50 KB.
4. Extrai FG daquele PDF. Se 0 atos: grava a pasta do dia e **não** marca sucesso definitivo. Não reutiliza extrato de outro dia.
5. Aplica SPGG em cima da pasta **anterior** (extra do mesmo dia não empilha na ordinária da manhã).
6. Copia para `resultados/AAAA-MM-DD/` o PDF, o extrato **deste** ciclo, SPGG, checagem e estado. Rede: até 4 tentativas. Falha no meio da varredura interrompe o catch-up para não furar a cadeia.

Tarefa Windows: **DOE Agente Autonomo**, 07:00 e 19:00, limite **16 h**. Vigia de 10 min só enquanto o dia não tiver pasta. Se o PC estiver desligado, rodar `python orquestrador_doe.py --vigia` nesta pasta.

Pastas vivas (dentro deste projeto):

- `AgenteDOE\`
- `Atualização DOE\`

---

## Histórico do que foi criado e feito

### 09/09/2026 — código e primeiro teste

- Criado/ajustado `orquestrador_doe.py` (a pasta no início só tinha este script).
- Primeiro teste: DOE ordinário **174**, 09/09/2026, id **54719**, 422 pág. Extração de 5 atos (incl. Disposição no Estado, depois filtrada no AgenteDOE). Aplicou EXONERA de Márcia Scherer na SPGG de produção — **efeito colateral**.
- Resultados passaram a ser copiados para a pasta do código (`resultados\`).

### 10/09/2026 — limpeza e pausa

- Segundo ciclo automático/manual: DOE **175**, 10/09/2026, id **54721**.
- Pedido para apagar testes: removidos PDFs, extratos datados, `estado_doe.json` da Atualização DOE e a pasta `resultados`.
- Produção restaurada:
  - `FG_Cargos_Comissao.xlsx` ← cópia de 04/09 (`FG_Cargos_Comissao_20260904_1108.xlsx`)
  - `SPGG_PARA_ENVIAR_NUVEM.xlsx` ← OneDrive `SPGG_PARA_ENVIAR_NUVEM 7.xlsx` (09/09 08:39)
  - `CHECAGEM_atualizacao_SPGG.xlsx` ← checagem 17 (08/09)
- Agente em modo **offline** (sem SPGG), válido só a partir de **11/09**. Tarefa adiada para 11/09 07:00.
- Criados `rodar_ciclo.cmd` e `agendar_tarefa.ps1` (tarefa `DOE Agente Autonomo`).

### 10/09 tarde — religar função original

- SPGG manualmente atualizada pelo usuário (~11:54).
- Orquestrador voltou ao ciclo completo **DOE → FG → SPGG**.
- Resultados em pasta **por dia**: `resultados/AAAA-MM-DD/`.
- Flag `--sem-spgg` no lugar do antigo `--com-spgg` (SPGG ligada por padrão).

### 11/09/2026 — primeiro ciclo válido

- Tarefa das 07:00 perdida (bateria). Ciclo manual ~09:49.
- DOE **176**, 11/09/2026, id **54723**, 148 pág.
- 2 atos: DISPENSA Gabriela Carvalho Nascimento (`102593000034`, FGT 03, a contar 01/09) e EXONERA Irineo Luiz Grings (`102224000431`, CCT 10, a contar 24/08). SPGG: 2 atualizadas, 0 novas.
- Pasta: `resultados\2026-09-11\`.

### 11/09 — regra da cadeia de SPGG

- Base de cada atualização = último `SPGG_PARA_ENVIAR_NUVEM.xlsx` em `resultados`.
- Só DOE com data **posterior** à dessa pasta.
- Confirmado: edição 176 passou a ser ignorada depois de gravada.

### 14/09/2026 — segundo ciclo + regeneração

- Tarefa 07:00: DOE **177**, 14/09/2026, id **54725**, 193 pág. Extrato 07:00 (35 registros) aplicado na SPGG de 11/09.
- 08:21 atualização do `agente_fg_comissao.py`; 08:22 reextração (`FG_Cargos_Comissao_20260914_0822.xlsx`).
- SPGG **regenerada** (~08:29) com base `resultados/2026-09-11` + extrato 08:22: **16** vagas atualizadas, **0** novas, **1** ignorada (data posterior na SPGG).
- `SPGG_PARA_ENVIAR_NUVEM.xlsx` igual em `Atualização DOE\` e em `resultados\2026-09-14\`. Essa de 14/09 é a base do próximo DOE.

### 22/09/2026 — PDF errado e refação

- A API sem `tipoDiario=1` devolveu o diário **tipo 0** (id **54736**, 9 páginas). Extração: 0 atos; o orquestrador reaplicou o extrato de 21/09.
- DOE oficial é **tipo 1**, id **54737**, 189 páginas. Corrigido `consultar_ultima()` para filtrar `tipoDiario=1`. Extração vazia **não** reutiliza mais o Excel do dia anterior.
- Refação: 107 registros (96 atos + 11 quadros); SPGG a partir de `resultados/2026-09-21`: **63** atualizadas, **2** novas, 0 ignoradas.

---

## 07/10/2026 — correções no orquestrador

O ciclo das 07:00 processou a ordinária 194 (id 54758, 400 pág.). O código passou a:

- recuperar dias em falta (não só `consultar_ultima()`);
- usar a última extra do dia via `nroEdicao`;
- rebaixar PDF se as páginas não baterem com a API;
- copiar só o FG gerado neste ciclo (não o mais recente da pasta AgenteDOE por mtime);
- avisar com clareza se o Excel estiver com a SPGG aberta;
- repetir a rede e parar o catch-up no primeiro erro do dia, para a cadeia da SPGG continuar contínua.

`--extras` permanece no CLI por compatibilidade; o padrão já é a última atualização do dia.

## Estado atual (07/10/2026)

Último processado: `54758|2026-10-07|194|ORDINARIA`

Última SPGG da cadeia: `resultados\2026-10-07\SPGG_PARA_ENVIAR_NUVEM.xlsx`

Próximo ciclo: varre a partir de 07/10 (para extra do mesmo dia) até hoje.

Conteúdo típico de `resultados/AAAA-MM-DD/`:

- `DOE_AAAA-MM-DD_ORDINARIA_<id>.pdf`
- `FG_Cargos_Comissao.xlsx` e `FG_Cargos_Comissao_YYYYMMDD_HHMM.xlsx`
- `SPGG_PARA_ENVIAR_NUVEM.xlsx`
- `CHECAGEM_atualizacao_SPGG.xlsx`
- `estado_doe.json`
- `ultimo_resultado.txt`

Na raiz de `resultados\`: `ultimo_ciclo.log`, `ultimo_resultado.txt`.

---

## O que este projeto não faz

- Não substitui a SPGG no OneDrive (`CASA_CIVIL/NOVA_ESTRUTURA`).
- Não altera as regras de extração do AgenteDOE (skill `agente-doe`).
- Não reprocessa consolidado histórico de PDFs (11/05–03/09). Isso continua no AgenteDOE, à mão.

Detalhe da extração FG e das regras SPGG: memória do workspace Grok e `Atualização DOE\ESTADO_TRABALHO.md`.
