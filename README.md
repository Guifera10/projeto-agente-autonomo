# Projeto Agente Autônomo (DOE → FG → SPGG)

Consulta o Diário Oficial do Estado (DOE), extrai atos de cargos em comissão e funções gratificadas e aplica o lote na planilha local da SPGG. O agente **não envia** nada para a nuvem: só atualiza `SPGG_PARA_ENVIAR_NUVEM.xlsx` neste computador.

O Git guarda o código. Planilhas, PDFs e o histórico de `resultados\` ficam na máquina que já está rodando o ciclo.

API: `https://doe-backend.pro.rs.gov.br`

---

## Como rodar em outro computador

### 1. Acesso

- Conta GitHub com permissão neste repositório (é privado)
- Git instalado
- Internet para a API do DOE

### 2. Máquina

- Windows (o ciclo usa `.cmd`, PowerShell e Tarefa Agendada)
- Usuário **logado** na sessão: a tarefa é Interactive e, sem sessão, não dispara
- Python 3.13 (ou 3.11+), disponível no PATH como `python`

### 3. Código

A pasta pode ser qualquer uma. O orquestrador usa as subpastas do próprio clone (`AgenteDOE\`, `Atualização DOE\`, `resultados\`).

```
git clone https://github.com/Guifera10/projeto-agente-autonomo.git
cd projeto-agente-autonomo
python -m pip install pandas openpyxl pdfplumber
```

### 4. Base da SPGG (obrigatória para atualizar a planilha)

Copiar da máquina que já roda o agente para a **mesma estrutura** no clone, usando a pasta de `resultados` mais recente. Exemplo com 19/10/2026:

- `resultados\2026-10-19\SPGG_PARA_ENVIAR_NUVEM.xlsx`
- `resultados\2026-10-19\CHECAGEM_atualizacao_SPGG.xlsx`

Sem esses dois arquivos o DOE e o extrato FG rodam; o patch da SPGG precisa da planilha.

Feche o Excel nesses arquivos antes de copiar e antes de rodar o ciclo.

### 5. Caminho do Python no agendador

`rodar_ciclo.cmd` e `agendar_tarefa.ps1` apontam para o Python desta máquina:

`C:\Users\guilherme-medina\AppData\Local\Programs\Python\Python313\python.exe`

No outro PC, troque pelo `python.exe` local **ou** rode na mão:

```
python orquestrador_doe.py --vigia
```

### 6. Tarefa das 07:00 e 19:00 (opcional)

Depois de ajustar o Python:

```
powershell -ExecutionPolicy Bypass -File .\agendar_tarefa.ps1
```

Isso cria a tarefa **DOE Agente Autonomo** naquele Windows.

### 7. Conferência

```
python orquestrador_doe.py --so-detectar
```

Se a API responder, o terminal está pronto. Ciclo completo (vigia de 10 em 10 min até existir pasta do dia):

```
python orquestrador_doe.py --vigia
```

---

## O que o ciclo faz

1. Varre cada dia desde a última pasta com SPGG em `resultados\` até hoje.
2. Consulta o DOE tipo 1; se houver extra, baixa a última atualização (`/pdf?nroEdicao=N`).
3. Extrai FG com `AgenteDOE\agente_fg_comissao.py`.
4. Aplica o lote em `Atualização DOE\` sobre a SPGG da pasta **anterior**.
5. Grava PDF, extrato, SPGG, checagem e estado em `resultados\AAAA-MM-DD\`.

Se o DOE do dia ainda não saiu, a vigia consulta a cada 10 minutos até gravar a pasta; com o resultado pronto, encerra. Extra nova entra na sessão das 19:00.

---

## Comandos úteis

```
python orquestrador_doe.py --so-detectar
python orquestrador_doe.py --so-baixar
python orquestrador_doe.py
python orquestrador_doe.py --vigia
python orquestrador_doe.py --sem-spgg
python orquestrador_doe.py --a-partir-de 2026-09-11
python orquestrador_doe.py --forcar
```

`--forcar` ignora `estado_doe.json` na varredura. Extra do mesmo dia entra mesmo com a pasta do dia já gravada.

Detalhes do pipeline e das regras: `MEMORIA.md`. Extração e cruzamento manuais: `Atualização DOE\COMO_RODAR_OFFLINE.md`.
