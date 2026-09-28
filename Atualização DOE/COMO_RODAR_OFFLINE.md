# Como rodar offline

Não precisa de internet depois de instalar as bibliotecas **uma vez**.

## Instalação (uma vez, com internet)

```
python -m pip install pandas openpyxl
```

Para extrair PDF do DOE (AgenteDOE):

```
python -m pip install pdfplumber openpyxl
```

## Duas etapas

### 1. Extrair os PDFs → Excel FG_Cargos

```
cd "C:\projetos Grok Build\Acoes documentos\AgenteDOE"
python rodar_agente.py
```

Escolha os PDFs (`1`, `1 3 5` ou `all`). Sai `FG_Cargos_Comissao_AAAAMMDD_HHMM.xlsx` na mesma pasta. O agente também atualiza `FG_Cargos_Comissao.xlsx`.

### 2. Cruzar com a SPGG → arquivo da nuvem + checagem

Na pasta `C:\projetos Grok Build\Acoes documentos\Atualização DOE` precisa existir a cópia original:

- `TB_R_CARGOS_SPGG.xlsx`

Depois:

```
cd "C:\projetos Grok Build\Acoes documentos\Atualização DOE"
python rodar_atualizacao.py
```

Menu:

- **1** — só extrair PDFs (chama o AgenteDOE)
- **2** — só gerar SPGG + checagem (escolhe o FG_Cargos)
- **3** — os dois

Ou já apontar o arquivo:

```
python rodar_atualizacao.py "C:\projetos Grok Build\Acoes documentos\AgenteDOE\FG_Cargos_Comissao_20260903_1306.xlsx"
```

**Uma fonte por geração:** só o Excel que você escolher. Não mistura com `Cargos em Comissão Atualizado`.

## O que sai

| Arquivo | Uso |
|---|---|
| `SPGG_PARA_ENVIAR_NUVEM.xlsx` | Substituir a SPGG no OneDrive |
| `CHECAGEM_atualizacao_SPGG.xlsx` | Conferência (não enviar) |

## Regras que o cruzamento aplica

- Titular: NOMEIA, DESIGNA, COLOCA, ATRIBUI, EXONERA, DISPENSA, CESSA
- Apostila: não é movimentação
- RETIFICA de tipo FG/CC: só CC/FG e CÓDIGO
- SEM EFEITO: só tira titular em posse/desistência/publicação indevida
- Data: **A contar de** se preenchida; senão Data DOE
- ENCARGO fica na vaga
- ID RH só se a pessoa já existia na SPGG
- `VAGA_1` na checagem com o número completo

## Enviar à nuvem

1. Fechar a SPGG no Excel Online  
2. Histórico de versões da `TB_R_CARGOS_SPGG.xlsx`  
3. Enviar `SPGG_PARA_ENVIAR_NUVEM.xlsx` e substituir (renomear se o OneDrive pedir)  
