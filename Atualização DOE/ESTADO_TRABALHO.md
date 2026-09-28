# Estado do trabalho — atualização SPGG × DOE

Última gravação: 08/09/2026. Checagem manual (CHECAGEM 17) conferida pelo usuário: **arquivo pronto para transferência**.

Checagem conferida de novo em 10/09/2026 (histórico 03/09 + lote `FG_Cargos_Comissao_DOE_04-10-09-2026`). **Tudo correto.** Pronto para substituir a SPGG na nuvem (`CASA_CIVIL/NOVA_ESTRUTURA`) com `SPGG_PARA_ENVIAR_NUVEM.xlsx`.

## Pasta

`C:\projetos Grok Build\Acoes documentos\Atualização DOE`

## Arquivo a usar daqui para frente

**`SPGG_PARA_ENVIAR_NUVEM.xlsx`** — arquivo da nuvem (não colar SPGG_TESTE).

**`CHECAGEM_atualizacao_SPGG.xlsx`** — conferência do período **11/05/2026 a 25/08/2026**:
- `Vagas_novas` (43 funções com VAGA_1 novo)
- `Movimentacoes` (1.116 vagas já existentes atualizadas)

- Montado por edição cirúrgica da original (não reescrever o xlsx no openpyxl — isso corrompia no Office 365).
- Abrir este no 365 para validar. Se abrir para **editar**, sem aviso de pasta danificada, pode substituir a original na nuvem.

**Não usar** gerações anteriores do mesmo nome feitas via openpyxl (abriam só em leitura no 365).

A original local (cópia baixada, intacta): `TB_R_CARGOS_SPGG.xlsx`  
DOE: somente `C:\projetos Grok Build\Acoes documentos\AgenteDOE\FG_Cargos_Comissao.xlsx` (gerado 29/08/2026 14:43, 2112 atos, 11/05–25/08).

## Decisões fechadas

1. Chave: `Vaga n°` (DOE) = `VAGA_1` (SPGG).
2. Um ato por vaga: último ato de **ocupação** (NOMEIA, DESIGNA, COLOCA, **ATRIBUI**, EXONERA, DISPENSA, CESSA).
3. RETIFICA/ALTERA/APOSTILA/RELOTA **não** trocam titular.
3b. TORNA/DECLARA SEM EFEITO: posse/desistência/publicação indevida cancela a pessoa; FGS/FGT (FG→CC, caso Evandro) **não** cancela a movimentação — a pessoa permanece e a vaga fica CC.
4. Vaga inexistente + NOMEIA/DESIGNA = **criar linha** (caso Bruna `102771000001`).
5. Data DOE vazia: usar a seção `DATA DO DOE: dd/mm/aaaa`.
6. **Banco de Talentos / TB_R_RH não é mais usado.** PROCV externo convertido em valor. Nome, telefone, partido e indicação ficam escritos na `TUDO POR VAGA`.
7. Nome já existente na SPGG: leva **ID RH**. Nome só do DOE: **ID RH em branco**.
8. Quem sai: `NOME ANTERIOR` + `ID_RH ANT.` (se houver ID).
9. Fórmulas que permanecem (da vaga): `NÍVEL` = RIGHT do CÓDIGO; `VALOR` = FATOR×1950 se CC, ×60% se FG (não depende mais de ter ID).
9b. **ENCARGO** é da `VAGA_1`, não do nome. Em vaga existente não se altera; só se preenche em vaga nova.
10. Sem Excel desktop: enviar pelo navegador (OneDrive → substituir). Histórico de versões antes de substituir.
11. STATUS na carga: OCUPADO na entrada, LIVRE na saída (ainda pode ajustar se a equipe quiser outro critério).

## Casos-guia

| Vaga | Resultado |
|---|---|
| `102221000010` | Evandro (ID 2458) titular, Ana anterior (11214), vaga **CC** CCS/FGS 13 |
| `102222000010` | Vinícius entra, Evandro sai |
| `102223000340` | Lea entra, Bruna sai |
| `102771000001` | **Nova.** Bruna titular, ID 10177, FG / CCT/FGT 10 |

## Números da última geração limpa (com DOE 26/08)

- 1107 vagas atualizadas
- 50 linhas novas (primeiro provimento)
- Exemplo 13/08: vaga `102616000001` → Fábio Hax Duro (ID 10128), anterior Eduardo Moura Mendes
- PROCV Banco de Talentos zerado na TUDO POR VAGA e INDIRETA

## Como rodar offline

Ver `COMO_RODAR_OFFLINE.md`. Lançador: `python rodar_atualizacao.py` nesta pasta (menu: extrair PDFs e/ou gerar SPGG + checagem).

## Scripts (somente o necessário)

- `rodar_atualizacao.py` — lançador offline
- `gerar_teste_atualizacao.py` — regras de cruzamento
- `gerar_spgg_nuvem_limpo.py` — gera SPGG da nuvem + checagem
- `COMO_RODAR_OFFLINE.md` — instruções

## Testes (conferência, não enviar)

`TESTE_atualizacao_SPGG.xlsx` … `_v5.xlsx` — evolução das regras. v5 tem abas Conferencia e CARGA_PROCV.

## Pendências

- Validar `SPGG_PARA_ENVIAR_NUVEM.xlsx` no Office 365 (edição, não só leitura).
- Se ok: substituir a `TB_R_CARGOS_SPGG.xlsx` na pasta OneDrive `CASA_CIVIL/NOVA_ESTRUTURA`.
- `Mat. Servidor` não foi definido se acompanha o nome.
- STATUS LIVRE na exoneração: confirmar com a equipe se for o caso.
