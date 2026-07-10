# Relatório de Validação

Gerado a partir da execução de `scripts/gerar_base.py`, `scripts/gerar_json.py` e `scripts/validar_dados.py` em **2026-07-10**, sobre os arquivos-fonte em `dados/fonte/` (ver README.md, seção 3).

## 1. Volumetria

| Item | Quantidade |
|---|---:|
| Linhas lidas de `Relatório Instituições` (fonte principal) | 488 |
| Instituições (IES) únicas após normalização | 416 |
| Campi únicos após normalização | 473 |
| Municípios da Bahia identificados entre os campi | 79 |
| Linhas lidas do catálogo de títulos (SITAC) | 56 |
| Códigos de título profissional únicos | 51 |
| Ocorrências alternativas preservadas em `TITULOS_EQUIVALENCIAS` | 5 |
| Cursos na base (todos de exemplo, arquivo `Consolidado`) | 7 |
| Vínculos curso→título (`CURSO_TITULO`) | 7 |
| Normas confirmadas em `BASE_LEGAL` | 3 |
| Vínculos título→atribuição→norma (`TITULO_ATRIBUICAO_LEGAL`) | 7 |
| Linhas em `BASE_DASHBOARD` | 475 |
| Registros publicáveis (`publicar_dashboard = SIM`) | 261 |
| Pendências abertas | 698 |
| % completude cadastral (indicador do dashboard) | 50,0% |

## 2. Qualidade cadastral das instituições (IES)

| Item | Quantidade | % de 416 |
|---|---:|---:|
| CNPJ inválido ou fictício (`00000000000000`) | 34 | 8,2% |
| Sem Código MEC informado | 192 | 46,2% |
| Categoria acadêmica PRIVADA | 123 | — |
| Categoria acadêmica PÚBLICA | 120 | — |
| Categoria acadêmica NÃO INFORMADA | 173 | — |

## 3. Qualidade cadastral dos campi

| Situação do campus | Quantidade |
|---|---:|
| Situação não informada (padrão) | 329 |
| Transferido para outro conselho | 63 |
| Não identificado (campo vazio ou só sigla de unidade) | 75 |
| Histórico (nome antigo) | 3 |
| Inativo (desativado/descredenciado/suspenso) | 3 |
| **Sem município confirmado** (dos não transferidos) | **144** |

## 4. Duplicidades e divergências tratadas

- **Códigos de título duplicados no catálogo de origem** (linha vazia + linha preenchida para o mesmo código, ou grafias divergentes): confirmados os casos citados no escopo original — `1210800`, `1210900`, `1211000`, `1221103` — todos consolidados em 1 descrição principal por código, com as demais ocorrências preservadas em `TITULOS_EQUIVALENCIAS` (5 linhas).
- **Distinção `1211101` × `1211102`**: confirmada no catálogo de origem como dois códigos distintos (`1211101` = Engenheiro Industrial - Elétrica; `1211102` = Engenheiro Industrial Eletrônica), sem ambiguidade no catálogo.
- **Divergência no exemplo do IFBA**: o registro de exemplo (`Consolidado/EXEMPLO`) vincula o curso "ENGENHARIA INDUSTRIAL - ELÉTRICA" ao código `1211102` (Eletrônica), quando o catálogo associa "Elétrica" ao código `1211101`. A divergência **não foi corrigida silenciosamente**: está registrada como pendência de gravidade ALTA em `PENDENCIAS`, com o código sugerido para verificação humana.
- **Registros duplicados na fonte de instituições**: linhas idênticas de instituição+campus foram consolidadas em um único registro de `CAMPI`, com o número de repetições anotado em `observacao` e uma pendência de gravidade BAIXA (nenhuma linha foi descartada — a fonte bruta permanece arquivada em `dados/fonte/Relatorio_Instituicoes.csv`).

## 5. Base legal

Nenhum texto de "lei tal" / "art. X e Y" presente nos arquivos de origem foi usado como base legal. Três normas foram confirmadas via consulta a fontes oficiais nesta sessão:

| Norma | Situação | Fonte |
|---|---|---|
| Lei nº 5.194/1966 | Vigente | planalto.gov.br |
| Resolução CONFEA nº 218/1973 | Parcialmente revogada (arts. 8º e 9º revogados pela Resolução 1.156/2025; art. 24 revogado pela Resolução 1.057/2014) | normativos.confea.org.br |
| Resolução CONFEA nº 1.156/2025 | Vigente — consolida as competências da engenharia elétrica/eletrônica/computação/industrial elétrica | normativos.confea.org.br |

Para os 7 vínculos título→atribuição→norma criados, a **norma** está confirmada, mas o **artigo específico** dentro da Resolução 1.156/2025 permanece `PENDENTE DE VALIDAÇÃO` — é necessário ler o texto integral/anexos da norma para extrair o artigo exato de cada título.

## 6. Cursos

Os únicos 7 registros de curso existentes vêm do arquivo `Consolidado` (aba `EXEMPLO`) e **não foram tratados como cadastro validado** em nenhum momento: `status_validacao = PENDENTE DE VALIDAÇÃO` e `publicar_dashboard = NÃO` em todos. Nenhum deles aparece no dashboard público por padrão (checkbox "somente publicáveis" ligado).

## 7. Pendências por tipo de registro

| Tipo de registro | Pendências |
|---|---:|
| IES | 516 |
| CAMPUS | 158 |
| TITULO | 10 |
| CURSO | 7 |
| TITULO_ATRIBUICAO_LEGAL | 7 |

Por gravidade: **ALTA** = 41 · **MÉDIA** = 466 · **BAIXA** = 191.

## 8. Verificações automáticas (`scripts/validar_dados.py`)

Resultado da última execução: **14 OK · 3 avisos · 0 erros**.

Avisos (não bloqueiam publicação, mas indicam trabalho pendente):

1. 34 instituições com CNPJ inválido ou fictício de 416.
2. 192 instituições sem Código MEC de 416.
3. 144 campi sem município identificado de 473.

Nenhum erro estrutural foi encontrado: sem IDs duplicados em nenhuma aba, sem coordenadas fora da faixa geográfica da Bahia, sem curso apontando para campus inexistente, sem vínculo curso→título para código inexistente, sem norma sem URL oficial, sem vínculo legal sem norma associada, e a contagem de registros bate entre a planilha e `dados/dados.json`.

## 9. Correções realizadas nesta rodada

- Separação de código MEC embutido no nome da instituição (`"500068070 - AGES - ..."` → nome limpo + `codigo_mec`).
- Extração de marcadores de situação (`[TRANSFERIDO PARA OUTRO CONSELHO]`, `[NOME ANTIGO]`, `DESATIVADA`, `[SEM CADASTRO]`, `[DESCREDENCIADA DO MEC]`, `CADASTRO SUSPENSO...`, `EXCLUIR`) do campo de campus para o campo `situacao_campus`, preservando o texto original em `campus_original`.
- Validação de dígito verificador de CNPJ (módulo 11) e marcação de CNPJ fictício `00000000000000`.
- Geocodificação municipal via centroide oficial do IBGE (79 municípios confirmados), com pequenas correções de grafia objetivas (ex. "TANCREDO NEVES" → "Presidente Tancredo Neves"; "CRUS DAS ALMAS" → "Cruz das Almas"; "Ihéus" → "Ilhéus"; "XIQUE - XIQUE" → "Xique-Xique").
- Separação de título feminino/masculino a partir da coluna `TRATAMENTO` do catálogo SITAC, usando a coluna `TÍTULO` (forma masculina) como referência de corte.
- Inferência de natureza administrativa (federal/estadual/municipal) **apenas** quando evidenciada no próprio nome oficial da instituição (ex. "UNIVERSIDADE FEDERAL..." → PÚBLICA FEDERAL); nos demais casos com categoria PÚBLICA, o campo permanece `PENDENTE DE VALIDAÇÃO`.

## 10. Itens não corrigidos por falta de evidência

- 144 campi sem município identificável automaticamente (siglas de unidade como "CETEP", "CFTB", "CEPA", campos vazios ou "SEM IDENTIFICAÇÃO") — exigem consulta ao SITAC ou conhecimento local para resolver.
- 192 instituições sem Código MEC — não foi feita busca externa no e-MEC nesta rodada (fora do escopo autorizado desta sessão).
- Artigo exato de cada título dentro da Resolução CONFEA 1.156/2025 — norma confirmada, mas o texto integral/anexos não foi lido artigo por artigo nesta sessão.
- Relação completa de cursos por instituição — não existe nos arquivos de origem; depende de extração futura de um relatório de cursos do SITAC (ver `INSTRUCOES_ATUALIZACAO.md`, passo 1).
