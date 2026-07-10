# Instruções de atualização

Fluxo operacional para atualizar a base e o dashboard quando novos dados chegarem do CREA-BA (ex.: relatório de cursos extraído do SITAC, novas planilhas de instituições, revisão de pendências).

## 0. Edição rápida de uma pendência pontual (via navegador, sem instalar nada)

Para corrigir um dado cadastral específico (ex.: preencher um CNPJ ou Código MEC faltante, corrigir a grafia de um município, ajustar uma categoria):

1. Abra a tabela correspondente no GitHub: [`dados/fonte/Relatorio_Instituicoes.csv`](https://github.com/adinailson88/cursos-crea/blob/master/dados/fonte/Relatorio_Instituicoes.csv) (instituições/campi) ou [`dados/fonte/Titulos_SITAC.csv`](https://github.com/adinailson88/cursos-crea/blob/master/dados/fonte/Titulos_SITAC.csv) (títulos profissionais).
2. Clique no ícone de lápis (✏️) no canto superior direito para editar.
3. Corrija a célula/linha necessária (mantendo a estrutura de colunas) e clique em **Commit changes...** direto na branch `master`.
4. O [Action "Atualizar dashboard CREA-BA"](https://github.com/adinailson88/cursos-crea/actions/workflows/atualizar_dashboard.yml) dispara sozinho a cada alteração em `dados/fonte/` ou `scripts/`, roda os três scripts abaixo e publica o resultado. Se quiser forçar manualmente, clique em **Run workflow** nesse link.
5. Em 1–2 minutos o GitHub Pages republica e o dashboard já reflete a mudança.

Isso cobre correções pontuais de dado cadastral. Para trocas grandes de fonte (nova extração completa do SITAC) ou para editar normas/base legal, siga o fluxo completo abaixo (passos 1–9), que também pode ser rodado localmente.

## 1. Atualizar dados de origem

- Substitua ou adicione os arquivos correspondentes em `dados/fonte/` (formato CSV, UTF-8) — pela edição rápida acima (passo 0) ou substituindo o arquivo inteiro localmente.
- Se o novo arquivo vier do Google Drive como planilha nativa, exporte como CSV (`Arquivo → Fazer download → Valores separados por vírgula`), preservando o cabeçalho original.
- Se houver um relatório de cursos completo do SITAC, salve-o como `dados/fonte/Cursos_SITAC.csv` com colunas equivalentes a: instituição, campus, curso, situação, modalidade, título/código, ato de registro — e ajuste `scripts/gerar_base.py` (seção 5, `CURSOS + CURSO_TITULO`) para lê-lo em vez do `Consolidado_EXEMPLO.csv`.

## 2. Executar a normalização

```bash
cd "Tabelas COREC"
python scripts/gerar_base.py
```

Isso reescreve `planilha/Base_IES_Cursos_CREA_BA.xlsx` do zero a partir de `dados/fonte/`.

## 3. Revisar pendências

- Abra a aba `PENDENCIAS` da planilha (ou `dados/validacao.json` após o passo 5) e revise os itens de gravidade `ALTA` primeiro.
- Correções pontuais de mapeamento (ex.: novo alias de município, nova sigla) devem ser feitas nas constantes no topo de `scripts/gerar_base.py` (`ALIASES_MUNICIPIO`, `MARCADORES_SITUACAO`) — não diretamente na planilha gerada, pois ela é sobrescrita a cada execução.
- Vínculos de base legal (norma/artigo) só devem ser adicionados em `BASE_LEGAL_ROWS` / `VINCULOS_TITULO_ATRIBUICAO` dentro do script quando confirmados em fonte oficial (CONFEA, CREA-BA, Diário Oficial) — nunca copiando texto genérico tipo "art. X e Y da lei tal".

## 4. Gerar BASE_DASHBOARD

A aba `BASE_DASHBOARD` é gerada automaticamente pelo próprio `gerar_base.py` (passo 2) — não é necessário nenhum passo manual adicional.

## 5. Gerar dados.json

```bash
python scripts/gerar_json.py
```

## 6. Validar

```bash
python scripts/validar_dados.py
```

Corrija qualquer item marcado `[ERRO]` antes de publicar. Itens `[AVISO]` são aceitáveis para publicação, desde que documentados em `RELATORIO_VALIDACAO.md`.

## 7. Publicar

- Local: `python -m http.server 8000` na raiz do projeto e abrir `http://localhost:8000/dashboard/index.html`.
- GitHub Pages / hospedagem estática: enviar a pasta inteira do projeto (ou `dashboard/` + `dados/`) para o repositório/host, mantendo o caminho relativo `../dados/dados.json` usado em `dashboard/app.js`.

## 8. Conferir o dashboard

- Indicadores no topo batem com `dados/dados.json → meta`;
- busca geral e filtros combinados retornam resultados coerentes;
- clique em um município no mapa filtra a tabela;
- exportação CSV funciona e reflete os filtros ativos;
- modal de detalhe abre ao clicar em uma linha da tabela.

## 9. Atualizar a documentação

Depois de rodar os três scripts, atualize os números citados em `RELATORIO_VALIDACAO.md` (e em `README.md` se a estrutura de arquivos mudou) com os valores reais impressos por `scripts/validar_dados.py` e pelo conteúdo de `dados/dados.json` — nunca reutilize números de uma execução anterior.
