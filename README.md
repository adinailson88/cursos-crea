# Base única padronizada e Dashboard — Instituições, Cursos, Títulos e Atribuições (CREA-BA)

### 🔗 [Abrir o dashboard ao vivo](https://adinailson88.github.io/cursos-crea/dashboard/index.html)

## 1. Objetivo

Consolidar em uma base única, normalizada e auditável as informações espalhadas em seis planilhas do Google Drive do CREA-BA sobre:

- instituições de ensino e seus campi cadastrados;
- cursos ofertados (quando disponíveis na fonte);
- títulos profissionais e códigos SITAC;
- referências normativas (leis, resoluções) que sustentam as atribuições profissionais;

e publicar essas informações em um **dashboard HTML** com busca, filtros combinados, tabela exportável e um **mapa interativo da Bahia**, deixando explícito o que está confirmado e o que ainda depende de validação.

Este painel é uma ferramenta de consulta e transparência. **Não substitui certidão, decisão de Câmara Especializada ou análise formal do CREA-BA.**

## 2. Estrutura de arquivos

```
Tabelas COREC/
├── dados/
│   ├── fonte/              # exports brutos das 6 planilhas do Drive (CSV) + referência de municípios IBGE
│   ├── processados/        # arquivos intermediários (vazio em uso normal)
│   ├── dados.json          # gerado por scripts/gerar_json.py — consumido pelo dashboard
│   └── validacao.json      # gerado por scripts/validar_dados.py
├── planilha/
│   └── Base_IES_Cursos_CREA_BA.xlsx   # planilha única normalizada (13 abas, ver seção 4)
├── dashboard/
│   ├── index.html
│   ├── styles.css
│   ├── app.js
│   └── assets/
├── scripts/
│   ├── gerar_base.py       # dados/fonte/*.csv -> planilha/Base_IES_Cursos_CREA_BA.xlsx
│   ├── gerar_json.py       # planilha -> dados/dados.json
│   └── validar_dados.py    # checagens automáticas -> dados/validacao.json
├── README.md
├── INSTRUCOES_ATUALIZACAO.md
└── RELATORIO_VALIDACAO.md
```

## 3. Fontes utilizadas

| Arquivo original (Drive) | Papel |
|---|---|
| `Relatório Instituições` | Fonte institucional **principal** (488 linhas → 416 instituições / 473 campi únicos) |
| `Relação de Instituições da Câmara de Elétrica` (+ "por modalidade") | Cruzamento/validação — **não** usadas como fonte primária |
| `sitac_relatorios_relatorio_generico_08-08-2025` | Catálogo de títulos profissionais (56 linhas → 51 códigos únicos) |
| `Consolidado` (aba `EXEMPLO`) | 7 registros de **exemplo** de curso+título+endereço — usados só como modelo/pista de cruzamento, nunca como cadastro validado |
| `Base_Padronizada_IES_Cursos_CREA_BA` | Diagnóstico anterior, usado apenas como **referência** de ordem de grandeza, não como modelo de colunas |
| `municipios_bahia_ibge.csv` | 417 municípios da Bahia com centroide, derivado do IBGE via [kelvins/municipios-brasileiros](https://github.com/kelvins/municipios-brasileiros) — usado para geocodificação municipal |
| Lei nº 5.194/1966, Resolução CONFEA nº 218/1973, Resolução CONFEA nº 1.156/2025 | Base legal confirmada via consulta a fontes oficiais (`planalto.gov.br`, `normativos.confea.org.br`) nesta sessão de trabalho |

Todas as fontes brutas usadas pelos scripts estão arquivadas em `dados/fonte/` para auditoria e reprocessamento.

## 4. Modelo de dados (planilha única)

`planilha/Base_IES_Cursos_CREA_BA.xlsx` contém:

- **RESUMO** — indicadores gerais;
- **IES** — 1 linha por instituição única (agrupada por CNPJ quando válido, por nome normalizado quando não há CNPJ confiável);
- **CAMPI** — 1 linha por campus único, com município, coordenadas e situação;
- **CURSOS** — cursos disponíveis (hoje, só os 7 de exemplo do arquivo `Consolidado`, todos marcados como não publicáveis);
- **TITULOS_PROFISSIONAIS** — 51 códigos únicos, com forma feminina/masculina separadas;
- **TITULOS_EQUIVALENCIAS** — ocorrências alternativas do mesmo código (duplicidades do catálogo de origem, preservadas);
- **CURSO_TITULO** — vínculo curso↔título;
- **BASE_LEGAL** — normas confirmadas (Lei 5.194/1966, Resolução CONFEA 218/1973, Resolução CONFEA 1.156/2025);
- **ATRIBUICOES** — áreas de atuação descritas nas normas acima;
- **TITULO_ATRIBUICAO_LEGAL** — vínculo título↔atribuição↔norma (artigo específico ainda pendente de extração integral do texto da norma);
- **FONTES** — 1 linha por arquivo/fonte utilizada;
- **PENDENCIAS** — registro estruturado de tudo que não pôde ser confirmado;
- **BASE_DASHBOARD** — tabela desnormalizada (1 linha por instituição+campus+curso+título+norma, quando existir) que alimenta o dashboard.

Princípios aplicados (detalhados em `scripts/gerar_base.py`): nada é inventado (campos não confirmados recebem `PENDENTE DE VALIDAÇÃO`); nenhum registro é apagado — duplicados, transferidos, históricos e inativos são preservados com a situação marcada; CNPJ/códigos são sempre texto; apenas registros com `publicar_dashboard = SIM` aparecem no modo público do dashboard.

## 5. Como executar/regerar

Pré-requisitos: Python 3.10+ com `openpyxl` (`pip install openpyxl`).

```bash
cd "Tabelas COREC"
python scripts/gerar_base.py      # dados/fonte/*.csv -> planilha/Base_IES_Cursos_CREA_BA.xlsx
python scripts/gerar_json.py      # planilha -> dados/dados.json
python scripts/validar_dados.py   # checagens -> dados/validacao.json
```

## 6. Como abrir o dashboard localmente

Navegadores bloqueiam `fetch()` em arquivos `file://`. É necessário servir a pasta por HTTP:

```bash
cd "Tabelas COREC"
python -m http.server 8000
# abrir http://localhost:8000/dashboard/index.html
```

## 7. Publicação

O dashboard é estático (HTML/CSS/JS + `dados.json`) e pode ser publicado em:

- **GitHub Pages**: publicar a pasta `Tabelas COREC/` (ou só `dashboard/` + `dados/`, ajustando o caminho `../dados/dados.json` em `app.js`) na branch do Pages;
- **Google Sites** (incorporação via iframe apontando para a URL publicada);
- qualquer servidor estático.

Bibliotecas externas usadas via CDN (documentadas): [Leaflet 1.9.4](https://leafletjs.com/) (mapa) + tiles do OpenStreetMap.

## 8. Limitações conhecidas

- **Cursos**: a base não contém a relação completa de cursos por instituição — só existem os 7 registros de exemplo do arquivo `Consolidado`, marcados como não publicáveis. É necessário extrair do SITAC um relatório de cursos completo (instituição, campus, curso, situação, modalidade, título/código, ato de registro).
- **Base legal por artigo**: as normas (Lei 5.194/1966, Resoluções CONFEA 218/1973 e 1.156/2025) foram confirmadas em fonte oficial, mas o artigo exato de cada título dentro da Resolução 1.156/2025 ainda não foi extraído do texto integral — o campo `artigo` permanece `PENDENTE DE VALIDAÇÃO` nesses casos.
- **Município do campus**: 144 de 473 campi não puderam ser associados automaticamente a um município da Bahia (nomes de unidade/sigla no lugar do município, ex. "CETEP", "CFTB", campos vazios) — precisam de revisão manual.
- Ver `RELATORIO_VALIDACAO.md` para os números completos e `dados/validacao.json` / aba `PENDENCIAS` para o detalhe de cada item.

## 9. Critérios de validação

Ver `scripts/validar_dados.py` (IDs duplicados, CNPJ inválido, coordenadas fora da faixa da Bahia, código de título inexistente, base legal sem link, divergência de contagem planilha×JSON, entre outros) e `RELATORIO_VALIDACAO.md`.

## 10. Política de atualização

Ver `INSTRUCOES_ATUALIZACAO.md`.
