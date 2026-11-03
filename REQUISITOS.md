# Documento de Requisitos — Refugo (Projeção de PPM BU1 SA)

| Item | Valor |
|---|---|
| Produto | Dashboard de avaliação/simulação de refugo |
| Versão do app | v3 (`app_refugo_v3.py`) |
| Data | 03/11/2026 |
| Status | Rascunho — documenta o comportamento atual |

## 1. Objetivo

Fornecer à gestão de qualidade da região **BU1 SA** uma ferramenta visual para:

1. acompanhar o refugo realizado (peças e PPM) por planta e para a região;
2. **projetar o PPM de fechamento do ano** a partir do realizado e de premissas de produção/refugo;
3. comparar com os **targets** e com o **ano anterior**;
4. simular cenários e descobrir **quanto cada planta (ou combinação de plantas) precisa reduzir** para que a
   região atinja o target.

## 2. Escopo

**Dentro do escopo:** leitura de planilha Excel no layout "Dados brutos", cálculos de PPM, projeção mensal,
análises de sensibilidade/combinações, simulador interativo, edição dos dados no app e exportação em Excel.

**Fora do escopo:** integração direta com SAP/BW, banco de dados, autenticação de usuários, persistência
de edições entre sessões (as edições vivem apenas na sessão do navegador).

## 3. Usuários

| Perfil | Necessidade |
|---|---|
| Gestor(a) de qualidade regional | Visão consolidada BU1 SA, gap vs target, plantas críticas |
| Analista de qualidade da planta | Situação da própria planta, histórico, quanto precisa reduzir |
| Analista de dados / PCP | Ajustar premissas de produção, lançar meses, exportar resultados |

## 4. Glossário

| Termo | Definição |
|---|---|
| **PPM** | Peças refugadas ÷ peças produzidas × 1.000.000 |
| **BU1 SA** | Região que agrega todas as plantas do arquivo |
| **Mês realizado** | Mês em que **todas** as plantas têm produção e refugo preenchidos |
| **Mês projetado** | Mês do ano da projeção que não é realizado |
| **Produção base** | Produção mensal de referência para os meses projetados |
| **PPM base** | PPM de referência usado para projetar o refugo |
| **Target** | Meta de PPM anual por planta e para a região |
| **Gap** | (PPM projetado − Target) ÷ Target × 100 |

## 5. Requisitos funcionais

### 5.1 Entrada de dados

- **RF-01** — O sistema deve aceitar upload de arquivo `.xlsx` no layout "Dados brutos" pela barra lateral.
- **RF-02** — Sem upload, o sistema deve ler o arquivo padrão `Dados brutos.xlsx` na pasta do app; se não
  existir, exibir erro orientando o upload.
- **RF-03** — O sistema deve localizar a linha de cabeçalho pelo texto `Calendar Year/Month` e reconhecer
  colunas de mês no formato `m.aaaa`, para **qualquer quantidade de anos**.
- **RF-04** — Linhas com conta **201** devem ser lidas como refugo e conta **202** como produção; código da
  planta na coluna 3 e nome na coluna 4.
- **RF-05** — Os targets devem ser lidos do bloco iniciado pela célula `Targets` (nome na coluna 0, valor na
  coluna 1) até a primeira linha vazia.
- **RF-06** — Erros de leitura (layout inesperado, arquivo sem dados) devem gerar mensagem clara e
  interromper a execução sem travar o app.
- **RF-07** — Anos sem nenhum dado devem ser ignorados nas análises.

### 5.2 Configuração (barra lateral)

- **RF-08** — Seleção do **ano da projeção** (padrão: último ano com dados); o ano anterior é usado como
  histórico/comparação quando tiver meses completos.
- **RF-09** — Seleção da metodologia de **produção base**:
  - Média do ano (meses realizados do ano da projeção);
  - Média móvel dos últimos X meses (pode cruzar anos; X configurável);
  - Meses selecionados (ano atual e anterior);
  - Média do ano anterior *(se houver ano anterior)*;
  - Mesmo mês do ano anterior *(se houver; sem dado → média do ano anterior)*.
- **RF-10** — Seleção do período do **PPM base**: média do ano, mesmo período da produção ou ano anterior.
- **RF-11** — Botão **Restaurar padrões** que limpa todo o estado da sessão.

### 5.3 Cabeçalho e indicadores

- **RF-12** — Exibir: PPM projetado BU1 SA (com gap vs target), target BU1 SA, PPM realizado até o momento,
  nº de plantas dentro do target e PPM do ano anterior (com variação).

### 5.4 Aba "Situação atual"

- **RF-13** — Seleção do mês de referência entre os meses realizados.
- **RF-14** — Cartões por planta e região para o **mês** e para o **acumulado do ano**: PPM, desvio vs meta
  (PPM, % e peças), status e comparação com o mesmo período do ano anterior.
- **RF-15** — Gráfico de histórico mensal (PPM, peças refugadas ou produção), visão região ou por planta,
  com comparação opcional com o ano anterior e linha de target.
- **RF-16** — Tabelas de desvio mensal vs meta (%) e de variação vs mesmo mês do ano anterior (%).

### 5.5 Aba "Resumo histórico"

- **RF-17** — Linha do tempo e comparação ano a ano do indicador escolhido, por planta ou região.
- **RF-18** — Tabela anual por planta, variação anual (mesmo intervalo de meses), tabela mês a mês e
  sazonalidade da produção (% da média mensal do ano). Anos incompletos indicam o nº de meses.

### 5.6 Aba "Produzido × Refugo"

- **RF-19** — Tabela editável de **pesos da produção projetada** por mês projetado (Geral e por planta, em %
  da base; peso efetivo = Geral × Planta).
- **RF-20** — Tabela editável de **variação do refugo projetado** (%) por planta.
- **RF-21** — Resumo anual por planta e região: produção/refugo real e projetado, PPM real, PPM ano
  projetado, target, gap, peças para atingir o target, status e comparação com o ano anterior.
- **RF-22** — Tabela mensal Jan–Dez de produção, refugo e PPM, destacando os meses projetados.
- **RF-23** — Download da projeção em Excel (`projecao_refugo.xlsx`).

### 5.7 Aba "Sensibilidade por planta"

- **RF-24** — Para uma planta e faixa de variação escolhidas, mostrar o efeito no PPM da planta e da região.
- **RF-25** — Tabela de **refugo permitido** por planta para atingir o próprio target (total e por mês) e
  variação necessária vs base.

### 5.8 Aba "Combinações BU1 SA vs Target"

- **RF-26** — Gráfico da influência de cada planta (variação −50% a +50%) no PPM BU1 SA, com linha de target.
- **RF-27** — Mapa de calor de combinações entre duas plantas (verde abaixo do target, vermelho acima).
- **RF-28** — Tabela da **redução uniforme necessária** para cada combinação de plantas, com peças a reduzir
  e indicação de viabilidade (variação ≥ −100%).

### 5.9 Aba "Simulador de combinações"

- **RF-29** — Selecionar 2 plantas (gráfico 2D) ou 3 plantas (gráfico 3D / plano de corte) e a faixa dos eixos.
- **RF-30** — Exibir a fronteira em que o PPM BU1 SA fica exatamente no target e o ponto analisado; clique
  no gráfico 2D move o ponto.
- **RF-31** — Tabela "Valores no ponto analisado" editável, permitindo fixar o refugo anual de plantas.

### 5.10 Aba "Dados de entrada"

- **RF-32** — Exibir e permitir editar produção, refugo (por ano, de 2020 até o ano mais recente) e targets;
  todas as telas devem ser recalculadas.
- **RF-33** — Preencher produção e refugo de todas as plantas em um mês o torna realizado; apagar o torna
  projetado.
- **RF-34** — Destacar valores alterados vs Excel original e informar a quantidade na barra lateral.
- **RF-35** — Botão para restaurar os dados originais do Excel.
- **RF-36** — Download dos dados editados e de um template em branco, ambos no layout "Dados brutos"
  (reenviáveis pelo upload).

### 5.11 Aba "Memorial de cálculo" e transparência

- **RF-37** — Documentar todas as fórmulas (PPM, produção base, projeção, fechamento, refugo permitido,
  combinações), com exemplo numérico de uma planta no cenário atual e glossário das colunas.
- **RF-38** — Toda coluna calculada deve ter tooltip explicativo e popover "ℹ️ Memorial de cálculo".

## 6. Regras de negócio

| ID | Regra |
|---|---|
| RN-01 | `PPM = Refugo ÷ Produção × 10⁶` |
| RN-02 | PPM da região = Σ refugo das plantas ÷ Σ produção das plantas × 10⁶ (ponderado pela produção) |
| RN-03 | Mês realizado exige produção **e** refugo de **todas** as plantas |
| RN-04 | `Prod_base = média da produção nos meses usados pela metodologia` |
| RN-05 | `Prod_m = Prod_base × Peso_Geral,m/100 × Peso_Planta,m/100` |
| RN-06 | `PPM_base = Σ refugo ÷ Σ produção × 10⁶` no período escolhido |
| RN-07 | `Refugo_m = PPM_base × Prod_m ÷ 10⁶ × (1 + Var%)` |
| RN-08 | `PPM_ano = (Refugo_real + Σ Refugo_proj) ÷ (Prod_real + Σ Prod_proj) × 10⁶` |
| RN-09 | `Gap % = (PPM_ano − Target) ÷ Target × 100`; status ✅ Dentro se Gap ≤ 0, ❌ Fora se Gap > 0 |
| RN-10 | `Peças p/ atingir target = Refugo_ano − Target × Prod_ano ÷ 10⁶` |
| RN-11 | `Refugo permitido = Target × Prod_ano ÷ 10⁶ − Refugo_real`; `Var_nec = Refugo_permitido ÷ Refugo_proj,base − 1` |
| RN-12 | Com produção fixa, a fronteira do target é linear nas quantidades refugadas (reta em 2D, plano em 3D) |
| RN-13 | Combinação viável se a variação uniforme necessária for ≥ −100% |
| RN-14 | Análises de projeção (Sensibilidade, Combinações, Simulador) só se aplicam quando há meses projetados |

## 7. Requisitos não funcionais

- **RNF-01 — Tecnologia:** Python 3.10+, Streamlit, pandas, NumPy, Plotly, openpyxl (ver `requirements.txt`).
- **RNF-02 — Idioma e formatos:** interface em pt-BR; números com milhar `.` e decimal `,`; meses `Jan/26`.
- **RNF-03 — Desempenho:** leitura do Excel em cache (`st.cache_data`); recálculo interativo em poucos
  segundos para ~8 plantas e vários anos.
- **RNF-04 — Usabilidade:** layout largo; cores fixas por planta; vermelho = piora, verde = melhora;
  linha tracejada para targets.
- **RNF-05 — Transparência:** todos os números devem ser rastreáveis até a fórmula (tooltips + memorial).
- **RNF-06 — Portabilidade dos dados:** arquivos exportados devem manter o layout de entrada.
- **RNF-07 — Confidencialidade:** dados reais de produção/refugo não devem ser versionados no repositório.
- **RNF-08 — Robustez:** células vazias/não numéricas tratadas como ausentes, sem quebrar o app.

## 8. Restrições e premissas

- O layout "Dados brutos" é a única fonte de dados suportada.
- Targets são únicos por planta (aplicados a todos os anos; o histórico compara com o target atual).
- Edições feitas no app não são salvas automaticamente; o usuário deve baixar o Excel editado.

## 9. Melhorias sugeridas (backlog)

- Testes automatizados das funções de cálculo (`base_production`, `project`, `summarize`).
- Modularizar o app (dados, cálculo, UI) mantendo o comportamento.
- Targets por ano.
- Arquivo de exemplo com dados fictícios para demonstração.
