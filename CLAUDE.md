# CLAUDE.md

Guia para o Claude Code (e para pessoas desenvolvedoras) trabalharem neste repositório.

## Visão geral

**Refugo** é um dashboard em **Streamlit** para acompanhar e **projetar o refugo (PPM) de fim de ano** da
região **BU1 SA** e de cada planta. Ele lê uma planilha no layout "Dados brutos" (exportação SAP/BW),
calcula o realizado, projeta os meses restantes e simula cenários para atingir os targets de PPM.

- Idioma da interface, do código (nomes, comentários) e da documentação: **português (pt-BR)**.
- Requisitos funcionais e regras de negócio: ver [`REQUISITOS.md`](REQUISITOS.md).

## Estrutura

```
app_refugo_v3.py   # aplicação inteira (arquivo único, ~1650 linhas)
requirements.txt   # dependências Python
REQUISITOS.md      # documento de requisitos
README.md
Dados brutos.xlsx  # (opcional, NÃO versionado) arquivo padrão lido se nenhum upload for feito
```

### Organização de `app_refugo_v3.py`

O script é executado de cima para baixo a cada interação (modelo do Streamlit). Seções, na ordem:

1. **Constantes** — `REGION`, `MESES`, paleta `CORES`, nomes das metodologias (`M_ANO`, `M_MOVEL`, ...).
2. **Dados** — `load_data()` (leitura do Excel, `@st.cache_data`), `build_data()`, `year_tables()`,
   `targets_table()`, `to_raw_layout()` (exporta de volta no layout "Dados brutos").
3. **Cálculo** — `base_production()`, `base_ppm()`, `project()`, `summarize()`, formatação (`fmt_num`, `fmt_pct`).
4. **Sidebar** — upload, ano da projeção, metodologia da produção base, período do PPM base, restaurar padrões.
5. **Edições** — `raw_store` em `st.session_state` guarda edições da aba "Dados de entrada" por (tipo, ano).
6. **Explicações** — dicionário `HELP` (tooltip de cada coluna) e `show_table()` (tabela + popover de memorial).
7. **Abas** (`tab0` … `tab6`): Situação atual, Resumo histórico, Produzido × Refugo, Sensibilidade por planta,
   Combinações BU1 SA vs Target, Simulador de combinações, Dados de entrada, Memorial de cálculo.

## Comandos

```bash
pip install -r requirements.txt
streamlit run app_refugo_v3.py          # abre em http://localhost:8501
python -m py_compile app_refugo_v3.py   # checagem rápida de sintaxe
```

Não há suíte de testes automatizados nem linter configurado. Para validar uma mudança, rode o app com uma
planilha de exemplo e confira as abas afetadas (e o Memorial de cálculo).

## Formato da planilha de entrada ("Dados brutos")

- Uma linha de cabeçalho contendo o texto `Calendar Year/Month`; as colunas de mês têm o formato `m.aaaa`
  (ex.: `1.2026`, `10.2026`) — qualquer quantidade de anos.
- Coluna 0 = conta: **201** = peças refugadas, **202** = peças produzidas. Coluna 3 = código da planta,
  coluna 4 = nome da planta.
- Bloco **Targets**: célula com o texto `Targets`; nas linhas abaixo, nome (planta ou `BU1 SA`) na coluna 0 e
  target PPM na coluna 1, até a primeira linha vazia.

## Regras de negócio essenciais

- `PPM = refugo / produção × 1.000.000`. Na região **BU1 SA** somam-se refugo e produção antes de dividir
  (média ponderada, **nunca** média simples dos PPMs).
- Mês **realizado** = todas as plantas com produção **e** refugo preenchidos; os demais são **projetados**.
- Produção projetada = produção base × peso Geral % × peso Planta %.
- Refugo projetado = PPM base × produção projetada ÷ 10⁶ × (1 + variação %).
- Fronteira do target é linear nas quantidades refugadas (produção fixa) — usada em Combinações/Simulador.

## Convenções

- Números no padrão brasileiro: use `fmt_num` / `fmt_pct` (milhar `.`, decimal `,`); gráficos Plotly com
  `separators=",."`.
- Cores: uma cor fixa por planta (`color_of`), target em `COR_TARGET`, linha tracejada para metas.
  Vermelho = piora / acima da meta; verde = melhora / dentro da meta.
- Toda coluna calculada nova deve ganhar uma explicação em `HELP` (ou `extra_help`) e, se for regra nova,
  uma entrada na aba **Memorial de cálculo**.
- Chaves de widgets (`key=`) devem ser únicas e, quando dependem de contexto (ano, plantas, arquivo),
  incluir esse contexto na chave (ex.: `W_KEY`, `raw_editor_key`).
- Estado que precisa sobreviver a widgets ocultos vai em estruturas próprias no `st.session_state`
  (padrão `raw_store`), não apenas no estado do widget.
- Use a API atual do Streamlit (`width="stretch"` em vez de `use_container_width`).
- Mantenha o app em arquivo único, salvo pedido explícito para modularizar.

## Cuidados

- **Não versionar dados reais** (`*.xlsx` com produção/refugo das plantas) — estão no `.gitignore`.
- Nomes de plantas não devem ser fixados no código (exceção existente: padrão `Itajubá` no mapa de combinações,
  com fallback).
- Alterações em fórmulas devem manter coerência entre cálculo, tooltips (`HELP`) e Memorial de cálculo.
