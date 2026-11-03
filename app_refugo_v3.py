"""Projeção do refugo (PPM) final do ano — BU1 SA.

Executar:  streamlit run app_refugo_v3.py
"""
import hashlib
from datetime import date
from io import BytesIO
from itertools import combinations
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

DEFAULT_FILE = Path(__file__).parent / "Dados brutos.xlsx"
REGION = "BU1 SA"
MESES = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
# Paleta categórica (ordem fixa, uma cor por planta)
CORES = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#6250d6", "#e34948"]
COR_TARGET = "#52514e"
COR_REGIAO = "#0b0b0b"
# Verde (abaixo do target) -> neutro (no target) -> vermelho (acima do target)
ESCALA_PPM = [[0, "#1a9850"], [0.25, "#91cf60"], [0.5, "#f7f7f2"], [0.75, "#fc8d59"], [1, "#d73027"]]

st.set_page_config(page_title="Projeção Refugo", page_icon="📉", layout="wide")


# --------------------------------------------------------------------------- dados
TEMPLATE_START = 2020  # primeiro ano do template padrão
M_ANO = "Média do ano"
M_MOVEL = "Média móvel (últimos X meses)"
M_SEL = "Meses selecionados"
M_ANO_ANT = "Média do ano anterior"
M_MESMO_MES = "Mesmo mês do ano anterior"


def _month_of(header):
    """'1.2026' / 1.2026 / '10.2026' -> número do mês (1..12) ou None."""
    try:
        txt = f"{float(header):.4f}" if not isinstance(header, str) else header.strip()
        m = int(txt.split(".")[0])
        return m if 1 <= m <= 12 else None
    except (ValueError, TypeError):
        return None


def _year_of(header):
    try:
        txt = f"{float(header):.4f}" if not isinstance(header, str) else header.strip()
        return int(txt.split(".")[1])
    except (ValueError, TypeError, IndexError):
        return None


def plabel(y, m):
    """Rótulo de período: Jan/26."""
    return f"{MESES[m - 1]}/{str(y)[2:]}"


@st.cache_data
def load_data(content: bytes):
    """Lê o layout 'Dados brutos' com qualquer quantidade de anos (colunas m.aaaa)."""
    raw = pd.read_excel(BytesIO(content), header=None)

    hdr_row = next(i for i in range(len(raw))
                   if raw.iloc[i].astype(str).str.contains("Calendar Year/Month", case=False).any())
    month_cols = {}
    for c in range(raw.shape[1]):
        m = _month_of(raw.iat[hdr_row, c])
        if m is not None:
            month_cols[c] = (_year_of(raw.iat[hdr_row, c]) or date.today().year, m)
    years = sorted({y for y, _ in month_cols.values()})

    prod = {y: {} for y in years}
    scrap = {y: {} for y in years}
    codes, plants = {}, []
    for i in range(hdr_row + 1, len(raw)):
        acc = pd.to_numeric(raw.iat[i, 0], errors="coerce")
        if acc not in (201, 202):
            continue
        plant = str(raw.iat[i, 4]).strip()
        codes[plant] = raw.iat[i, 3]
        if plant not in plants:
            plants.append(plant)
        dest = scrap if acc == 201 else prod
        for c, (y, m) in month_cols.items():
            dest[y].setdefault(plant, {})[m] = pd.to_numeric(raw.iat[i, c], errors="coerce")

    def frame(d):
        return pd.DataFrame(d, dtype=float).T.reindex(index=plants, columns=range(1, 13)).astype(float)

    # Targets: bloco abaixo da célula "Targets"
    targets = {}
    pos = np.argwhere(raw.astype(str).apply(lambda s: s.str.strip().str.lower()).values == "targets")
    if len(pos):
        r = pos[0][0] + 1
        while r < len(raw) and pd.notna(raw.iat[r, 0]):
            val = pd.to_numeric(raw.iat[r, 1], errors="coerce")
            if pd.notna(val):
                targets[str(raw.iat[r, 0]).strip()] = float(val)
            r += 1

    return {"plants": plants, "codes": codes, "years": years, "targets": targets,
            "prod": {y: frame(prod[y]) for y in years}, "scrap": {y: frame(scrap[y]) for y in years}}


def complete_months(prod_df, scrap_df):
    """Meses com produção e refugo de todas as plantas."""
    return [m for m in range(1, 13) if prod_df[m].notna().all() and scrap_df[m].notna().all()]


def build_data(plants, prod_df, scrap_df, targets, year):
    """Estrutura do ano da projeção; mês realizado = todas as plantas com produção e refugo."""
    realized = complete_months(prod_df, scrap_df)
    projected = [m for m in range(1, 13) if m not in realized]
    return {"plants": plants, "prod": prod_df, "scrap": scrap_df, "targets": targets,
            "realized": realized, "projected": projected, "year": year}


def empty_year(plants):
    return pd.DataFrame(np.nan, index=plants, columns=range(1, 13), dtype=float)


def year_tables(raw, y):
    """Tabelas editáveis (template) de um ano, com os valores originais do arquivo."""
    def to_tbl(df):
        t = df.copy()
        t.columns = MESES
        t.index.name = "Planta"
        return t
    return (to_tbl(raw["prod"].get(y, empty_year(raw["plants"]))),
            to_tbl(raw["scrap"].get(y, empty_year(raw["plants"]))))


def targets_table(raw):
    tg = pd.DataFrame({"Target PPM": [raw["targets"].get(p, np.nan) for p in raw["plants"] + [REGION]]},
                      index=raw["plants"] + [REGION])
    tg.index.name = "Planta"
    return tg


def to_raw_layout(plants, codes, prod_tbls, scrap_tbls, tg_tbl, years, blank=False):
    """Gera um .xlsx no layout 'Dados brutos' (colunas m.aaaa de todos os anos) — pode ser reenviado no app."""
    head = ["Account number QM", None, "Value Type", "Plant", "Location|Calendar Year/Month"] + \
           [f"{m}.{y}" for y in years for m in range(1, 13)]
    rows = [head]
    for acc, name, tbls in [(201, "Scrapped parts", scrap_tbls), (202, "Number of parts produced", prod_tbls)]:
        for pl in plants:
            vals = []
            for y in years:
                if blank or y not in tbls:
                    vals += [None] * 12
                else:
                    vals += [None if pd.isna(v) else float(v) for v in tbls[y].loc[pl, MESES]]
            rows.append([acc, name, 10, codes.get(pl), pl] + vals)
    rows += [[None] * len(head) for _ in range(5)]
    rows.append([None, "Targets"] + [None] * (len(head) - 2))
    for pl, r in tg_tbl.iterrows():
        rows.append([pl, None if blank or pd.isna(r["Target PPM"]) else float(r["Target PPM"])]
                    + [None] * (len(head) - 2))
    buf = BytesIO()
    pd.DataFrame(rows).to_excel(buf, header=False, index=False, sheet_name="Planilha1")
    return buf.getvalue()


# --------------------------------------------------------------------------- cálculo
def period_sum(hist, periods, plants):
    """Σ por planta de um conjunto de períodos (ano, mês)."""
    if not periods:
        return pd.Series(0.0, index=plants)
    return pd.concat([hist[y][m] for y, m in periods], axis=1).sum(axis=1)


def base_production(hist_prod, timeline, year, projected, plants, method, n_last, sel_periods):
    """Produção base por planta e mês projetado (DataFrame plantas × meses projetados) e períodos usados."""
    cur = [p for p in timeline if p[0] == year]
    prev = [p for p in timeline if p[0] == year - 1]
    if method == M_MOVEL:
        used = timeline[-n_last:]
    elif method == M_SEL:
        used = sel_periods or cur
    elif method == M_ANO_ANT:
        used = prev or cur
    elif method == M_MESMO_MES:
        used = [(year - 1, m) for m in projected if (year - 1, m) in prev]
    else:
        used = cur
    fallback = period_sum(hist_prod, prev or cur, plants) / max(len(prev or cur), 1)
    if method == M_MESMO_MES:
        base = pd.DataFrame({m: hist_prod[year - 1][m] if (year - 1, m) in prev else fallback
                             for m in projected}, index=plants)
    else:
        avg = period_sum(hist_prod, used, plants) / max(len(used), 1)
        base = pd.DataFrame({m: avg for m in projected}, index=plants)
    return base, used


def base_ppm(hist_prod, hist_scrap, periods, plants):
    """PPM médio do refugo no período (Σ refugo / Σ produção)."""
    return period_sum(hist_scrap, periods, plants) / period_sum(hist_prod, periods, plants) * 1e6


def project(data, prod_base, ppm_base, weights, variation):
    """Monta produção e refugo jan–dez do ano da projeção (real + projetado).

    prod_base: DataFrame plantas × meses projetados
    weights:   DataFrame meses projetados × plantas (fração, 1 = 100%)
    variation: Series planta -> variação do refugo (fração, 0 = base)
    """
    prod = data["prod"].copy()
    scrap = data["scrap"].copy()
    for m in data["projected"]:
        prod[m] = prod_base[m] * weights.loc[m]
        scrap[m] = ppm_base * prod[m] / 1e6 * (1 + variation)
    return prod, scrap


def pct(a, b):
    """Variação percentual (a / b − 1) × 100."""
    return (a / b - 1) * 100


def summarize(data, prod, scrap, prev=None):
    real = data["realized"]
    proj = data["projected"]
    df = pd.DataFrame({
        "Produção real": prod[real].sum(axis=1),
        "Refugo real": scrap[real].sum(axis=1),
        "Produção projetada": prod[proj].sum(axis=1),
        "Refugo projetado": scrap[proj].sum(axis=1),
    })
    df.loc[REGION] = df.sum()
    df["Produção ano"] = df["Produção real"] + df["Produção projetada"]
    df["Refugo ano"] = df["Refugo real"] + df["Refugo projetado"]
    df["PPM real"] = df["Refugo real"] / df["Produção real"] * 1e6
    df["PPM ano (proj.)"] = df["Refugo ano"] / df["Produção ano"] * 1e6
    df["Target PPM"] = [data["targets"].get(p, np.nan) for p in df.index]
    df["Gap vs target (%)"] = (df["PPM ano (proj.)"] - df["Target PPM"]) / df["Target PPM"] * 100
    df["Peças p/ atingir target"] = df["Refugo ano"] - df["Target PPM"] * df["Produção ano"] / 1e6
    df["Status"] = np.where(df["Target PPM"].isna(), "—",
                            np.where(df["Gap vs target (%)"] <= 0, "✅ Dentro", "❌ Fora"))
    if prev is not None:
        pp = prev["prod"].sum(axis=1)
        ps = prev["scrap"].sum(axis=1)
        pp.loc[REGION], ps.loc[REGION] = pp.sum(), ps.sum()
        df["PPM ano anterior"] = ps / pp * 1e6
        df["Δ PPM vs ano ant. (%)"] = pct(df["PPM ano (proj.)"], df["PPM ano anterior"])
        df["Δ Refugo vs ano ant. (%)"] = pct(df["Refugo ano"], ps)
        df["Δ Produção vs ano ant. (%)"] = pct(df["Produção ano"], pp)
    return df


def fmt_num(x, dec=0):
    if pd.isna(x):
        return ""
    return f"{x:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_pct(x):
    return "" if pd.isna(x) else f"{fmt_num(x, 1)}%"


def default_tables(data):
    w = pd.DataFrame(100.0, index=[MESES[m - 1] for m in data["projected"]],
                     columns=["Geral"] + data["plants"])
    w.index.name = "Mês"
    v = pd.DataFrame({"Variação refugo (%)": 0.0}, index=data["plants"])
    v.index.name = "Planta"
    return w, v


def apply_edits(default, key, none_is_nan=False):
    """Aplica as edições guardadas pelo st.data_editor (session_state[key]) à tabela padrão."""
    df = default.copy()
    state = st.session_state.get(key) or {}
    for r, changes in state.get("edited_rows", {}).items():
        for col, val in changes.items():
            j = df.columns.get_loc(col)
            if val is None:
                val = np.nan if none_is_nan else default.iat[int(r), j]
            df.iat[int(r), j] = val
    return df


def _diff_cells(edited, original):
    a, b = edited.astype(float), original.astype(float)
    return ~((a == b) | (a.isna() & b.isna()))


# --------------------------------------------------------------------------- sidebar
st.sidebar.title("⚙️ Configuração")
upload = st.sidebar.file_uploader("Arquivo de dados (layout 'Dados brutos')", type=["xlsx"])
if upload is not None:
    content = upload.getvalue()
elif DEFAULT_FILE.exists():
    content = DEFAULT_FILE.read_bytes()
else:
    st.error(f"Arquivo padrão não encontrado: {DEFAULT_FILE}. Envie um arquivo na barra lateral.")
    st.stop()

try:
    raw = load_data(content)
except Exception as exc:  # layout inesperado
    st.error(f"Não foi possível ler o arquivo: {exc}")
    st.stop()

plants = raw["plants"]

# Edições dos dados originais feitas na aba "Dados de entrada" (template), por ano
# O editor de um ano só é exibido quando o ano está selecionado, e o Streamlit descarta o estado de
# widgets não exibidos; por isso as edições ficam num armazenamento próprio: (tipo, ano) -> {(linha, col): valor}.
EDIT_YEARS = sorted(set(raw["years"]) | set(range(TEMPLATE_START, max(raw["years"] + [date.today().year]) + 1)))
FILE_ID = hashlib.md5(content).hexdigest()[:8]
STORE_KEY = f"raw_store_{FILE_ID}"
raw_store = st.session_state.setdefault(STORE_KEY, {})


def edited_table(default, kind, year):
    df = default.copy()
    for (r, col), val in raw_store.get((kind, year), {}).items():
        df.iat[r, df.columns.get_loc(col)] = np.nan if val is None else val
    return df


def raw_editor_key(kind, year):
    return f"raw_{kind}_{year}_{FILE_ID}_{st.session_state.get('raw_store_ver', 0)}"


def on_raw_edit(kind, year, key):
    store = st.session_state.setdefault(STORE_KEY, {})
    cells = store.setdefault((kind, year), {})
    for r, ch in st.session_state[key].get("edited_rows", {}).items():
        for col, val in ch.items():
            cells[(int(r), col)] = val
    st.session_state["raw_store_ver"] = st.session_state.get("raw_store_ver", 0) + 1


raw_default, raw_edit, raw_changed = {}, {}, {}
hist_prod, hist_scrap = {}, {}
for _y in EDIT_YEARS:
    _pd, _sd = year_tables(raw, _y)
    _pe = edited_table(_pd, "prod", _y)
    _se = edited_table(_sd, "scrap", _y)
    raw_default[_y] = {"prod": _pd, "scrap": _sd}
    raw_edit[_y] = {"prod": _pe, "scrap": _se}
    raw_changed[_y] = {"prod": _diff_cells(_pe, _pd), "scrap": _diff_cells(_se, _sd)}
    hist_prod[_y] = _pe.set_axis(range(1, 13), axis=1).astype(float)
    hist_scrap[_y] = _se.set_axis(range(1, 13), axis=1).astype(float)
raw_tg_default = targets_table(raw)
raw_tg_tbl = edited_table(raw_tg_default, "tg", 0)
raw_changed["tg"] = _diff_cells(raw_tg_tbl, raw_tg_default)
n_raw_changed = int(sum(raw_changed[y][k].values.sum() for y in EDIT_YEARS for k in ("prod", "scrap"))
                    + raw_changed["tg"].values.sum())
targets = {p: float(v) for p, v in raw_tg_tbl["Target PPM"].items() if pd.notna(v)}

# Anos sem nenhum dado são ignorados
data_years = [y for y in EDIT_YEARS if hist_prod[y].notna().any().any() or hist_scrap[y].notna().any().any()]
if not data_years:
    st.error("O arquivo não tem dados de produção/refugo. Preencha a aba Dados de entrada ou envie outro arquivo.")
    st.stop()

proj_year = st.sidebar.selectbox("📅 Ano da projeção", data_years, index=len(data_years) - 1, key="proj_year",
                                 help="Ano projetado (padrão: o último ano com dados). Os anos anteriores "
                                      "servem de histórico e de base de comparação.")
prev_year = proj_year - 1
data = build_data(plants, hist_prod[proj_year], hist_scrap[proj_year], targets, proj_year)
if not data["realized"]:
    st.error(f"{proj_year}: nenhum mês com produção e refugo completos para todas as plantas.")
    st.stop()
HAS_PREV = prev_year in data_years and bool(complete_months(hist_prod[prev_year], hist_scrap[prev_year]))
prev = {"prod": hist_prod[prev_year], "scrap": hist_scrap[prev_year]} if HAS_PREV else None
# Linha do tempo de meses completos até o ano da projeção (para média móvel / meses selecionados)
timeline = [(y, m) for y in data_years if y <= proj_year
            for m in complete_months(hist_prod[y], hist_scrap[y])]

if n_raw_changed:
    st.sidebar.warning(f"✏️ {n_raw_changed} valor(es) dos dados originais editado(s) no app "
                       "(aba Dados de entrada).")
HAS_PROJ = bool(data["projected"])

w_default, v_default = default_tables(data)
W_KEY = f"w_editor_{proj_year}_" + "_".join(map(str, data["projected"]))
V_KEY = f"v_editor_{proj_year}"
color_of = {p: CORES[i % len(CORES)] for i, p in enumerate(plants)}
real_labels = [MESES[m - 1] for m in data["realized"]]

st.sidebar.subheader("Produção base (meses projetados)")
methods = [M_ANO, M_MOVEL, M_SEL] + ([M_ANO_ANT, M_MESMO_MES] if HAS_PREV else [])
method = st.sidebar.radio("Metodologia", methods, key="method")
cur_periods = [p for p in timeline if p[0] == proj_year]
n_last = len(cur_periods)
sel_periods = cur_periods
if method == M_MOVEL:
    n_last = st.sidebar.slider("X (meses)", 1, len(timeline), min(3, len(timeline)), key="n_last",
                               help="Conta para trás a partir do último mês realizado; pode incluir meses "
                                    "dos anos anteriores.")
    st.sidebar.caption("Meses usados: " + ", ".join(plabel(*p) for p in timeline[-n_last:]))
elif method == M_SEL:
    opts = [p for p in timeline if p[0] >= prev_year]
    labels = [plabel(*p) for p in opts]
    chosen = st.sidebar.multiselect("Meses", labels, default=[plabel(*p) for p in cur_periods[-3:]],
                                    key="sel_months")
    sel_periods = [opts[labels.index(c)] for c in chosen]
    if not sel_periods:
        st.sidebar.warning("Nenhum mês selecionado — usando todos os realizados do ano.")
elif method == M_MESMO_MES:
    st.sidebar.caption(f"Cada mês projetado usa a produção do mesmo mês de {prev_year} "
                       "(sem dado → média do ano anterior).")

st.sidebar.subheader("PPM base do refugo")
ppm_opts = ["Média do ano", "Mesmo período da produção"] + (["Ano anterior"] if HAS_PREV else [])
ppm_method = st.sidebar.radio(
    "Período do PPM base", ppm_opts, key="ppm_method",
    help="Refugo projetado = PPM base × produção projetada × (1 + variação).",
)

if st.sidebar.button("↺ Restaurar padrões", width="stretch"):
    for k in list(st.session_state.keys()):
        del st.session_state[k]
    st.rerun()

prod_base, prod_cols = base_production(hist_prod, timeline, proj_year, data["projected"], plants,
                                       method, n_last, sel_periods)
if ppm_method == "Ano anterior":
    ppm_cols = [p for p in timeline if p[0] == prev_year]
elif ppm_method == "Mesmo período da produção":
    ppm_cols = prod_cols or cur_periods
else:
    ppm_cols = cur_periods
ppm_b = base_ppm(hist_prod, hist_scrap, ppm_cols, plants)

# Tabelas editáveis atuais (convertidas em frações)
w_df = apply_edits(w_default, W_KEY)
v_df = apply_edits(v_default, V_KEY)
weights = pd.DataFrame(
    {p: w_df["Geral"].values / 100 * w_df[p].values / 100 for p in plants},
    index=data["projected"],
)
variation = v_df["Variação refugo (%)"].reindex(plants).fillna(0) / 100

prod, scrap = project(data, prod_base, ppm_b, weights, variation)
summary = summarize(data, prod, scrap, prev)
region_target = data["targets"].get(REGION, np.nan)


def region_ppm_for(var_series):
    """PPM anual BU1 SA para um vetor de variações (fração) por planta."""
    _, s = project(data, prod_base, ppm_b, weights, var_series)
    return s.sum().sum() / prod.sum().sum() * 1e6


# Refugo projetado com variação 0 (base) — usado nas análises lineares
_, scrap0 = project(data, prod_base, ppm_b, weights, pd.Series(0.0, index=plants))
proj_scrap0 = scrap0[data["projected"]].sum(axis=1)
real_scrap = data["scrap"][data["realized"]].sum(axis=1)
prod_year = prod.sum(axis=1)
prev_ppm = summary["PPM ano anterior"] if HAS_PREV else pd.Series(np.nan, index=summary.index)

# --------------------------------------------------------------------------- explicações
prod_cols_txt = ", ".join(plabel(*p) for p in prod_cols)
ppm_cols_txt = ", ".join(plabel(*p) for p in ppm_cols)
proj_txt = ", ".join(MESES[m - 1] for m in data["projected"]) or "—"
real_txt = ", ".join(real_labels)

HELP = {
    # Resumo anual
    "Produção real": f"Σ peças produzidas nos meses realizados ({real_txt}). Fonte: conta 202 dos dados brutos.",
    "Refugo real": f"Σ peças refugadas nos meses realizados ({real_txt}). Fonte: conta 201 dos dados brutos.",
    "Produção projetada": f"Σ dos meses projetados ({proj_txt}) de: Produção base × Peso Geral % × Peso Planta %. "
                          f"Produção base = {method} ({prod_cols_txt}).",
    "Refugo projetado": "Σ dos meses projetados de: PPM base × Produção projetada do mês ÷ 1.000.000 × (1 + Variação refugo %). "
                        f"PPM base = Σ refugo ÷ Σ produção × 10⁶ em {ppm_cols_txt}.",
    "Produção ano": "Produção real + Produção projetada.",
    "Refugo ano": "Refugo real + Refugo projetado.",
    "PPM real": "Refugo real ÷ Produção real × 1.000.000 (somente meses realizados).",
    "PPM ano (proj.)": "Refugo ano ÷ Produção ano × 1.000.000. Na linha BU1 SA as quantidades são somadas antes de dividir "
                       "(média ponderada pela produção, não média simples dos PPMs).",
    "Target PPM": "Meta de PPM lida do bloco 'Targets' dos dados brutos.",
    "Gap vs target (%)": "(PPM ano (proj.) − Target PPM) ÷ Target PPM × 100. Negativo = abaixo da meta (bom).",
    "Peças p/ atingir target": "Refugo ano − Target PPM × Produção ano ÷ 10⁶. Positivo = peças refugadas a menos "
                               "necessárias para atingir o target; negativo = folga (peças que ainda cabem).",
    "Status": "✅ Dentro se Gap ≤ 0; ❌ Fora se Gap > 0.",
    "PPM ano anterior": f"PPM do ano {prev_year}: Σ refugo ÷ Σ produção × 10⁶ (meses com dados).",
    "Δ PPM vs ano ant. (%)": f"(PPM {proj_year} ÷ PPM {prev_year} − 1) × 100. Positivo = piora.",
    "Δ Refugo vs ano ant. (%)": f"(Refugo ano {proj_year} ÷ Refugo {prev_year} − 1) × 100.",
    "Δ Produção vs ano ant. (%)": f"(Produção ano {proj_year} ÷ Produção {prev_year} − 1) × 100.",
    "Total / Ano": "Produção e Refugo: soma de Jan–Dez. PPM: Refugo do ano ÷ Produção do ano × 10⁶.",
    # Sensibilidade
    "Refugo projetado base (total)": "Refugo projetado com variação 0%: PPM base × Produção projetada ÷ 10⁶, somado nos meses projetados.",
    "Refugo permitido (total meses proj.)": "Target PPM × Produção ano ÷ 10⁶ − Refugo real. Máximo de peças que ainda podem ser "
                                            "refugadas nos meses projetados para fechar o ano no target.",
    "Refugo permitido por mês": "Refugo permitido (total meses proj.) ÷ nº de meses projetados.",
    "Variação necessária vs base (%)": "(Refugo permitido ÷ Refugo projetado base − 1) × 100. Negativo = redução necessária.",
    "PPM projetado": "PPM ano (proj.) da planta no cenário atual.",
    # Combinações
    "Nº plantas": "Quantidade de plantas que recebem a mesma variação.",
    "Variação necessária (%)": "v tal que Σ Refugo real (todas) + Σ Refugo proj. das demais (cenário atual) + "
                               "Σ Refugo proj. base da combinação × (1 + v) = Target BU1 SA × Produção BU1 SA ÷ 10⁶.",
    "Peças a reduzir (meses proj.)": "−v × Σ Refugo projetado base das plantas da combinação.",
    "Viável": "✅ se a variação necessária ≥ −100% (não exige refugo negativo).",
    # Situação atual / simulador
    "PPM": "Refugo ÷ Produção × 1.000.000.",
    "Desvio PPM": "PPM − Target PPM. Positivo = acima da meta.",
    "Desvio %": "(PPM − Target) ÷ Target × 100.",
    "Desvio meta (pç)": "Refugo − Target PPM × Produção ÷ 10⁶. Positivo = peças acima da meta; negativo = folga.",
    "Desvio vs ano ant. (pç)": "Refugo atual − refugo do mesmo período do ano anterior. Positivo = mais peças refugadas.",
    "Refugo ano": "Refugo real + Refugo projetado.",
    "Refugo meses proj.": "Refugo ano − Refugo real: peças a refugar nos meses projetados.",
    "Refugo médio/mês proj.": "Refugo meses proj. ÷ nº de meses projetados.",
    "Variação vs base (%)": "(Refugo meses proj. ÷ Refugo projetado base − 1) × 100.",
}
for _m in range(1, 13):
    HELP[MESES[_m - 1]] = (f"{MESES[_m - 1]}: valor realizado (dados brutos)." if _m in data["realized"] else
                           f"{MESES[_m - 1]} projetado: Produção = base × pesos; Refugo = PPM base × Produção ÷ 10⁶ × "
                           "(1 + variação); PPM = Refugo ÷ Produção × 10⁶.")


def show_table(obj, title=None, extra_help=None, **kwargs):
    """st.dataframe com dica (tooltip) em cada coluna calculada + popover com o memorial da tabela."""
    df = obj.data if hasattr(obj, "data") else obj
    helps = {**HELP, **(extra_help or {})}
    cols = [c for c in df.columns if c in helps]
    cfg = {c: st.column_config.Column(help=helps[c]) for c in cols}
    st.dataframe(obj, column_config=cfg, **kwargs)
    with st.popover("ℹ️ Memorial de cálculo" + (f" — {title}" if title else "")):
        st.caption("Passe o mouse sobre o cabeçalho de cada coluna para ver a mesma explicação.")
        seen = set()
        for c in cols:
            if c in MESES:
                continue
            if c not in seen:
                st.markdown(f"**{c}** — {helps[c]}")
                seen.add(c)
        if any(c in MESES for c in df.columns):
            st.markdown(f"**Meses realizados ({real_txt})** — valores reais dos dados brutos.")
            st.markdown(f"**Meses projetados ({proj_txt})** — {HELP[MESES[data['projected'][0] - 1]].split(': ', 1)[1]}"
                        if data["projected"] else "")


# --------------------------------------------------------------------------- cabeçalho
st.title(f"📉 Projeção de Refugo {proj_year} — {REGION}")
st.caption(
    f"Realizado: {', '.join(real_labels)} · Projetado: "
    f"{', '.join(MESES[m - 1] for m in data['projected']) or '—'} · "
    f"Produção base: {method} ({prod_cols_txt})"
    + (f" · Comparação: {prev_year}" if HAS_PREV else " · Sem dados do ano anterior")
)

reg = summary.loc[REGION]
k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("PPM projetado BU1 SA", fmt_num(reg["PPM ano (proj.)"]),
          delta=f"{fmt_pct(reg['Gap vs target (%)'])} vs target", delta_color="inverse",
          help=HELP["Gap vs target (%)"])
k2.metric("Target BU1 SA", fmt_num(region_target))
k3.metric("PPM realizado (até agora)", fmt_num(reg["PPM real"]))
k4.metric("Plantas dentro do target",
          f"{(summary.drop(REGION)['Status'] == '✅ Dentro').sum()} / {len(plants)}")
if HAS_PREV:
    k5.metric(f"PPM BU1 SA {prev_year}", fmt_num(reg["PPM ano anterior"]),
              delta=f"{fmt_pct(reg['Δ PPM vs ano ant. (%)'])} projetado vs {prev_year}", delta_color="inverse",
              help=HELP["Δ PPM vs ano ant. (%)"])
else:
    k5.metric("Ano anterior", "—", help="Sem dados do ano anterior no arquivo.")

tab0, tabH, tab1, tab2, tab3, tab4, tab6, tab5 = st.tabs([
    "🏠 Situação atual", "📚 Resumo histórico", "📋 Produzido × Refugo", "📈 Sensibilidade por planta",
    "🧩 Combinações BU1 SA vs Target", "🎯 Simulador de combinações", "✏️ Dados de entrada",
    "📘 Memorial de cálculo"])


def actual_frame(prod_df, scrap_df, months):
    """Produção, refugo, PPM e desvio vs target (dados reais) para um conjunto de meses."""
    pr = prod_df[months].sum(axis=1, min_count=1)
    sc = scrap_df[months].sum(axis=1, min_count=1)
    df = pd.DataFrame({"Produção": pr, "Refugo": sc})
    df.loc[REGION] = df.sum(min_count=len(plants))
    df["PPM"] = df["Refugo"] / df["Produção"] * 1e6
    df["Target PPM"] = [data["targets"].get(p, np.nan) for p in df.index]
    df["Desvio PPM"] = df["PPM"] - df["Target PPM"]
    df["Desvio %"] = df["Desvio PPM"] / df["Target PPM"] * 100
    df["Desvio meta (pç)"] = df["Refugo"] - df["Target PPM"] * df["Produção"] / 1e6
    df["Status"] = np.where(df["Desvio PPM"] <= 0, "✅ Dentro", "❌ Fora")
    return df


def colored_pct(d, higher_is_bad=True):
    """Variação % colorida para st.caption/markdown (vermelho = piora)."""
    if pd.isna(d):
        return "—"
    bad = d > 0 if higher_is_bad else d < 0
    return f":{'red' if bad else 'green'}[{'▲' if d > 0 else '▼'} {fmt_pct(d)}]"


def status_cards(df, prev_df=None, prev_lbl="", cols_per_row=3):
    names = [REGION] + plants
    for i in range(0, len(names), cols_per_row):
        cols = st.columns(cols_per_row)
        for col, p in zip(cols, names[i:i + cols_per_row]):
            r = df.loc[p]
            with col.container(border=True):
                st.metric(
                    ("🌎 " if p == REGION else "🏭 ") + p, fmt_num(r["PPM"]) + " PPM",
                    delta=f"{fmt_num(r['Desvio PPM'])} ({fmt_num(r['Desvio %'], 1)}%) vs meta",
                    delta_color="inverse",
                    help=f"PPM = Refugo ÷ Produção × 10⁶ = {fmt_num(r['Refugo'])} ÷ {fmt_num(r['Produção'])} × 10⁶. "
                         f"Desvio = PPM − Target ({fmt_num(r['Target PPM'])}).",
                )
                st.caption(f"{r['Status']} · Meta {fmt_num(r['Target PPM'])} · "
                           f"Refugo {fmt_num(r['Refugo'])} pç · Produção {fmt_num(r['Produção'])} pç")
                desvio_meta = r["Desvio meta (pç)"]
                st.caption(f"**Desvio da meta:** {fmt_num(desvio_meta)} pç "
                           f"({'acima' if desvio_meta > 0 else 'de folga' if desvio_meta < 0 else 'na meta'})")
                if prev_df is not None and pd.notna(prev_df.loc[p, "PPM"]):
                    pv = prev_df.loc[p]
                    delta_refugo_pecas = r["Refugo"] - pv["Refugo"]
                    st.caption(f"**{prev_lbl}:** {fmt_num(pv['PPM'])} PPM → PPM {colored_pct(pct(r['PPM'], pv['PPM']))}"
                               f" · refugo {colored_pct(pct(r['Refugo'], pv['Refugo']))}"
                               f" · produção {colored_pct(pct(r['Produção'], pv['Produção']), higher_is_bad=False)}")
                    st.caption(f"**Desvio vs {prev_lbl}:** {fmt_num(delta_refugo_pecas)} pç de refugo")


def entity_series(prod_df, scrap_df, months, p):
    pr, sc = prod_df[months], scrap_df[months]
    if p == REGION:
        return pr.sum(min_count=len(plants)), sc.sum(min_count=len(plants))
    return pr.loc[p], sc.loc[p]


# =========================================================================== TELA 0
with tab0:
    sel_label = st.selectbox(f"Mês de referência ({proj_year})", real_labels, index=len(real_labels) - 1,
                             key="cur_month")
    cur_m = data["realized"][real_labels.index(sel_label)]
    ytd_months = [m for m in data["realized"] if m <= cur_m]
    cur_df = actual_frame(data["prod"], data["scrap"], [cur_m])
    ytd_df = actual_frame(data["prod"], data["scrap"], ytd_months)
    prev_cur = actual_frame(prev["prod"], prev["scrap"], [cur_m]) if HAS_PREV else None
    prev_ytd = actual_frame(prev["prod"], prev["scrap"], ytd_months) if HAS_PREV else None

    st.subheader(f"Refugo do mês — {sel_label}/{proj_year}")
    if HAS_PREV:
        st.caption(f"Comparação com o mesmo mês de {prev_year}. Vermelho = piora (PPM/refugo maior ou produção menor).")
    status_cards(cur_df, prev_cur, f"{sel_label}/{prev_year}")
    st.subheader(f"Acumulado no ano — Jan a {sel_label}/{proj_year}")
    if HAS_PREV:
        st.caption(f"Comparação com o mesmo intervalo de {prev_year} (Jan a {sel_label}).")
    status_cards(ytd_df, prev_ytd, f"Jan–{sel_label}/{prev_year}")

    st.subheader("Histórico mensal")
    h1, h2, h3, h4 = st.columns([1, 1, 2, 1])
    ind = h1.radio("Indicador", ["PPM", "Peças refugadas", "Produção"], key="hist_ind", horizontal=True)
    view = h2.radio("Visão", ["Região BU1 SA", "Por planta"], key="hist_view", horizontal=True)
    hist_plants = h3.multiselect("Plantas", plants, default=plants, key="hist_plants",
                                 disabled=view != "Por planta")
    cmp_prev = h4.toggle(f"Comparar com {prev_year}", value=True, key="hist_cmp", disabled=not HAS_PREV) and HAS_PREV

    months = data["realized"]
    x = [MESES[m - 1] for m in months]

    def series_for(p, src_prod, src_scrap):
        pr, sc = entity_series(src_prod, src_scrap, months, p)
        return {"PPM": sc / pr * 1e6, "Peças refugadas": sc, "Produção": pr}[ind]

    fig = go.Figure()
    show_prev = cmp_prev and (view == "Região BU1 SA" or len(hist_plants) == 1)
    targets_line = REGION if view == "Região BU1 SA" else (hist_plants[0] if len(hist_plants) == 1 else None)
    if view == "Região BU1 SA" or len(hist_plants) == 1:
        ent = REGION if view == "Região BU1 SA" else hist_plants[0]
        base_color = CORES[0] if ent == REGION else color_of[ent]
        if show_prev:
            yp = series_for(ent, prev["prod"], prev["scrap"])
            fig.add_bar(x=x, y=yp.values, name=f"{ent} {prev_year}", marker_color="#b8b7b0",
                        hovertemplate=f"{prev_year} · %{{x}}: %{{y:,.0f}}<extra></extra>")
        y = series_for(ent, data["prod"], data["scrap"])
        fig.add_bar(x=x, y=y.values, name=f"{ent} {proj_year}", marker_color=base_color,
                    text=[fmt_num(v) for v in y.values], textposition="outside",
                    hovertemplate=f"{proj_year} · %{{x}}: %{{y:,.0f}}<extra></extra>")
    else:
        for p in hist_plants:
            y = series_for(p, data["prod"], data["scrap"])
            fig.add_bar(x=x, y=y.values, name=p, marker_color=color_of[p],
                        hovertemplate=f"{p}<br>%{{x}}: %{{y:,.0f}}<extra></extra>")
        if cmp_prev:
            st.caption(f"A comparação com {prev_year} aparece na visão Região ou com 1 planta selecionada.")
    if ind == "PPM" and targets_line:
        t = data["targets"].get(targets_line, np.nan)
        fig.add_hline(y=t, line_dash="dash", line_color=COR_TARGET,
                      annotation_text=f"Target {fmt_num(t)}", annotation_position="top left")
    fig.update_layout(barmode="group", bargap=0.2, bargroupgap=0.05, height=440, yaxis_title=ind,
                      title=f"{ind} por mês" + (" — linha tracejada = target" if ind == "PPM" else ""),
                      legend=dict(orientation="h", y=-0.15), separators=",.")
    fig.update_yaxes(tickformat=",.0f")
    st.plotly_chart(fig, width="stretch")

    st.subheader("Desvio mensal do PPM vs meta (%)")
    dev_rows = {}
    for p in plants + [REGION]:
        pr, sc = entity_series(data["prod"], data["scrap"], months, p)
        tg = data["targets"].get(p, np.nan)
        dev_rows[p] = ((sc / pr * 1e6 - tg) / tg * 100).values
    dev = pd.DataFrame(dev_rows, index=x).T
    dev.index.name = "Planta"
    show_table(
        dev.style.format(fmt_pct).map(lambda v: "color:#c62828" if v > 0 else "color:#2e7d32"),
        title="Desvio mensal",
        extra_help={m: f"{m}: (PPM do mês − Target PPM) ÷ Target PPM × 100, com PPM do mês = Refugo ÷ Produção × 10⁶. "
                       "Positivo (vermelho) = acima da meta." for m in x},
        width="stretch",
    )

    if HAS_PREV:
        st.subheader(f"Variação do PPM vs mesmo mês de {prev_year} (%)")
        yoy_rows = {}
        for p in plants + [REGION]:
            pr, sc = entity_series(data["prod"], data["scrap"], months, p)
            ppr, psc = entity_series(prev["prod"], prev["scrap"], months, p)
            vals = list(pct(sc / pr, psc / ppr).values)
            vals.append(pct(sc.sum() / pr.sum(), psc.sum(min_count=1) / ppr.sum(min_count=1)))
            yoy_rows[p] = vals
        yoy = pd.DataFrame(yoy_rows, index=x + ["Acumulado"]).T
        yoy.index.name = "Planta"
        show_table(
            yoy.style.format(fmt_pct).map(lambda v: "color:#c62828" if v > 0 else "color:#2e7d32"),
            title=f"Variação vs {prev_year}",
            extra_help={**{m: f"{m}: (PPM {m}/{proj_year} ÷ PPM {m}/{prev_year} − 1) × 100. Positivo (vermelho) = piora."
                           for m in x},
                        "Acumulado": f"(PPM Jan–{x[-1]}/{proj_year} ÷ PPM Jan–{x[-1]}/{prev_year} − 1) × 100."},
            width="stretch",
        )

    st.subheader("Tabela do mês selecionado")
    month_tbl = cur_df.copy()
    if HAS_PREV:
        month_tbl[f"PPM {sel_label}/{prev_year}"] = prev_cur["PPM"]
        month_tbl["Δ PPM vs ano ant. (%)"] = pct(cur_df["PPM"], prev_cur["PPM"])
        month_tbl[f"Refugo {sel_label}/{prev_year}"] = prev_cur["Refugo"]
        month_tbl["Desvio vs ano ant. (pç)"] = cur_df["Refugo"] - prev_cur["Refugo"]
    show_table(month_tbl.style.format({c: (fmt_pct if "%" in c else fmt_num)
                                       for c in month_tbl.columns if c != "Status"}),
               title="Mês selecionado",
               extra_help={"Produção": "Peças produzidas no mês (conta 202).",
                           "Refugo": "Peças refugadas no mês (conta 201).",
                           "Desvio meta (pç)": "Refugo do mês − Target PPM × Produção do mês ÷ 1.000.000. Positivo = peças acima da meta; negativo = folga.",
                           f"PPM {sel_label}/{prev_year}": "PPM do mesmo mês no ano anterior.",
                           "Δ PPM vs ano ant. (%)": "(PPM do mês ÷ PPM do mesmo mês do ano anterior − 1) × 100.",
                           f"Refugo {sel_label}/{prev_year}": "Peças refugadas no mesmo mês do ano anterior.",
                           "Desvio vs ano ant. (pç)": "Refugo do mês selecionado − refugo do mesmo mês do ano anterior. Positivo = mais peças refugadas."},
               width="stretch")

# =========================================================================== RESUMO HISTÓRICO
with tabH:
    st.caption(f"Anos com dados: {', '.join(map(str, data_years))} (anos sem dados são ignorados). "
               "Ano incompleto = soma dos meses disponíveis. Target = meta atual do bloco Targets.")
    g1, g2, g3 = st.columns([1, 1, 2])
    h_ind = g1.radio("Indicador", ["PPM", "Peças refugadas", "Produção"], key="hh_ind", horizontal=True)
    h_ent = g2.selectbox("Planta / região", [REGION] + plants, key="hh_ent")
    h_years = sorted(g3.multiselect("Anos", data_years, default=data_years, key="hh_years"))
    year_color = {y: CORES[i % len(CORES)] for i, y in enumerate(data_years)}
    all_m = list(range(1, 13))

    def ind_values(pr, sc):
        return {"PPM": sc / pr * 1e6, "Peças refugadas": sc, "Produção": pr}[h_ind]

    def months_with_data(y, p):
        pr, sc = entity_series(hist_prod[y], hist_scrap[y], all_m, p)
        return [m for m in all_m if pd.notna(pr[m]) and pd.notna(sc[m])]

    def total_value(y, p, months_sel):
        pr, sc = entity_series(hist_prod[y], hist_scrap[y], all_m, p)
        if not months_sel:
            return np.nan
        prs, scs = pr[months_sel].sum(min_count=len(months_sel)), sc[months_sel].sum(min_count=len(months_sel))
        return {"PPM": scs / prs * 1e6, "Peças refugadas": scs, "Produção": prs}[h_ind]

    if not h_years:
        st.info("Selecione ao menos um ano.")
    else:
        # Linha do tempo
        xs, ys = [], []
        for y in h_years:
            v = ind_values(*entity_series(hist_prod[y], hist_scrap[y], all_m, h_ent))
            for m in all_m:
                if pd.notna(v[m]):
                    xs.append(pd.Timestamp(y, m, 1))
                    ys.append(v[m])
        fig = go.Figure()
        fig.add_scatter(x=xs, y=ys, mode="lines+markers", name=h_ent,
                        line=dict(color=CORES[0] if h_ent == REGION else color_of[h_ent], width=2),
                        marker=dict(size=8), hovertemplate="%{x|%b/%Y}: %{y:,.0f}<extra></extra>")
        t = data["targets"].get(h_ent, np.nan)
        if h_ind == "PPM" and pd.notna(t):
            fig.add_hline(y=t, line_dash="dash", line_color=COR_TARGET,
                          annotation_text=f"Target atual {fmt_num(t)}", annotation_position="top left")
        fig.update_layout(title=f"{h_ind} mensal — {h_ent} (linha do tempo)", height=400, yaxis_title=h_ind,
                          separators=",.", showlegend=False)
        fig.update_xaxes(tickformat="%b/%y")
        fig.update_yaxes(tickformat=",.0f")
        st.plotly_chart(fig, width="stretch")

        # Comparação ano a ano
        fig2 = go.Figure()
        for y in h_years:
            v = ind_values(*entity_series(hist_prod[y], hist_scrap[y], all_m, h_ent))
            fig2.add_scatter(x=MESES, y=v.values, mode="lines+markers", name=str(y),
                             line=dict(color=year_color[y], width=3 if y == proj_year else 2),
                             marker=dict(size=8), hovertemplate=f"{y} · %{{x}}: %{{y:,.0f}}<extra></extra>")
        if h_ind == "PPM" and pd.notna(t):
            fig2.add_hline(y=t, line_dash="dash", line_color=COR_TARGET,
                           annotation_text=f"Target atual {fmt_num(t)}", annotation_position="top left")
        fig2.update_layout(title=f"{h_ind} por mês — comparação entre anos ({h_ent})", height=420,
                           yaxis_title=h_ind, hovermode="x unified", separators=",.",
                           legend=dict(orientation="h", y=-0.15))
        fig2.update_yaxes(tickformat=",.0f")
        st.plotly_chart(fig2, width="stretch")

        # Tabela anual e variação ano a ano (mesmo intervalo de meses)
        ents = plants + [REGION]
        annual, yoy, n_months = {}, {}, {}
        for y in h_years:
            n_months[y] = len(months_with_data(y, REGION))
            col = f"{y}" + (f" ({n_months[y]} m)" if n_months[y] < 12 else "")
            annual[col] = [total_value(y, p, months_with_data(y, p)) for p in ents]
            if y - 1 in data_years:
                vals = []
                for p in ents:
                    ms = [m for m in months_with_data(y, p) if m in months_with_data(y - 1, p)]
                    vals.append(pct(total_value(y, p, ms), total_value(y - 1, p, ms)) if ms else np.nan)
                yoy[f"{y} vs {y - 1}"] = vals
        annual_df = pd.DataFrame(annual, index=ents)
        annual_df.index.name = "Planta"
        st.subheader(f"{h_ind} anual por planta")
        show_table(annual_df.style.format(fmt_num), title=f"{h_ind} anual", width="stretch",
                   extra_help={c: f"{h_ind} do ano {c[:4]}" + (": soma dos meses disponíveis" if "m)" in c else "")
                               + (" (Σ refugo ÷ Σ produção × 10⁶)." if h_ind == "PPM" else ".")
                               for c in annual_df.columns})
        if yoy:
            yoy_df = pd.DataFrame(yoy, index=ents)
            yoy_df.index.name = "Planta"
            st.subheader(f"Variação anual do {h_ind} (%)")
            show_table(yoy_df.style.format(fmt_pct).map(
                lambda v: ("color:#c62828" if v > 0 else "color:#2e7d32") if h_ind != "Produção" else ""),
                title="Variação anual", width="stretch",
                extra_help={c: f"({h_ind} {c[:4]} ÷ {h_ind} {c[-4:]} − 1) × 100, usando apenas os meses presentes "
                               "nos dois anos (mesmo intervalo)." for c in yoy_df.columns})

        # Matriz mês × ano
        st.subheader(f"{h_ind} mês a mês — {h_ent}")
        mat = pd.DataFrame({MESES[m - 1]: [ind_values(*entity_series(hist_prod[y], hist_scrap[y], all_m, h_ent))[m]
                                           for y in h_years] for m in all_m}, index=[str(y) for y in h_years])
        mat["Ano"] = [total_value(y, h_ent, months_with_data(y, h_ent)) for y in h_years]
        mat.index.name = "Ano"
        show_table(mat.style.format(fmt_num), title="Mês a mês", width="stretch",
                   extra_help={**{m: f"{h_ind} de {m} no ano da linha (dados brutos)." for m in MESES},
                               "Ano": f"{h_ind} do ano (meses disponíveis)."})

        # Sazonalidade da produção
        st.subheader(f"Sazonalidade da produção — {h_ent} (% da média mensal do ano)")
        st.caption("Referência para a tabela de pesos da produção projetada (tela Produzido × Refugo).")
        seas = {}
        for y in h_years:
            pr, _ = entity_series(hist_prod[y], hist_scrap[y], all_m, h_ent)
            if pr.notna().sum() >= 2:
                seas[str(y)] = (pr / pr.mean() * 100).values
        if seas:
            seas_df = pd.DataFrame(seas, index=MESES).T
            full = [str(y) for y in h_years if months_with_data(y, h_ent) == all_m]
            if full:
                seas_df.loc["Média anos completos"] = seas_df.loc[full].mean()
            seas_df.index.name = "Ano"
            show_table(seas_df.style.format(fmt_pct), title="Sazonalidade", width="stretch",
                       extra_help={m: f"Produção de {m} ÷ média mensal da produção do ano (meses disponíveis) × 100."
                                   for m in MESES})

# =========================================================================== TELA 1
with tab1:
    c1, c2 = st.columns([3, 2])
    with c1:
        st.subheader("Pesos da produção projetada (% da base)")
        st.caption("Peso efetivo = Geral × Planta. Ex.: Dez = 30% ⇒ produção 70% menor que a base. "
                   "O refugo acompanha a produção (PPM base × produção).")
        cfg = {c: st.column_config.NumberColumn(c, min_value=0.0, max_value=500.0, step=5.0,
                                                format="%.0f%%")
               for c in w_default.columns}
        st.data_editor(w_default, column_config=cfg, width="stretch", key=W_KEY)
    with c2:
        st.subheader("Variação do refugo projetado")
        st.caption("Aplicada sobre o refugo base dos meses projetados. −10% = 10% menos peças refugadas.")
        st.data_editor(
            v_default,
            column_config={"Variação refugo (%)": st.column_config.NumberColumn(
                min_value=-100.0, max_value=500.0, step=1.0, format="%.1f%%")},
            width="stretch", key=V_KEY,
        )

    st.subheader("Resumo anual por planta")
    sum_show = summary.copy()
    num_cols = [c for c in sum_show.columns if c != "Status"]
    show_table(
        sum_show.style.format({c: (fmt_pct if "%" in c else fmt_num) for c in num_cols})
        .map(lambda v: "color:#c62828" if v > 0 else "color:#2e7d32",
             subset=[c for c in ["Gap vs target (%)", "Peças p/ atingir target", "Δ PPM vs ano ant. (%)"]
                     if c in sum_show.columns])
        .apply(lambda r: ["font-weight:600" if r.name == REGION else "" for _ in r], axis=1),
        title="Resumo anual", width="stretch",
    )

    st.subheader("Tabela mensal — Produção, Refugo e PPM")
    st.caption("Colunas destacadas = meses projetados.")
    rows = []
    for p in plants + [REGION]:
        pr = prod.loc[p] if p != REGION else prod.sum()
        sc = scrap.loc[p] if p != REGION else scrap.sum()
        ppm = sc / pr * 1e6
        items = [("Produção", pr, pr.sum()), ("Refugo", sc, sc.sum()), ("PPM", ppm, sc.sum() / pr.sum() * 1e6)]
        if HAS_PREV:
            ppr, psc = entity_series(prev["prod"], prev["scrap"], list(range(1, 13)), p)
            pppm = psc / ppr * 1e6
            pppm_tot = psc.sum(min_count=1) / ppr.sum(min_count=1) * 1e6
            items += [(f"PPM {prev_year}", pppm, pppm_tot),
                      ("Δ PPM vs ano ant. (%)", pct(ppm, pppm), pct(sc.sum() / pr.sum() * 1e6, pppm_tot))]
        for name, ser, total in items:
            row = {"Planta": p, "Indicador": name}
            row.update({MESES[m - 1]: ser[m] for m in range(1, 13)})
            row["Total / Ano"] = total
            rows.append(row)
    monthly = pd.DataFrame(rows).set_index(["Planta", "Indicador"])
    proj_labels = [MESES[m - 1] for m in data["projected"]]
    pct_rows = [i for i in monthly.index if "%" in i[1]]
    show_table(
        monthly.style.format(fmt_num).format(fmt_pct, subset=(pct_rows, monthly.columns))
        .set_properties(subset=proj_labels, **{"background-color": "rgba(42,120,214,0.12)",
                                               "font-style": "italic"}),
        title="Tabela mensal", width="stretch", height=min(38 * (len(monthly) + 1), 900),
        extra_help={"Total / Ano": "Produção e Refugo: soma de Jan–Dez. PPM: Refugo do ano ÷ Produção do ano × 10⁶. "
                                   f"Δ PPM vs ano ant.: (PPM {proj_year} ÷ PPM {prev_year} − 1) × 100."},
    )

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        summary.to_excel(xw, sheet_name="Resumo")
        monthly.to_excel(xw, sheet_name="Mensal")
        w_df.to_excel(xw, sheet_name="Pesos produção")
        v_df.to_excel(xw, sheet_name="Variação refugo")
    st.download_button("⬇️ Baixar projeção (Excel)", buf.getvalue(), "projecao_refugo.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # Gráfico PPM mensal
    fig = go.Figure()
    for p in plants:
        fig.add_scatter(x=MESES, y=(scrap.loc[p] / prod.loc[p] * 1e6).values, name=p,
                        mode="lines+markers", line=dict(color=color_of[p], width=2),
                        marker=dict(size=8))
    fig.add_scatter(x=MESES, y=(scrap.sum() / prod.sum() * 1e6).values, name=f"{REGION} {proj_year}",
                    mode="lines+markers", line=dict(color=COR_REGIAO, width=3), marker=dict(size=9))
    if HAS_PREV:
        ppr, psc = entity_series(prev["prod"], prev["scrap"], list(range(1, 13)), REGION)
        fig.add_scatter(x=MESES, y=(psc / ppr * 1e6).values, name=f"{REGION} {prev_year}", mode="lines+markers",
                        line=dict(color=COR_TARGET, width=2, dash="dash"), marker=dict(size=7))
    if data["projected"]:
        fig.add_vrect(x0=MESES[data["projected"][0] - 1], x1=MESES[-1], fillcolor="#8888",
                      opacity=0.12, line_width=0, annotation_text="Projetado",
                      annotation_position="top left")
    fig.update_layout(title="PPM mensal (real + projetado)", yaxis_title="PPM",
                      hovermode="x unified", height=420, legend=dict(orientation="h", y=-0.15))
    st.plotly_chart(fig, width="stretch")

# =========================================================================== TELA 2
with tab2:
    if not HAS_PROJ:
        st.info(f"Não há meses projetados em {proj_year} — análise disponível apenas para anos em andamento.")
    else:
        cA, cB = st.columns([1, 3])
        with cA:
            plant = st.selectbox("Planta", plants, key="sens_plant")
            rng = st.slider("Faixa de variação do refugo (%)", -100, 200, (-50, 50), step=5,
                            key="sens_range")
            n_proj = max(len(data["projected"]), 1)
            tgt = data["targets"].get(plant, np.nan)
            prod_proj_p = prod.loc[plant, data["projected"]].sum()
            allowed_total = tgt * prod_year[plant] / 1e6 - real_scrap[plant]
            req_var = allowed_total / proj_scrap0[plant] - 1 if proj_scrap0[plant] else np.nan

            st.metric("PPM ano projetado", fmt_num(summary.loc[plant, "PPM ano (proj.)"]),
                      delta=(f"{fmt_pct(summary.loc[plant, 'Δ PPM vs ano ant. (%)'])} vs {prev_year} "
                             f"({fmt_num(prev_ppm[plant])})" if HAS_PREV else None), delta_color="inverse")
            st.metric("Produção ano projetada", fmt_num(prod_year[plant]),
                      delta=(f"{fmt_pct(summary.loc[plant, 'Δ Produção vs ano ant. (%)'])} vs {prev_year}"
                             if HAS_PREV else None), delta_color="off")
            st.metric("Produção meses projetados", fmt_num(prod_proj_p))
            st.metric("Refugo médio mensal (ano real)",
                      fmt_num(data["scrap"].loc[plant, data["realized"]].mean()))
            st.metric("Refugo médio mensal projetado (base)", fmt_num(proj_scrap0[plant] / n_proj))
            st.metric("Máx. refugo/mês p/ atingir target", fmt_num(allowed_total / n_proj),
                      delta=(f"{fmt_num(req_var * 100, 1)}% vs base" if pd.notna(req_var) else None),
                      delta_color="off")
            if pd.notna(req_var) and allowed_total < 0:
                st.error("Target inatingível este ano: o refugo já realizado excede o permitido.")

        with cB:
            vs = np.linspace(rng[0], rng[1], 81) / 100
            scrap_month = proj_scrap0[plant] * (1 + vs) / n_proj
            ppm_plant = (real_scrap[plant] + proj_scrap0[plant] * (1 + vs)) / prod_year[plant] * 1e6
            base_var = variation.copy()
            ppm_region = []
            for v in vs:
                base_var[plant] = v
                ppm_region.append(region_ppm_for(base_var))
            hover = [f"Variação: {v * 100:+.0f}%" for v in vs]

            fig = go.Figure()
            fig.add_scatter(x=scrap_month, y=ppm_plant, name=f"PPM ano — {plant}", mode="lines",
                            line=dict(color=color_of[plant], width=2), customdata=hover,
                            hovertemplate="Refugo/mês: %{x:,.0f}<br>PPM: %{y:,.0f}<br>%{customdata}<extra></extra>")
            x0 = proj_scrap0[plant] * (1 + variation[plant]) / n_proj
            y0 = summary.loc[plant, "PPM ano (proj.)"]
            fig.add_scatter(x=[x0], y=[y0], name="Cenário atual", mode="markers",
                            marker=dict(size=12, color=color_of[plant], line=dict(color="white", width=2)))
            if pd.notna(tgt):
                fig.add_hline(y=tgt, line_dash="dash", line_color=COR_TARGET,
                              annotation_text=f"Target {fmt_num(tgt)}", annotation_position="top left")
            if HAS_PREV and pd.notna(prev_ppm[plant]):
                fig.add_hline(y=prev_ppm[plant], line_dash="dot", line_color="#8a8984",
                              annotation_text=f"PPM {prev_year} {fmt_num(prev_ppm[plant])}",
                              annotation_position="bottom right")
            fig.update_layout(title=f"{plant}: PPM anual × refugo médio mensal projetado",
                              xaxis_title="Peças refugadas por mês projetado", yaxis_title="PPM ano",
                              height=420, hovermode="closest", legend=dict(orientation="h", y=-0.2))
            st.plotly_chart(fig, width="stretch")

            fig2 = go.Figure()
            fig2.add_scatter(x=scrap_month, y=ppm_region, name=f"PPM {REGION}", mode="lines",
                             line=dict(color=COR_REGIAO, width=2), customdata=hover,
                             hovertemplate="Refugo/mês planta: %{x:,.0f}<br>PPM região: %{y:,.0f}<br>%{customdata}<extra></extra>")
            if pd.notna(region_target):
                fig2.add_hline(y=region_target, line_dash="dash", line_color=COR_TARGET,
                               annotation_text=f"Target {REGION} {fmt_num(region_target)}",
                               annotation_position="top left")
            fig2.update_layout(title=f"Impacto de {plant} no PPM {REGION} (demais plantas no cenário atual)",
                               xaxis_title=f"Peças refugadas por mês projetado — {plant}",
                               yaxis_title=f"PPM {REGION}", height=380, showlegend=False)
            st.plotly_chart(fig2, width="stretch")

        st.subheader("Refugo permitido por planta para atingir o próprio target")
        rows = []
        for p in plants:
            t = data["targets"].get(p, np.nan)
            allowed = t * prod_year[p] / 1e6 - real_scrap[p]
            rows.append({
                "Planta": p, "Target PPM": t, "PPM projetado": summary.loc[p, "PPM ano (proj.)"],
                "Refugo projetado base (total)": proj_scrap0[p],
                "Refugo permitido (total meses proj.)": allowed,
                "Refugo permitido por mês": allowed / n_proj,
                "Variação necessária vs base (%)": (allowed / proj_scrap0[p] - 1) * 100 if proj_scrap0[p] else np.nan,
            })
        show_table(pd.DataFrame(rows).set_index("Planta").style.format(fmt_num),
                   title="Refugo permitido", width="stretch")

# =========================================================================== TELA 3
with tab3:
    if not HAS_PROJ:
        st.info(f"Não há meses projetados em {proj_year} — análise disponível apenas para anos em andamento.")
    else:
        st.caption("Demais plantas (não analisadas em cada gráfico) usam a variação definida na "
                   "tabela da tela 'Produzido × Refugo'.")
        total_prod = prod.sum().sum()
        allowed_region = region_target * total_prod / 1e6

        # Gráfico 1 — influência de cada planta
        vs = np.linspace(-0.5, 0.5, 41)
        fig = go.Figure()
        for p in plants:
            ys = []
            for v in vs:
                vv = variation.copy()
                vv[p] = v
                ys.append(region_ppm_for(vv))
            fig.add_scatter(x=vs * 100, y=ys, name=p, mode="lines", line=dict(color=color_of[p], width=2))
        fig.add_hline(y=region_target, line_dash="dash", line_color=COR_TARGET,
                      annotation_text=f"Target {fmt_num(region_target)}", annotation_position="top left")
        fig.update_layout(title=f"Influência de cada planta no PPM {REGION}",
                          xaxis_title="Variação do refugo projetado da planta (%)",
                          yaxis_title=f"PPM {REGION}", hovermode="x unified", height=430,
                          legend=dict(orientation="h", y=-0.2))
        st.plotly_chart(fig, width="stretch")

        # Gráfico 2 — mapa de combinações de 2 plantas
        st.subheader("Mapa de combinações entre duas plantas")
        g1, g2, g3 = st.columns(3)
        pa = g1.selectbox("Planta (eixo X)", plants, index=plants.index("Itajubá") if "Itajubá" in plants else 0,
                          key="heat_a")
        pb_opts = [p for p in plants if p != pa]
        pb = g2.selectbox("Planta (eixo Y)", pb_opts, key="heat_b")
        hr = g3.slider("Faixa (%)", -100, 100, (-50, 50), step=5, key="heat_range")
        grid = np.arange(hr[0], hr[1] + 1, 5) / 100
        z = np.empty((len(grid), len(grid)))
        for i, vb in enumerate(grid):
            for j, va in enumerate(grid):
                vv = variation.copy()
                vv[pa], vv[pb] = va, vb
                z[i, j] = region_ppm_for(vv)
        fig = go.Figure()
        fig.add_heatmap(x=grid * 100, y=grid * 100, z=z, zmid=region_target,
                        colorscale=ESCALA_PPM,
                        colorbar=dict(title="PPM"),
                        hovertemplate=f"{pa}: %{{x:+.0f}}%<br>{pb}: %{{y:+.0f}}%<br>PPM {REGION}: %{{z:,.0f}}<extra></extra>")
        fig.add_contour(x=grid * 100, y=grid * 100, z=z, showscale=False, hoverinfo="skip",
                        contours=dict(start=region_target, end=region_target, size=1, coloring="lines",
                                      showlabels=True),
                        line=dict(width=3, color=COR_REGIAO), name="Target")
        fig.update_layout(title=f"PPM {REGION} — azul = dentro do target, vermelho = fora "
                                f"(linha = target {fmt_num(region_target)})",
                          xaxis_title=f"Variação refugo {pa} (%)", yaxis_title=f"Variação refugo {pb} (%)",
                          height=520)
        st.plotly_chart(fig, width="stretch")

        # Curva de fronteira: variação de pb necessária para cada variação de pa
        others = variation.drop([pa, pb])
        fixed = real_scrap.sum() + (proj_scrap0[others.index] * (1 + others)).sum()
        va_line = np.linspace(hr[0], hr[1], 81) / 100
        vb_line = (allowed_region - fixed - proj_scrap0[pa] * (1 + va_line)) / proj_scrap0[pb] - 1
        fig = go.Figure()
        fig.add_scatter(x=va_line * 100, y=np.minimum(vb_line * 100, hr[1]), mode="lines", fill="tozeroy",
                        fillcolor="rgba(42,120,214,0.10)", line=dict(width=0), hoverinfo="skip",
                        showlegend=False)
        fig.add_scatter(x=va_line * 100, y=vb_line * 100, mode="lines", name="Fronteira do target",
                        line=dict(color=COR_REGIAO, width=2),
                        hovertemplate=f"{pa}: %{{x:+.1f}}%<br>{pb} máx.: %{{y:+.1f}}%<extra></extra>")
        fig.update_yaxes(range=[min(hr[0], np.nanmin(vb_line * 100)), hr[1]])
        fig.update_layout(title=f"Curva de combinações {pa} × {pb} que mantém {REGION} no target "
                                "(área abaixo da linha = dentro do target)",
                          xaxis_title=f"Variação refugo {pa} (%)", yaxis_title=f"Variação máx. refugo {pb} (%)",
                          height=420)
        st.plotly_chart(fig, width="stretch")

        # Combinações de plantas com redução uniforme
        st.subheader("Redução uniforme necessária por combinação de plantas")
        st.caption("Para cada combinação, a variação % (igual em todas as plantas da combinação) que "
                   f"leva o {REGION} exatamente ao target. Valores < −100% = impossível só com essas plantas.")
        rows = []
        for k in range(1, len(plants) + 1):
            for combo in combinations(plants, k):
                combo = list(combo)
                rest = variation.drop(combo)
                fixed_c = real_scrap.sum() + (proj_scrap0[rest.index] * (1 + rest)).sum()
                denom = proj_scrap0[combo].sum()
                v_req = (allowed_region - fixed_c) / denom - 1 if denom else np.nan
                rows.append({"Combinação": " + ".join(combo), "Nº plantas": k,
                             "Variação necessária (%)": v_req * 100,
                             "Peças a reduzir (meses proj.)": -denom * v_req,
                             "Viável": "✅" if v_req >= -1 else "❌"})
        combo_df = pd.DataFrame(rows).sort_values("Variação necessária (%)", ascending=False)
        show_table(combo_df.set_index("Combinação").style.format(
            {"Variação necessária (%)": lambda x: fmt_num(x, 1), "Peças a reduzir (meses proj.)": fmt_num}),
            title="Combinações", width="stretch", height=420)

        sel = st.multiselect("Combinação para simular", plants, default=plants, key="combo_sel")
        if sel:
            vs = np.linspace(-0.6, 0.4, 51)
            ys = []
            for v in vs:
                vv = variation.copy()
                vv[sel] = v
                ys.append(region_ppm_for(vv))
            fig = go.Figure()
            fig.add_scatter(x=vs * 100, y=ys, mode="lines", name=" + ".join(sel),
                            line=dict(color=COR_REGIAO, width=2),
                            hovertemplate="Variação: %{x:+.0f}%<br>PPM: %{y:,.0f}<extra></extra>")
            fig.add_hline(y=region_target, line_dash="dash", line_color=COR_TARGET,
                          annotation_text=f"Target {fmt_num(region_target)}", annotation_position="top left")
            fig.update_layout(title=f"PPM {REGION} × variação uniforme em: {' + '.join(sel)}",
                              xaxis_title="Variação do refugo projetado (%)", yaxis_title=f"PPM {REGION}",
                              height=400, showlegend=False)
            st.plotly_chart(fig, width="stretch")

# =========================================================================== TELA 4
with tab4:
    if not HAS_PROJ:
        st.info(f"Não há meses projetados em {proj_year} — análise disponível apenas para anos em andamento.")
    else:
        st.caption("Eixos = PPM anual de cada planta (refugo do ano real + projetado ÷ produção do ano × 10⁶). A produção fica fixa na projeção atual; "
                   "as plantas não selecionadas ficam no cenário atual (tabela de variação). "
                   "Fronteira = combinações em que o PPM da BU1 SA fica exatamente no target.")
        s1, s2 = st.columns([2, 1])
        sim_pl = s1.multiselect("Plantas a combinar (2 = gráfico 2D, 3 = gráfico 3D)", plants,
                                default=plants[:2], max_selections=3, key="sim_plants")
        rr = s2.slider("Faixa dos eixos (variação do refugo projetado, %)", -100, 300, (-100, 50), step=5,
                       key="sim_range")

        if len(sim_pl) < 2:
            st.info("Selecione 2 ou 3 plantas.")
        else:
            ss = st.session_state
            others = [p for p in plants if p not in sim_pl]
            # Edições manuais da tabela "Valores no ponto analisado": planta -> refugo anual fixado
            manual = ss.setdefault("sim_manual", {})
            cur_annual = real_scrap + proj_scrap0 * (1 + variation)
            for p, ed in manual.items():
                if p not in sim_pl:
                    cur_annual[p] = ed["annual"]
            total_prod = prod.sum().sum()
            allowed = region_target * total_prod / 1e6
            others_sum = cur_annual[others].sum()
            free = allowed - others_sum  # soma máxima das plantas selecionadas
            lo = {p: real_scrap[p] + proj_scrap0[p] * (1 + rr[0] / 100) for p in sim_pl}
            hi = {p: real_scrap[p] + proj_scrap0[p] * (1 + rr[1] / 100) for p in sim_pl}
            for p in sim_pl:  # amplia o eixo se o valor editado estiver fora da faixa
                if p in manual:
                    lo[p] = min(lo[p], manual[p]["annual"])
                    hi[p] = max(hi[p], manual[p]["annual"])
            own = {p: data["targets"].get(p, np.nan) * prod_year[p] / 1e6 for p in sim_pl}
            step = {p: max(1, int((hi[p] - lo[p]) / 400)) for p in sim_pl}
            tag = "_".join(sim_pl)
            keys = {p: f"sim_{tag}_{p}" for p in sim_pl}
            for p in sim_pl:
                if p in manual and keys[p] not in ss:
                    ss[keys[p]] = int(round(manual[p]["annual"]))
            for p in sim_pl:
                ss[keys[p]] = int(np.clip(ss.get(keys[p], cur_annual[p]), lo[p], hi[p]))

            if len(sim_pl) == 3:
                fixed_p = st.radio("Plano de corte (planta com valor fixo)", sim_pl, horizontal=True,
                                   key=f"sim_plane_{tag}")
                pa, pb = [p for p in sim_pl if p != fixed_p]
            else:
                fixed_p = None
                pa, pb = sim_pl
            chart_key = f"sim_chart_{tag}_{fixed_p}"

            # Clique no gráfico 2D -> move o ponto analisado
            try:
                pts = ss[chart_key]["selection"]["points"]
            except (KeyError, TypeError):
                pts = []
            pts = [q for q in pts if "x" in q and "y" in q]
            if pts:
                sig = (round(pts[0]["x"], 1), round(pts[0]["y"], 1))  # eixos em PPM
                if ss.get(f"{chart_key}_last") != sig:
                    ss[f"{chart_key}_last"] = sig
                    ss[keys[pa]] = int(np.clip(sig[0] * prod_year[pa] / 1e6, lo[pa], hi[pa]))
                    ss[keys[pb]] = int(np.clip(sig[1] * prod_year[pb] / 1e6, lo[pb], hi[pb]))

            def reset_point():
                for k in list(ss.keys()):
                    if str(k).startswith(("sim_chart_", f"sim_{tag}_")):
                        del ss[k]
                ss["sim_manual"] = {}
                ss["pt_editor_ver"] = ss.get("pt_editor_ver", 0) + 1

            h1, h2 = st.columns([3, 1])
            h1.markdown("**Ponto analisado** — ajuste nos sliders ou clique no gráfico 2D")
            h2.button("↺ Voltar à projeção atual", on_click=reset_point, key="sim_reset", width="stretch",
                      help="Remove as simulações (sliders, cliques e edições manuais da tabela) e volta ao "
                           "cenário projetado a partir dos dados do arquivo Excel.")
            slider_order = ([fixed_p] if fixed_p else []) + [pa, pb]
            for col, p in zip(st.columns(len(slider_order)), slider_order):
                col.slider(f"Refugo anual — {p}" + (" (plano)" if p == fixed_p else ""),
                           int(lo[p]), int(hi[p]), step=step[p], key=keys[p], format="%d")
                col.caption(f"= **{fmt_num(ss[keys[p]] / prod_year[p] * 1e6)} PPM** · real {fmt_num(real_scrap[p])} pç · "
                            f"projeção atual {fmt_num(cur_annual[p])} pç · meta {fmt_num(own[p])} pç")
            point = {p: float(ss[keys[p]]) for p in sim_pl}
            for p in sim_pl:  # slider/clique moveu a planta -> deixa de ser edição manual
                if p in manual and int(point[p]) != int(round(manual[p]["annual"])):
                    del manual[p]
            ppm_pt = (others_sum + sum(point.values())) / total_prod * 1e6
            slack = allowed - others_sum - sum(point.values())

            m1, m2, m3 = st.columns(3)
            m1.metric(f"PPM {REGION} no ponto", fmt_num(ppm_pt),
                      delta=f"{fmt_num(ppm_pt - region_target)} vs target", delta_color="inverse",
                      help="(Σ refugo anual das plantas selecionadas no ponto + Σ refugo anual das demais) ÷ "
                           "Produção total BU1 SA × 10⁶.")
            m2.metric("Folga até o target (peças)", fmt_num(slack),
                      help="Target BU1 SA × Produção BU1 SA ÷ 10⁶ − Σ refugo anual de todas as plantas. "
                           "Positivo = ainda cabe esse refugo; negativo = precisa reduzir isso.")
            m3.metric("Situação vs target", "✅ Dentro" if slack >= 0 else "❌ Fora")

            def status_marker(ok, size=12):
                return dict(size=size, color=CORES[2] if ok else CORES[7], line=dict(color="white", width=2),
                            symbol="diamond")

            # Eixos dos gráficos em PPM da planta (refugo anual ÷ produção anual × 10⁶);
            # o hover mostra as peças e a diferença em relação à projeção atual.
            def ppm_of(p, q):
                return q / prod_year[p] * 1e6

            N_CD = 4  # valores de customdata por planta

            def hover_lines(p, idx):
                axis = "xyz"[idx]
                k = N_CD * idx
                return (f"<b>{p}</b>: %{{{axis}:,.0f}} PPM<br>"
                        f"&nbsp;&nbsp;refugo: %{{customdata[{k}]:,.0f}} pç<br>"
                        f"&nbsp;&nbsp;Δ vs projeção atual: %{{customdata[{k + 1}]:+,.0f}} pç<br>"
                        f"&nbsp;&nbsp;variação necessária: <b>%{{customdata[{k + 2}]:+,.1f}}%</b> no ano · "
                        f"<b>%{{customdata[{k + 3}]:+,.1f}}%</b> nos meses projetados<br>")

            def custom(pairs):
                """pairs: [(planta, array de peças)] -> customdata no último eixo:
                [pç, Δpç vs projeção atual, Δ% no ano, Δ% no refugo dos meses projetados, ...]."""
                cols = []
                for p, q in pairs:
                    delta = q - cur_annual[p]
                    proj_cur = cur_annual[p] - real_scrap[p]
                    cols += [q, delta, delta / cur_annual[p] * 100,
                             delta / proj_cur * 100 if proj_cur else np.full_like(np.asarray(q, float), np.nan)]
                return np.stack(cols, axis=-1)

            # ---------------- 3D
            if len(sim_pl) == 3:
                px_, py_, pz_ = sim_pl
                gx = np.linspace(lo[px_], hi[px_], 40)
                gy = np.linspace(lo[py_], hi[py_], 40)
                X, Y = np.meshgrid(gx, gy)
                Z = free - X - Y
                Z[(Z < lo[pz_]) | (Z > hi[pz_])] = np.nan
                hov3 = hover_lines(px_, 0) + hover_lines(py_, 1) + hover_lines(pz_, 2)
                fig3 = go.Figure()
                fig3.add_surface(
                    x=ppm_of(px_, X), y=ppm_of(py_, Y), z=ppm_of(pz_, Z), name="Fronteira do target",
                    showscale=False, opacity=0.85, colorscale=[[0, "#cfe0f5"], [1, "#2a78d6"]],
                    customdata=custom([(px_, X), (py_, Y), (pz_, Z)]),
                    hovertemplate=hov3 + "<extra>Fronteira do target</extra>")
                # plano de corte
                u, v = np.meshgrid([0.0, 1.0], [0.0, 1.0])
                c = {fixed_p: np.full_like(u, point[fixed_p]),
                     pa: lo[pa] + u * (hi[pa] - lo[pa]),
                     pb: lo[pb] + v * (hi[pb] - lo[pb])}
                fig3.add_surface(x=ppm_of(px_, c[px_]), y=ppm_of(py_, c[py_]), z=ppm_of(pz_, c[pz_]),
                                 showscale=False, opacity=0.3, colorscale=[[0, "#8a8984"], [1, "#8a8984"]],
                                 hoverinfo="skip", name=f"Plano {fixed_p}")
                # interseção plano × fronteira
                t = np.linspace(lo[pa], hi[pa], 100)
                line = {fixed_p: np.full_like(t, point[fixed_p]), pa: t, pb: free - point[fixed_p] - t}
                ok_ = (line[pb] >= lo[pb]) & (line[pb] <= hi[pb])
                ln = {p: line[p][ok_] for p in sim_pl}
                fig3.add_scatter3d(x=ppm_of(px_, ln[px_]), y=ppm_of(py_, ln[py_]), z=ppm_of(pz_, ln[pz_]),
                                   mode="lines", line=dict(color=COR_REGIAO, width=6), name="Interseção com o plano",
                                   customdata=custom([(px_, ln[px_]), (py_, ln[py_]), (pz_, ln[pz_])]),
                                   hovertemplate=hov3 + "<extra>Interseção</extra>")
                cur3 = {p: np.array([cur_annual[p]]) for p in sim_pl}
                fig3.add_scatter3d(x=ppm_of(px_, cur3[px_]), y=ppm_of(py_, cur3[py_]), z=ppm_of(pz_, cur3[pz_]),
                                   mode="markers", marker=dict(size=6, color=COR_TARGET), name="Projeção atual",
                                   customdata=custom([(px_, cur3[px_]), (py_, cur3[py_]), (pz_, cur3[pz_])]),
                                   hovertemplate=hov3 + "<extra>Projeção atual</extra>")
                pt3 = {p: np.array([point[p]]) for p in sim_pl}
                fig3.add_scatter3d(x=ppm_of(px_, pt3[px_]), y=ppm_of(py_, pt3[py_]), z=ppm_of(pz_, pt3[pz_]),
                                   mode="markers+text", marker=status_marker(slack >= 0, 8),
                                   text=[f"PPM {REGION} {fmt_num(ppm_pt)}"], textposition="top center",
                                   name="Ponto analisado",
                                   customdata=custom([(px_, pt3[px_]), (py_, pt3[py_]), (pz_, pt3[pz_])]),
                                   hovertemplate=hov3 + f"PPM {REGION}: {fmt_num(ppm_pt)}<extra>Ponto</extra>")
                fig3.update_layout(
                    title=f"Fronteira do target {REGION} — abaixo da superfície = dentro do target",
                    scene=dict(xaxis=dict(title=f"{px_} (PPM)", tickformat=",.0f"),
                               yaxis=dict(title=f"{py_} (PPM)", tickformat=",.0f"),
                               zaxis=dict(title=f"{pz_} (PPM)", tickformat=",.0f")),
                    separators=",.",
                    height=650, legend=dict(orientation="h", y=-0.05), margin=dict(l=0, r=0, t=50, b=0))
                st.plotly_chart(fig3, width="stretch", key=f"sim3d_{tag}")

            # ---------------- 2D (gráfico próprio ou corte do 3D)
            fixed_val = point[fixed_p] if fixed_p else 0.0
            base_sum = others_sum + fixed_val
            free_ab = allowed - base_sum
            ga = np.linspace(lo[pa], hi[pa], 80)
            gb = np.linspace(lo[pb], hi[pb], 80)
            A, B = np.meshgrid(ga, gb)
            Zp = (base_sum + A + B) / total_prod * 1e6
            hov2 = hover_lines(pa, 0) + hover_lines(pb, 1)
            fig = go.Figure()
            fig.add_contour(x=ppm_of(pa, ga), y=ppm_of(pb, gb), z=Zp, zmid=region_target, hoverinfo="skip",
                            ncontours=14, colorscale=ESCALA_PPM,
                            contours=dict(showlabels=True, labelformat=",.0f",
                                          labelfont=dict(size=10, color="#3a3a38")),
                            colorbar=dict(title=f"PPM {REGION}", tickformat=",.0f"),
                            line=dict(width=0.5, color="rgba(0,0,0,0.25)"))
            fa = np.linspace(lo[pa], hi[pa], 200)
            fb = free_ab - fa
            ok_ = (fb >= lo[pb]) & (fb <= hi[pb])
            fig.add_scatter(x=ppm_of(pa, fa[ok_]), y=ppm_of(pb, fb[ok_]), mode="lines", name="Fronteira do target",
                            line=dict(color=COR_REGIAO, width=3), customdata=custom([(pa, fa[ok_]), (pb, fb[ok_])]),
                            hovertemplate=hov2 + f"PPM {REGION}: {fmt_num(region_target)} (target)<extra>Fronteira</extra>")
            for p, fn in [(pa, fig.add_vline), (pb, fig.add_hline)]:
                tg = data["targets"].get(p, np.nan)
                if ppm_of(p, lo[p]) <= tg <= ppm_of(p, hi[p]):
                    fn(tg, line_dash="dot", line_color=COR_TARGET,
                       annotation_text=f"meta {p} ({fmt_num(tg)} PPM)",
                       annotation_position="top left" if p == pb else "top right")
            # grade invisível para hover/clique
            cg_a, cg_b = np.meshgrid(np.linspace(lo[pa], hi[pa], 35), np.linspace(lo[pb], hi[pb], 35))
            cg_a, cg_b = cg_a.ravel(), cg_b.ravel()
            cg_ppm = (base_sum + cg_a + cg_b) / total_prod * 1e6
            cg_cd = np.column_stack([custom([(pa, cg_a), (pb, cg_b)]), cg_ppm,
                                     cg_a + cg_b - cur_annual[pa] - cur_annual[pb]])
            fig.add_scatter(x=ppm_of(pa, cg_a), y=ppm_of(pb, cg_b), mode="markers", showlegend=False,
                            marker=dict(size=14, color="rgba(0,0,0,0)"), customdata=cg_cd,
                            hovertemplate=hov2 + f"<b>PPM {REGION}: %{{customdata[{2 * N_CD}]:,.0f}}</b><br>"
                                                 f"Δ total vs projeção atual: %{{customdata[{2 * N_CD + 1}]:+,.0f}} pç"
                                                 "<extra>clique para analisar</extra>")
            cur_ab = {p: np.array([cur_annual[p]]) for p in (pa, pb)}
            fig.add_scatter(x=ppm_of(pa, cur_ab[pa]), y=ppm_of(pb, cur_ab[pb]), mode="markers", name="Projeção atual",
                            marker=dict(size=10, color=COR_TARGET, line=dict(color="white", width=2)),
                            customdata=custom([(pa, cur_ab[pa]), (pb, cur_ab[pb])]),
                            hovertemplate=hov2 + "<extra>Projeção atual</extra>")
            pt_ab = {p: np.array([point[p]]) for p in (pa, pb)}
            fig.add_scatter(x=ppm_of(pa, pt_ab[pa]), y=ppm_of(pb, pt_ab[pb]), mode="markers+text",
                            name="Ponto analisado", marker=status_marker(slack >= 0),
                            text=[f"PPM {REGION} {fmt_num(ppm_pt)}"], textposition="top center",
                            customdata=custom([(pa, pt_ab[pa]), (pb, pt_ab[pb])]),
                            hovertemplate=hov2 + f"PPM {REGION}: {fmt_num(ppm_pt)}<extra>Ponto analisado</extra>")
            fig.update_layout(
                title=(f"Corte no plano {fixed_p} = {fmt_num(ppm_of(fixed_p, fixed_val))} PPM — " if fixed_p else "")
                      + f"{pa} × {pb}: abaixo da linha preta = dentro do target",
                xaxis_title=f"Refugo {pa} (PPM ano)", yaxis_title=f"Refugo {pb} (PPM ano)", height=560,
                legend=dict(orientation="h", y=-0.15), separators=",.")
            fig.update_xaxes(range=[ppm_of(pa, lo[pa]), ppm_of(pa, hi[pa])], tickformat=",.0f")
            fig.update_yaxes(range=[ppm_of(pb, lo[pb]), ppm_of(pb, hi[pb])], tickformat=",.0f")
            st.plotly_chart(fig, width="stretch", key=chart_key, on_select="rerun", selection_mode="points")

            # ---------------- Tabela do ponto (editável: PPM ou Refugo ano)
            n_proj = max(len(data["projected"]), 1)
            rows = []
            for p in plants:
                annual = point.get(p, cur_annual[p])
                tg = data["targets"].get(p, np.nan)
                ppm_p = annual / prod_year[p] * 1e6
                ed = manual.get(p)
                rows.append({"Planta": p + (" ★" if p in sim_pl else ""),
                             "Edição": f"✏️ {ed['col']} manual" if ed else "",
                             "PPM": ppm_p, "Refugo ano": annual,
                             "Refugo meses proj.": annual - real_scrap[p],
                             "Refugo médio/mês proj.": (annual - real_scrap[p]) / n_proj,
                             "Variação vs base (%)": ((annual - real_scrap[p]) / proj_scrap0[p] - 1) * 100,
                             "Produção ano": prod_year[p], "Target PPM": tg,
                             "Desvio %": (ppm_p - tg) / tg * 100, "Status": "✅ Dentro" if ppm_p <= tg else "❌ Fora",
                             "Δ PPM vs ano ant. (%)": pct(ppm_p, prev_ppm[p])})
            pt_df = pd.DataFrame(rows).set_index("Planta")
            tot = pt_df[["Refugo ano", "Refugo meses proj.", "Refugo médio/mês proj.", "Produção ano"]].sum()
            pt_df.loc[REGION] = {**tot.to_dict(), "Edição": "", "Variação vs base (%)": np.nan, "PPM": ppm_pt,
                                 "Target PPM": region_target, "Desvio %": (ppm_pt - region_target) / region_target * 100,
                                 "Status": "✅ Dentro" if slack >= 0 else "❌ Fora",
                                 "Δ PPM vs ano ant. (%)": pct(ppm_pt, prev_ppm[REGION])}
            row_plants = plants + [REGION]
            editable = ["PPM", "Refugo ano"]

            def on_point_edit(editor_key, keys_map):
                changes = ss[editor_key].get("edited_rows", {})
                man = ss.setdefault("sim_manual", {})
                for r, ch in changes.items():
                    p = row_plants[int(r)]
                    if p == REGION:
                        continue
                    for col, val in ch.items():
                        if val is None or col not in editable:
                            continue
                        annual = float(val) * prod_year[p] / 1e6 if col == "PPM" else float(val)
                        annual = max(annual, 0.0)
                        man[p] = {"col": col, "annual": annual}
                        if p in keys_map:
                            ss[keys_map[p]] = int(round(annual))
                ss["pt_editor_ver"] = ss.get("pt_editor_ver", 0) + 1

            st.subheader("Valores no ponto analisado (★ = plantas do gráfico)")
            st.caption("✏️ Edite **PPM** ou **Refugo ano** de qualquer planta (ex.: Itajubá = 39.000 PPM) para simular "
                       "o cenário. Linhas editadas ficam destacadas em amarelo; plantas ★ movem o ponto do gráfico, "
                       "as demais mudam a fronteira. A linha BU1 SA é calculada (edições nela são ignoradas).")
            edit_key = f"pt_editor_{ss.get('pt_editor_ver', 0)}"
            hl = "background-color: rgba(255, 213, 79, 0.45); font-weight: 600"
            styler = (pt_df.style
                      .format({c: (lambda x: fmt_num(x, 1) + "%") if "%" in c else fmt_num
                               for c in pt_df.columns if c not in ("Status", "Edição", *editable)})
                      .apply(lambda r: [hl if r["Edição"] else "" for _ in r], axis=1))
            cfg = {c: st.column_config.Column(help=HELP.get(c)) for c in pt_df.columns if c in HELP}
            cfg["PPM"] = st.column_config.NumberColumn(
                "PPM ✏️", format="localized", min_value=0.0, step=1.0,
                help="EDITÁVEL. Refugo ano ÷ Produção ano × 10⁶. Ao editar: Refugo ano = PPM × Produção ano ÷ 10⁶.")
            cfg["Refugo ano"] = st.column_config.NumberColumn(
                "Refugo ano ✏️", format="localized", min_value=0.0, step=1.0,
                help="EDITÁVEL. Peças refugadas no ano (real + projetado). Ao editar, o PPM é recalculado.")
            cfg["Edição"] = st.column_config.TextColumn("Edição", help="Mostra se o valor da planta foi editado manualmente.")
            st.data_editor(styler, column_config=cfg, disabled=[c for c in pt_df.columns if c not in editable],
                           width="stretch", key=edit_key, on_change=on_point_edit, args=(edit_key, keys))
            e1, e2 = st.columns([1, 3])
            if e1.button("🧹 Limpar edições manuais", disabled=not manual, key="clear_manual"):
                ss["sim_manual"] = {}
                for p in sim_pl:
                    ss.pop(keys[p], None)
                ss["pt_editor_ver"] = ss.get("pt_editor_ver", 0) + 1
                st.rerun()
            if manual:
                e2.info("Edições ativas: " + "; ".join(
                    f"{p}: {fmt_num(v['annual'])} pç/ano ({fmt_num(v['annual'] / prod_year[p] * 1e6)} PPM)"
                    for p, v in manual.items()))
            with st.popover("ℹ️ Memorial de cálculo — Ponto analisado"):
                for c in pt_df.columns:
                    if cfg.get(c, {}).get("help"):
                        st.markdown(f"**{c}** — {cfg[c]['help']}")

# =========================================================================== TELA 5
with tab5:
    st.header("Memorial de cálculo")
    st.markdown(f"**Fonte dos dados:** arquivo *Dados brutos* — conta **201** = peças refugadas, conta **202** = "
                f"peças produzidas, bloco **Targets** = metas em PPM. Meses realizados: **{real_txt}**. "
                f"Meses projetados: **{proj_txt}**.")
    st.subheader("1. PPM")
    st.latex(r"PPM = \frac{\text{Peças refugadas}}{\text{Peças produzidas}} \times 10^6")
    st.markdown("Para a região BU1 SA somam-se refugo e produção de todas as plantas antes de dividir "
                "(média ponderada pela produção).")
    st.subheader("2. Produção base (por planta)")
    st.markdown(f"Metodologia atual: **{method}** — meses usados: **{prod_cols_txt}**.")
    st.latex(r"Prod_{base} = \frac{1}{n}\sum_{m \in \text{meses usados}} Prod_m")
    st.markdown(f"- *Média do ano*: todos os meses realizados de {proj_year}.\n"
                "- *Média móvel (últimos X meses)*: os X últimos meses realizados (pode incluir anos anteriores).\n"
                "- *Meses selecionados*: apenas os meses escolhidos (ano atual e anterior).\n"
                f"- *Média do ano anterior*: média mensal de {prev_year}.\n"
                f"- *Mesmo mês do ano anterior*: cada mês projetado usa a produção do mesmo mês de {prev_year}.")
    st.subheader("3. Produção projetada (mês m projetado)")
    st.latex(r"Prod_{m} = Prod_{base} \times \frac{Peso_{Geral,m}}{100} \times \frac{Peso_{Planta,m}}{100}")
    st.subheader("4. PPM base e refugo projetado")
    st.markdown(f"PPM base calculado em: **{ppm_cols_txt}** ({ppm_method}).")
    st.latex(r"PPM_{base} = \frac{\sum Refugo}{\sum Prod} \times 10^6 \qquad "
             r"Refugo_m = PPM_{base} \times \frac{Prod_m}{10^6} \times (1 + Var\%)")
    st.subheader("5. Fechamento do ano")
    st.latex(r"PPM_{ano} = \frac{Refugo_{real} + \sum Refugo_{proj}}{Prod_{real} + \sum Prod_{proj}} \times 10^6"
             r"\qquad Gap = PPM_{ano} - Target")
    st.subheader("6. Refugo permitido para atingir o target")
    st.latex(r"Refugo_{permitido} = \frac{Target \times Prod_{ano}}{10^6} - Refugo_{real}"
             r"\qquad Var_{nec} = \frac{Refugo_{permitido}}{Refugo_{proj,base}} - 1")
    st.subheader("7. Combinações e simulador")
    st.latex(r"\sum_{i} Refugo_{ano,i} \le \frac{Target_{BU1\,SA} \times \sum_i Prod_{ano,i}}{10^6}")
    st.markdown("Como a produção é fixa, a fronteira do target é **linear** nas quantidades refugadas: "
                "uma reta (2 plantas) ou um plano (3 plantas). Variação uniforme v de uma combinação C:")
    st.latex(r"v = \frac{Refugo_{permitido,BU1} - \sum Refugo_{real} - \sum_{i \notin C} Refugo_{proj,i}(1+v_i)}"
             r"{\sum_{i \in C} Refugo_{proj,base,i}} - 1")

    st.subheader("8. Exemplo numérico com o cenário atual")
    ex = st.selectbox("Planta", plants, key="memo_plant")
    n_proj = len(data["projected"])
    r = summary.loc[ex]
    pb_ = prod_base.loc[ex] if HAS_PROJ else pd.Series(dtype=float)
    lines = [
        f"- Produção base ({method}) a partir de {prod_cols_txt}"
        + (f" = **{fmt_num(pb_.iloc[0])}** pç/mês" if HAS_PROJ and method != M_MESMO_MES else " (varia por mês)"),
        f"- PPM base ({ppm_cols_txt}) = {fmt_num(period_sum(hist_scrap, ppm_cols, plants)[ex])} ÷ "
        f"{fmt_num(period_sum(hist_prod, ppm_cols, plants)[ex])} × 10⁶ = **{fmt_num(ppm_b[ex], 1)}**",
        f"- Variação do refugo = **{fmt_num(variation[ex] * 100, 1)}%**",
    ]
    for m in data["projected"]:
        lines.append(f"- {MESES[m - 1]}: Produção = {fmt_num(pb_[m])} × {fmt_num(weights.loc[m, ex] * 100)}% = "
                     f"{fmt_num(prod.loc[ex, m])}; Refugo = {fmt_num(ppm_b[ex], 1)} × {fmt_num(prod.loc[ex, m])} "
                     f"÷ 10⁶ × (1 {variation[ex] * 100:+.1f}%) = {fmt_num(scrap.loc[ex, m])}")
    lines += [
        f"- Refugo ano = {fmt_num(r['Refugo real'])} + {fmt_num(r['Refugo projetado'])} = **{fmt_num(r['Refugo ano'])}**",
        f"- Produção ano = {fmt_num(r['Produção real'])} + {fmt_num(r['Produção projetada'])} = "
        f"**{fmt_num(r['Produção ano'])}**",
        f"- PPM ano = {fmt_num(r['Refugo ano'])} ÷ {fmt_num(r['Produção ano'])} × 10⁶ = "
        f"**{fmt_num(r['PPM ano (proj.)'])}** (target {fmt_num(r['Target PPM'])}, gap {fmt_pct(r['Gap vs target (%)'])}; "
        f"peças p/ atingir target {fmt_num(r['Peças p/ atingir target'])})",
    ]
    if n_proj:
        allowed_ex = r["Target PPM"] * r["Produção ano"] / 1e6 - r["Refugo real"]
        lines.append(f"- Refugo permitido nos meses projetados = {fmt_num(r['Target PPM'])} × "
                     f"{fmt_num(r['Produção ano'])} ÷ 10⁶ − {fmt_num(r['Refugo real'])} = **{fmt_num(allowed_ex)}** "
                     f"({fmt_num(allowed_ex / n_proj)} pç/mês)")
    st.markdown("\n".join(lines))

    st.subheader("9. Histórico e comparação com o ano anterior")
    st.markdown(f"Ano da projeção: **{proj_year}**. Anos com dados: **{', '.join(map(str, data_years))}** "
                "(anos sem dados são ignorados). Comparações com o ano anterior usam o mesmo intervalo de meses "
                "quando indicado (mês / acumulado) ou o ano anterior inteiro (fechamento projetado).")
    st.latex(r"\Delta\% = \left(\frac{Valor_{ano}}{Valor_{ano-1}} - 1\right) \times 100")

    st.subheader("10. Glossário das colunas")
    gl = pd.DataFrame({"Coluna": list(HELP), "Cálculo": list(HELP.values())})
    st.dataframe(gl[~gl["Coluna"].isin(MESES)], hide_index=True, width="stretch")

# =========================================================================== TELA 6
with tab6:
    st.header("Dados de entrada (template)")
    st.caption(f"Template padrão de {TEMPLATE_START} a {max(EDIT_YEARS)}. Anos sem dados são ignorados nas análises. "
               "Edite qualquer célula para simular ou corrigir dados — todas as telas são recalculadas. Para lançar um "
               "mês novo, preencha produção e refugo de todas as plantas; apague as células de um mês do ano da "
               "projeção para que ele volte a ser projetado.")

    def restore_raw():
        st.session_state[STORE_KEY] = {}
        st.session_state["raw_store_ver"] = st.session_state.get("raw_store_ver", 0) + 1

    tpl_years = list(range(TEMPLATE_START, max(EDIT_YEARS) + 1))
    b1, b2, b3 = st.columns(3)
    b1.button("↺ Restaurar dados originais do Excel", on_click=restore_raw, disabled=not n_raw_changed,
              width="stretch", key="raw_restore")
    b2.download_button("⬇️ Baixar dados editados (layout Dados brutos)",
                       to_raw_layout(plants, raw["codes"], {y: raw_edit[y]["prod"] for y in EDIT_YEARS},
                                     {y: raw_edit[y]["scrap"] for y in EDIT_YEARS}, raw_tg_tbl, EDIT_YEARS),
                       "Dados brutos editado.xlsx", width="stretch",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       help="Mesmo layout do arquivo original, com todos os anos; pode ser reenviado pelo upload.")
    b3.download_button(f"⬇️ Baixar template em branco ({TEMPLATE_START}–{max(tpl_years)})",
                       to_raw_layout(plants, raw["codes"], {}, {}, raw_tg_tbl, tpl_years, blank=True),
                       "Template Dados brutos.xlsx", width="stretch",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       help="Layout vazio com plantas e todos os meses desde 2020, para preencher fora do app e "
                            "enviar pelo upload. Anos deixados em branco são ignorados.")
    if n_raw_changed:
        st.warning(f"✏️ {n_raw_changed} valor(es) diferente(s) do Excel — veja a lista de alterações abaixo.")
    else:
        st.success("Dados iguais ao arquivo Excel carregado.")

    edit_year = st.selectbox("Ano para visualizar/editar", EDIT_YEARS, index=EDIT_YEARS.index(proj_year),
                             key="raw_edit_year",
                             format_func=lambda y: f"{y}" + ("" if y in data_years else " (sem dados — ignorado)"))
    num_cfg = {m: st.column_config.NumberColumn(m, format="localized", min_value=0.0, step=1.0,
                                                help=f"{m}/{edit_year} — vazio = sem dado")
               for m in MESES}
    e1, e2 = st.columns(2)
    for col, kind, title in [(e1, "prod", "Peças produzidas (conta 202)"), (e2, "scrap", "Peças refugadas (conta 201)")]:
        with col:
            st.subheader(f"{title} — {edit_year}")
            k = raw_editor_key(kind, edit_year)
            st.data_editor(raw_edit[edit_year][kind], column_config=num_cfg, key=k, width="stretch",
                           on_change=on_raw_edit, args=(kind, edit_year, k))
    t1, t2 = st.columns([1, 2])
    with t1:
        st.subheader("Targets (PPM)")
        st.caption(f"Aplicados ao ano da projeção ({proj_year}).")
        k = raw_editor_key("tg", 0)
        st.data_editor(raw_tg_tbl, key=k, width="stretch", on_change=on_raw_edit, args=("tg", 0, k),
                       column_config={"Target PPM": st.column_config.NumberColumn(
                           format="localized", min_value=0.0, help="Meta de PPM da planta / região.")})
    with t2:
        st.subheader(f"PPM resultante — {edit_year}")
        pe, se = raw_edit[edit_year]["prod"].astype(float), raw_edit[edit_year]["scrap"].astype(float)
        ppm_tbl = se / pe * 1e6
        ppm_tbl.loc[REGION] = se.sum(min_count=len(plants)) / pe.sum(min_count=len(plants)) * 1e6
        show_table(ppm_tbl.style.format(fmt_num), title="PPM dos dados de entrada", width="stretch",
                   extra_help={m: f"{m}: Refugo ÷ Produção × 10⁶ com os valores (editados) desta aba." for m in MESES})

    if n_raw_changed:
        st.subheader("Valores alterados vs Excel (destacados)")
        hl = "background-color: rgba(255, 213, 79, 0.55); font-weight: 600"
        blocks = [(f"Produção {y}", raw_edit[y]["prod"], raw_default[y]["prod"], raw_changed[y]["prod"])
                  for y in EDIT_YEARS] + \
                 [(f"Refugo {y}", raw_edit[y]["scrap"], raw_default[y]["scrap"], raw_changed[y]["scrap"])
                  for y in EDIT_YEARS] + [("Targets", raw_tg_tbl, raw_tg_default, raw_changed["tg"])]
        for title, tbl, orig, mask in blocks:
            if not mask.values.any():
                continue
            st.markdown(f"**{title}** — valores no app; lista 'original → novo' abaixo")
            st.dataframe(tbl.style.format(fmt_num).apply(lambda _, mk=mask: np.where(mk, hl, ""), axis=None),
                         width="stretch")
            changes = [f"{r} · {c}: {fmt_num(orig.loc[r, c]) or 'vazio'} → {fmt_num(tbl.loc[r, c]) or 'vazio'}"
                       for r in tbl.index for c in tbl.columns if mask.loc[r, c]]
            st.caption(" | ".join(changes))
