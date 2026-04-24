import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px

st.set_page_config(page_title="Cost Database Explorer", layout="wide")
st.title("Cost Database Explorer")

# ── File upload ────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("Data Source")
    uploaded_file = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx", "xls"])
    st.caption("If no file is uploaded the bundled example database is used.")


@st.cache_data
def load_data(source):
    xl = pd.ExcelFile(source)
    df_gen  = pd.read_excel(xl, sheet_name="General")
    df_cost = pd.read_excel(xl, sheet_name="Cost")
    return df_gen, df_cost


try:
    source = uploaded_file if uploaded_file is not None else "Costdatabase v.3.1.xlsx"
    df_gen, df_cost = load_data(source)
except Exception as e:
    st.error(f"Could not load data: {e}")
    st.stop()

# ── Sidebar: FX rates for non-USD currencies ──────────────────────────────────
_curr_col      = next((c for c in df_gen.columns  if c.strip().lower() == "native currency"), None)
_fx_col        = next((c for c in df_gen.columns  if c.strip().lower() == "fx to usd"), None)
_cost_curr_col = next((c for c in df_cost.columns if c.strip().lower() == "currency"), None)
_cost_fx_col   = next((c for c in df_cost.columns if c.strip().lower() == "fx to usd"), None)

fx_rates = {"USD": 1.0}   # currency → rate applied to convert native → USD

_all_non_usd = set()
if _curr_col:
    _all_non_usd |= {str(c).strip().upper() for c in df_gen[_curr_col].dropna()
                     if str(c).strip().upper() != "USD"}
if _cost_curr_col:
    _all_non_usd |= {str(c).strip().upper() for c in df_cost[_cost_curr_col].dropna()
                     if str(c).strip().upper() != "USD"}

if _all_non_usd:
    with st.sidebar:
        st.markdown("---")
        st.subheader("Currency conversion to USD")
        for curr_str in sorted(_all_non_usd):
            default = 1.0
            if _fx_col and _curr_col:
                mask = df_gen[_curr_col].astype(str).str.strip().str.upper() == curr_str
                _sug = df_gen.loc[mask, _fx_col].apply(pd.to_numeric, errors="coerce").mean()
                if pd.notna(_sug):
                    default = float(_sug)
            if default == 1.0 and _cost_fx_col and _cost_curr_col:
                mask = df_cost[_cost_curr_col].astype(str).str.strip().str.upper() == curr_str
                _sug = df_cost.loc[mask, _cost_fx_col].apply(pd.to_numeric, errors="coerce").mean()
                if pd.notna(_sug):
                    default = float(_sug)
            fx_rates[curr_str] = st.number_input(
                f"{curr_str} → USD",
                min_value=0.0001,
                value=round(default, 6),
                format="%.6f",
                key=f"fx_{curr_str}",
            )

def get_fx(currency):
    """Return the USD conversion rate for a currency string."""
    if pd.isna(currency):
        return 1.0
    return fx_rates.get(str(currency).strip().upper(), 1.0)


# ── Cost-sheet aggregation ─────────────────────────────────────────────────────

_c_pid   = next((c for c in df_cost.columns if c.strip() == "Project ID"),    None)
_c_phase = next((c for c in df_cost.columns if c.strip() == "Cost Phase"),    None)
_c_grp   = next((c for c in df_cost.columns if c.strip() == "Cost Group"),    None) or \
           next((c for c in df_cost.columns if c.strip() == "Component"),     None)
_c_val   = next((c for c in df_cost.columns if c.strip() == "Native Value"),  None)
_c_unit  = next((c for c in df_cost.columns if c.strip() == "Unit Basis"),    None)
_c_dc    = next((c for c in df_cost.columns if c.strip() == "Project DC MWp"), None)
_c_inc   = next((c for c in df_cost.columns if c.strip() == "Include?"),      None)


def cost_totals(project_ids, phase_values, to_m_usd=True):
    """
    Aggregate Cost sheet rows for the given projects and Cost Phase values.
    to_m_usd=True  → _value column in m$ USD  (for CAPEX chart)
    to_m_usd=False → _value column in USD/year (for OPEX chart)
    Returns DataFrame[Project ID, Cost Group, _value].
    """
    if _c_pid is None or _c_phase is None or not project_ids:
        return pd.DataFrame(columns=["Project ID", "Cost Group", "_value"])

    df = df_cost[df_cost[_c_pid].isin(project_ids)].copy()
    df = df[df[_c_phase].isin(phase_values)]
    if _c_inc:
        df = df[df[_c_inc].astype(str).str.strip().str.upper() != "N"]

    df["_native"] = pd.to_numeric(df[_c_val], errors="coerce") if _c_val else np.nan
    df["_fx"]     = df[_cost_curr_col].apply(get_fx) if _cost_curr_col else 1.0
    df["_dc_kwp"] = (pd.to_numeric(df[_c_dc], errors="coerce") * 1000
                     if _c_dc else pd.Series(0.0, index=df.index))

    units = df[_c_unit].astype(str).str.strip() if _c_unit else pd.Series("", index=df.index)
    base  = df["_native"] * df["_fx"]

    df["_usd"] = np.where(
        units == "m$",    base * 1e6,
        np.where(
            units == "$/kWp", base * df["_dc_kwp"].where(df["_dc_kwp"] > 0),
            base              # $/year or $
        )
    )

    df["_value"] = df["_usd"] / 1e6 if to_m_usd else df["_usd"]

    grp_col = _c_grp or "Component"
    agg = (
        df[df["_value"].notna() & (df["_value"] > 0)]
        .groupby([_c_pid, grp_col], as_index=False)["_value"]
        .sum()
        .rename(columns={_c_pid: "Project ID", grp_col: "Cost Group"})
    )
    return agg


# ── Helpers ───────────────────────────────────────────────────────────────────
def numeric_cols(df):
    return df.select_dtypes(include="number").columns.tolist()

COLOUR_SEQ = px.colors.qualitative.Plotly

# ══════════════════════════════════════════════════════════════════════════════
# CHART 1 — Bubble chart
# ══════════════════════════════════════════════════════════════════════════════
st.header("Projects Overview")

gen_all = df_gen.columns.tolist()
gen_num = numeric_cols(df_gen)

c1, c2, c3 = st.columns(3)
with c1:
    x_col = st.selectbox("X-axis (grouping)", options=gen_all,
                         index=gen_all.index("Region") if "Region" in gen_all else 0)
with c2:
    y1_col = st.selectbox("Y-axis", options=gen_num, index=0)
with c3:
    size_options = ["— none —"] + [c for c in gen_num if c != y1_col]
    size_col = st.selectbox("Bubble size & colour", options=size_options, index=0)
    size_col = None if size_col == "— none —" else size_col

x_unique = sorted(df_gen[x_col].dropna().astype(str).unique().tolist())
tag_filter = st.multiselect(f"Filter by {x_col}", options=x_unique, default=x_unique)

df_f = (df_gen[df_gen[x_col].astype(str).isin(tag_filter)] if tag_filter else df_gen) \
       .reset_index(drop=True).copy()

# Completeness: share of non-NaN columns per row
df_f["_completeness"] = (df_f.notna().sum(axis=1) / len(df_f.columns) * 100).round(0).astype(int)

hover_col = next((c for c in ["Project Name", "Project ID"] if c in df_f.columns), None)

# X positions with jitter for categorical axes
_x_is_num = pd.api.types.is_numeric_dtype(df_gen[x_col])
if _x_is_num:
    df_f["_x_pos"] = pd.to_numeric(df_f[x_col], errors="coerce")
    _cat_map = {}
else:
    _cats = df_f[x_col].astype(str).unique().tolist()
    _cat_map = {c: i for i, c in enumerate(_cats)}
    rng = np.random.default_rng(42)
    df_f["_x_pos"] = df_f[x_col].astype(str).map(_cat_map) + rng.uniform(-0.3, 0.3, len(df_f))

# 3-bucket sizes relative to filtered range
_metric_col = size_col if size_col else y1_col
_metric = pd.to_numeric(df_f[_metric_col], errors="coerce").abs()
_lo, _hi = _metric.min(), _metric.max()
if pd.notna(_hi) and _hi > _lo:
    _pct = (_metric - _lo) / (_hi - _lo) * 100
    df_f["_size"] = _pct.apply(lambda p: 16 if p <= 30 else (28 if p <= 60 else 44))
    df_f["_size"] = df_f["_size"].where(_metric.notna(), 16)
else:
    df_f["_size"] = 28

_color_vals = _metric.fillna(_lo if pd.notna(_lo) else 0)

# Hover text
def _fmt(s):
    return pd.to_numeric(s, errors="coerce").round(2).astype(str)

hover_parts = []
if hover_col:
    hover_parts.append(df_f[hover_col].astype(str))
hover_parts.append("<b>" + x_col + ":</b> " + df_f[x_col].astype(str))
hover_parts.append("<b>" + y1_col + ":</b> " + _fmt(df_f[y1_col]))
if size_col:
    hover_parts.append("<b>" + size_col + ":</b> " + _fmt(df_f[size_col]))
hover_parts.append("<b>Completeness:</b> " + df_f["_completeness"].astype(str) + "%")
hover_text = ["<br>".join(row) for row in zip(*hover_parts)]

bubble_label = df_f.apply(
    lambda r: f"{r['_completeness']}%" if r["_size"] >= 28 else "", axis=1
).tolist()

fig1 = go.Figure(go.Scatter(
    x=df_f["_x_pos"].tolist(),
    y=pd.to_numeric(df_f[y1_col], errors="coerce").tolist(),
    mode="markers+text",
    text=bubble_label,
    textposition="middle center",
    textfont=dict(size=8, color="white", family="Arial Black"),
    hovertext=hover_text,
    hovertemplate="%{hovertext}<extra></extra>",
    marker=dict(
        size=df_f["_size"].tolist(),
        sizemode="diameter",
        color=_color_vals.tolist(),
        colorscale="RdYlBu_r",
        showscale=True,
        colorbar=dict(title=dict(text=_metric_col, side="right"), thickness=14, len=0.8),
        cmin=float(_lo) if pd.notna(_lo) else None,
        cmax=float(_hi) if pd.notna(_hi) else None,
        opacity=0.85,
        line=dict(width=0.6, color="white"),
    ),
))

if not _x_is_num and _cat_map:
    fig1.update_xaxes(tickvals=list(_cat_map.values()),
                      ticktext=list(_cat_map.keys()), tickangle=-30)

fig1.update_layout(height=540, xaxis_title=x_col, yaxis_title=y1_col,
                   hovermode="closest", margin=dict(r=120, b=80))
st.plotly_chart(fig1, use_container_width=True)

with st.expander("Show filtered data"):
    st.dataframe(
        df_f.drop(columns=["_x_pos", "_size", "_completeness"], errors="ignore")
            .reset_index(drop=True),
        use_container_width=True,
    )

# Shared scope across Chart 2 and 3
label_col = "Project Name" if "Project Name" in df_f.columns else "Project ID"
_proj_ids = set(df_f["Project ID"].dropna()) if "Project ID" in df_f.columns else set()

# ── General-sheet column lookups used by Charts 2 & 3 ────────────────────────
_capex_col = next((c for c in df_gen.columns if c.strip().upper() == "CAPEX"), None)
_opex_col  = next((c for c in df_gen.columns if c.strip().upper() == "OPEX"),  None)


def _gen_totals(value_col):
    """Return DataFrame[Project ID, label_col, _value] from the General sheet (m$ USD)."""
    if value_col is None:
        return pd.DataFrame(columns=["Project ID", label_col, "_value"])
    cols = ["Project ID", label_col, value_col] + ([_curr_col] if _curr_col else [])
    df = df_f[cols].copy()
    df["_value"] = (pd.to_numeric(df[value_col], errors="coerce")
                    * (df[_curr_col].apply(get_fx) if _curr_col else 1.0))
    return df[df["_value"].notna() & (df["_value"] > 0)][["Project ID", label_col, "_value"]]


def _build_chart_df(gen_df, cost_phase_values):
    """
    Combine General-sheet totals with Cost-sheet breakdown.
    Projects that have Cost-sheet rows → stacked by Cost Group.
    Projects without Cost-sheet rows → single bar labelled 'Total'.
    """
    detail = cost_totals(_proj_ids, cost_phase_values, to_m_usd=True)
    detail = detail.merge(
        df_f[["Project ID", label_col]].drop_duplicates(), on="Project ID", how="left"
    )
    projects_with_detail = set(detail["Project ID"])

    fallback = gen_df[~gen_df["Project ID"].isin(projects_with_detail)].copy()
    fallback["Cost Group"] = "Total"

    combined = pd.concat(
        [detail[["Project ID", label_col, "Cost Group", "_value"]],
         fallback[["Project ID", label_col, "_value"]].assign(**{"Cost Group": fallback["Cost Group"]})],
        ignore_index=True,
    )
    return combined


# ══════════════════════════════════════════════════════════════════════════════
# CHART 2 — CAPEX by project
# ══════════════════════════════════════════════════════════════════════════════
st.header("CAPEX by Project")

if _capex_col is None:
    st.info('No "CAPEX" column found in the General sheet.')
else:
    df_cap2 = _build_chart_df(_gen_totals(_capex_col), ["CAPEX", "Hybrid Add-on"])
    if df_cap2.empty:
        st.info("No CAPEX data for the selected projects.")
    else:
        proj_order_cap = (
            df_cap2.groupby(label_col)["_value"].sum()
            .sort_values(ascending=False).index.tolist()
        )
        fig2 = px.bar(
            df_cap2,
            x=label_col, y="_value", color="Cost Group",
            barmode="stack",
            category_orders={label_col: proj_order_cap},
            labels={"_value": "CAPEX (m$ USD)", label_col: "Project"},
            color_discrete_sequence=COLOUR_SEQ,
            height=480,
        )
        fig2.update_layout(
            yaxis_title="CAPEX (m$ USD)", xaxis_tickangle=-35, legend_title="Cost Group"
        )
        st.plotly_chart(fig2, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# CHART 3 — OPEX by project
# ══════════════════════════════════════════════════════════════════════════════
st.header("OPEX by Project (m$/year)")

if _opex_col is None:
    st.info('No "OPEX" column found in the General sheet.')
else:
    df_opx2 = _build_chart_df(_gen_totals(_opex_col), ["OPEX"])
    if df_opx2.empty:
        st.info("No OPEX data for the selected projects.")
    else:
        proj_order_opx = (
            df_opx2.groupby(label_col)["_value"].sum()
            .sort_values(ascending=False).index.tolist()
        )
        fig3 = px.bar(
            df_opx2,
            x=label_col, y="_value", color="Cost Group",
            barmode="stack",
            category_orders={label_col: proj_order_opx},
            labels={"_value": "OPEX (m$/year)", label_col: "Project"},
            color_discrete_sequence=COLOUR_SEQ,
            height=480,
        )
        fig3.update_layout(
            yaxis_title="OPEX (m$/year)", xaxis_tickangle=-35, legend_title="Cost Group"
        )
        st.plotly_chart(fig3, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# SUMMARY TABLE — per-project component costs in USD
# ══════════════════════════════════════════════════════════════════════════════
st.header("Project Cost Summary")


def build_cost_summary():
    if not _proj_ids:
        return pd.DataFrame()

    # -- Per-component USD breakdown from Cost sheet --
    df = df_cost[df_cost[_c_pid].isin(_proj_ids)].copy() if _c_pid else pd.DataFrame()
    if not df.empty and _c_inc:
        df = df[df[_c_inc].astype(str).str.strip().str.upper() != "N"]

    if not df.empty and _c_val:
        df["_native"] = pd.to_numeric(df[_c_val], errors="coerce")
        df["_fx"]     = df[_cost_curr_col].apply(get_fx) if _cost_curr_col else 1.0
        df["_dc_kwp"] = (pd.to_numeric(df[_c_dc], errors="coerce") * 1000
                         if _c_dc else pd.Series(0.0, index=df.index))
        units = df[_c_unit].astype(str).str.strip() if _c_unit else pd.Series("", index=df.index)
        base  = df["_native"] * df["_fx"]
        df["_usd"] = np.where(
            units == "m$",    base * 1e6,
            np.where(
                units == "$/kWp", base * df["_dc_kwp"].where(df["_dc_kwp"] > 0),
                base,
            ),
        )
        comp_col = "Component" if "Component" in df.columns else (_c_grp or "Component")
        df["_col"] = df[_c_phase].astype(str) + " | " + df[comp_col].astype(str)
        pivot = (
            df[df["_usd"].notna()]
            .groupby([_c_pid, "_col"])["_usd"]
            .sum()
            .unstack(fill_value=np.nan)
            .reset_index()
            .rename(columns={_c_pid: "Project ID"})
        )
    else:
        pivot = pd.DataFrame(columns=["Project ID"])

    # -- Total CAPEX (USD): Cost sheet preferred, else General × 1e6 --
    cap_detail = (
        cost_totals(_proj_ids, ["CAPEX", "Hybrid Add-on"], to_m_usd=False)
        .groupby("Project ID")["_value"].sum()
    )
    cap_gen = (
        _gen_totals(_capex_col).set_index("Project ID")["_value"] * 1e6
        if _capex_col else pd.Series(dtype=float)
    )
    total_cap = cap_detail.combine_first(cap_gen).rename("Total CAPEX (USD)")

    # -- Total OPEX (USD/year): Cost sheet preferred, else General × 1e6 --
    opx_detail = (
        cost_totals(_proj_ids, ["OPEX"], to_m_usd=False)
        .groupby("Project ID")["_value"].sum()
    )
    opx_gen = (
        _gen_totals(_opex_col).set_index("Project ID")["_value"] * 1e6
        if _opex_col else pd.Series(dtype=float)
    )
    total_opx = opx_detail.combine_first(opx_gen).rename("Total OPEX (USD/year)")

    base_df = df_f[["Project ID", label_col]].drop_duplicates()
    result = (
        base_df
        .merge(pivot if not pivot.empty else pd.DataFrame(columns=["Project ID"]),
               on="Project ID", how="left")
        .merge(total_cap.reset_index(), on="Project ID", how="left")
        .merge(total_opx.reset_index(), on="Project ID", how="left")
        .drop(columns=["Project ID"])
        .set_index(label_col)
    )
    return result


tbl = build_cost_summary()
if tbl.empty:
    st.info("No cost data available for the selected projects.")
else:
    num_cols = tbl.select_dtypes("number").columns
    st.dataframe(
        tbl.style.format({c: "{:,.0f}" for c in num_cols}, na_rep="—"),
        use_container_width=True,
    )
