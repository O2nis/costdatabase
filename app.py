import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots

st.set_page_config(page_title="Cost Database Explorer", layout="wide")
st.title("Cost Database Explorer")

# ── Sidebar: file upload ───────────────────────────────────────────────────────
with st.sidebar:
    st.header("Data Source")
    uploaded_file = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx", "xls"])
    st.markdown("---")
    st.caption("If no file is uploaded the bundled example database is used.")


@st.cache_data
def load_data(source):
    xl = pd.ExcelFile(source)
    df_gen = pd.read_excel(xl, sheet_name="General")
    df_cost = pd.read_excel(xl, sheet_name="Cost")
    return df_gen, df_cost


try:
    source = uploaded_file if uploaded_file is not None else "Costdatabase v.3.1.xlsx"
    df_gen, df_cost = load_data(source)
except Exception as e:
    st.error(f"Could not load data: {e}")
    st.stop()

# ── Normalise Cost sheet to USD/MWp ───────────────────────────────────────────

def compute_usd_per_mwp(row):
    """Return USD/MWp regardless of original unit basis."""
    unit = row["Unit Basis"]
    usd = row["USD Value"]
    dc_mwp = row["Project DC MWp"]
    if pd.isna(usd):
        return np.nan
    if unit == "m$":
        if pd.isna(dc_mwp) or dc_mwp == 0:
            return np.nan
        return (usd * 1_000_000) / dc_mwp
    if unit == "$/kWp":
        return usd * 1_000          # $/kWp → $/MWp
    # $/year and anything else: not comparable on a per-MWp basis
    return np.nan


def compute_usd_per_year(row):
    """Return annualised USD value (for $/year rows)."""
    if row["Unit Basis"] == "$/year":
        return row["USD Value"]
    return np.nan


df_cost = df_cost.copy()
df_cost["USD_per_MWp"] = df_cost.apply(compute_usd_per_mwp, axis=1)
df_cost["USD_per_year"] = df_cost.apply(compute_usd_per_year, axis=1)

# ── Helpers ───────────────────────────────────────────────────────────────────

def numeric_cols(df):
    return df.select_dtypes(include="number").columns.tolist()


def all_cols(df):
    return df.columns.tolist()


CHART_TYPES = ["Bar", "Scatter", "Line"]
COLOUR_SEQ = px.colors.qualitative.Plotly

# ══════════════════════════════════════════════════════════════════════════════
# CHART 1 — General sheet
# ══════════════════════════════════════════════════════════════════════════════
st.header("Chart 1 — General Overview")

gen_all = all_cols(df_gen)
gen_num = numeric_cols(df_gen)

c1, c2, c3, c4 = st.columns([2, 2, 2, 2])
with c1:
    x_col = st.selectbox("X-axis", options=gen_all, index=gen_all.index("Project ID") if "Project ID" in gen_all else 0)
with c2:
    y1_default = gen_num[0] if gen_num else None
    y1_col = st.selectbox("Y-axis 1", options=gen_num, index=0)
with c3:
    y2_options = ["— none —"] + [c for c in gen_num if c != y1_col]
    y2_col = st.selectbox("Y-axis 2 (optional)", options=y2_options, index=0)
    y2_col = None if y2_col == "— none —" else y2_col
with c4:
    chart1_type = st.selectbox("Chart type", options=CHART_TYPES, index=0, key="ct1")

# Tag filter — values of the selected X-axis column
x_unique = sorted(df_gen[x_col].dropna().astype(str).unique().tolist())
tag_filter = st.multiselect(
    f"Filter by {x_col}",
    options=x_unique,
    default=x_unique,
    key="tag1",
)

df_gen_f = df_gen[df_gen[x_col].astype(str).isin(tag_filter)].copy() if tag_filter else df_gen.copy()

# Build chart 1
fig1 = make_subplots(specs=[[{"secondary_y": bool(y2_col)}]])

x_vals = df_gen_f[x_col].astype(str)

def add_trace(fig, x, y, name, ctype, secondary=False, color=None):
    kw = dict(name=name, marker_color=color) if color else dict(name=name)
    if ctype == "Bar":
        fig.add_trace(go.Bar(x=x, y=y, **kw), secondary_y=secondary)
    elif ctype == "Scatter":
        kw.pop("marker_color", None)
        fig.add_trace(go.Scatter(x=x, y=y, mode="markers", name=name,
                                 marker=dict(color=color, size=9)), secondary_y=secondary)
    else:  # Line
        kw.pop("marker_color", None)
        fig.add_trace(go.Scatter(x=x, y=y, mode="lines+markers", name=name,
                                 line=dict(color=color)), secondary_y=secondary)

add_trace(fig1, x_vals, df_gen_f[y1_col], y1_col, chart1_type, secondary=False,
          color=COLOUR_SEQ[0])
if y2_col:
    add_trace(fig1, x_vals, df_gen_f[y2_col], y2_col,
              "Line" if chart1_type == "Bar" else chart1_type,
              secondary=True, color=COLOUR_SEQ[1])

fig1.update_layout(
    height=460,
    barmode="group",
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    xaxis_title=x_col,
)
fig1.update_yaxes(title_text=y1_col, secondary_y=False)
if y2_col:
    fig1.update_yaxes(title_text=y2_col, secondary_y=True)

st.plotly_chart(fig1, use_container_width=True)

# Data table toggle
with st.expander("Show filtered General data"):
    st.dataframe(df_gen_f.reset_index(drop=True), use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# CHART 2 — Cost breakdown for the projects visible in Chart 1
# ══════════════════════════════════════════════════════════════════════════════
st.header("Chart 2 — Cost Breakdown by Component")
st.caption("Scope: projects currently shown in Chart 1 (after tag filter).")

# Derive project IDs from the filtered General view
active_pids = df_gen_f["Project ID"].dropna().unique().tolist()
df_cost_f = df_cost[df_cost["Project ID"].isin(active_pids)].copy()

# Optional: respect "Include?" flag
with st.sidebar:
    st.markdown("---")
    st.subheader("Cost chart options")
    only_included = st.checkbox('Only rows where Include? = "Y"', value=True)
    cost_phase_opts = df_cost["Cost Phase"].dropna().unique().tolist()
    selected_phases = st.multiselect("Cost phases", options=cost_phase_opts, default=cost_phase_opts)

if only_included and "Include?" in df_cost_f.columns:
    df_cost_f = df_cost_f[df_cost_f["Include?"] == "Y"]

df_cost_f = df_cost_f[df_cost_f["Cost Phase"].isin(selected_phases)]

if df_cost_f.empty:
    st.info("No cost data for the selected projects / filters.")
    st.stop()

# ── 2a  CAPEX / comparable rows → USD/MWp stacked bar ────────────────────────
df_capex = df_cost_f[df_cost_f["USD_per_MWp"].notna()].copy()
df_opex  = df_cost_f[df_cost_f["USD_per_year"].notna()].copy()

c1, c2 = st.columns([3, 1])
with c2:
    chart2_mode = st.radio("Group by", ["Cost Group", "Component"], index=1)
    show_pct = st.checkbox("Show as % of total", value=False)

if not df_capex.empty:
    st.subheader("CAPEX / Capital-equivalent — USD / MWp (DC)")

    pivot = (
        df_capex.groupby(["Project ID", chart2_mode])["USD_per_MWp"]
        .sum()
        .reset_index()
    )

    if show_pct:
        totals = pivot.groupby("Project ID")["USD_per_MWp"].transform("sum")
        pivot["USD_per_MWp"] = 100 * pivot["USD_per_MWp"] / totals
        y_label = "Share (%)"
    else:
        y_label = "USD / MWp (DC)"

    fig2 = px.bar(
        pivot,
        x="Project ID",
        y="USD_per_MWp",
        color=chart2_mode,
        barmode="stack",
        text_auto=".3s",
        labels={"USD_per_MWp": y_label},
        color_discrete_sequence=COLOUR_SEQ,
        height=460,
    )
    fig2.update_traces(textposition="inside", textfont_size=11)
    fig2.update_layout(
        yaxis_title=y_label,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    st.plotly_chart(fig2, use_container_width=True)
else:
    st.info("No CAPEX / comparable cost data in the current selection.")

# ── 2b  OPEX → USD/year bar ───────────────────────────────────────────────────
if not df_opex.empty:
    st.subheader("OPEX — USD / year")

    pivot_opex = (
        df_opex.groupby(["Project ID", chart2_mode])["USD_per_year"]
        .sum()
        .reset_index()
    )

    fig3 = px.bar(
        pivot_opex,
        x="Project ID",
        y="USD_per_year",
        color=chart2_mode,
        barmode="stack",
        text_auto=".3s",
        labels={"USD_per_year": "USD / year"},
        color_discrete_sequence=COLOUR_SEQ,
        height=380,
    )
    fig3.update_traces(textposition="inside", textfont_size=11)
    fig3.update_layout(
        yaxis_title="USD / year",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    st.plotly_chart(fig3, use_container_width=True)

# Detail table
with st.expander("Show cost detail table"):
    show_cols = ["Project ID", "Cost Phase", "Cost Group", "Component",
                 "USD Value", "Unit Basis", "USD_per_MWp", "USD_per_year",
                 "Source", "Confidence", "Comments"]
    show_cols = [c for c in show_cols if c in df_cost_f.columns]
    st.dataframe(df_cost_f[show_cols].reset_index(drop=True), use_container_width=True)
