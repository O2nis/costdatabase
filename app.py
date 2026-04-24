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
_curr_col = next((c for c in df_gen.columns if c.strip().lower() == "native currency"), None)
_fx_col   = next((c for c in df_gen.columns if c.strip().lower() == "fx to usd"), None)

fx_rates = {"USD": 1.0}   # currency → rate applied to convert native → USD

if _curr_col:
    non_usd = [c for c in df_gen[_curr_col].dropna().unique() if str(c).strip().upper() != "USD"]
    if non_usd:
        with st.sidebar:
            st.markdown("---")
            st.subheader("Currency conversion to USD")
            for curr in sorted(non_usd):
                curr_str = str(curr).strip().upper()
                # Suggest rate from FX to USD column if present, else 1.0
                if _fx_col:
                    mask = df_gen[_curr_col].astype(str).str.strip().str.upper() == curr_str
                    suggested = df_gen.loc[mask, _fx_col].apply(pd.to_numeric, errors="coerce").mean()
                    default = float(suggested) if pd.notna(suggested) else 1.0
                else:
                    default = 1.0
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

# ══════════════════════════════════════════════════════════════════════════════
# CHART 2 — CAPEX by project (General sheet, FX-converted to USD)
# ══════════════════════════════════════════════════════════════════════════════
st.header("CAPEX by Project")

_capex_col = next((c for c in df_gen.columns if c.strip().upper() == "CAPEX"), None)

if _capex_col is None:
    st.info('No "CAPEX" column found in the General sheet.')
else:
    df_cap = df_f[[label_col, _capex_col] + ([_curr_col] if _curr_col else [])].copy()
    df_cap["_capex_native"] = pd.to_numeric(df_cap[_capex_col], errors="coerce")
    df_cap["_fx"] = df_cap[_curr_col].apply(get_fx) if _curr_col else 1.0
    # CAPEX column is in m$ native → convert to m$ USD
    df_cap["_capex_usd"] = df_cap["_capex_native"] * df_cap["_fx"]
    df_cap = df_cap[df_cap["_capex_usd"].notna() & (df_cap["_capex_usd"] > 0)] \
               .sort_values("_capex_usd", ascending=False)

    if df_cap.empty:
        st.info("No CAPEX data for the selected projects.")
    else:
        fig2 = px.bar(
            df_cap,
            x=label_col,
            y="_capex_usd",
            text_auto=".3s",
            color="_capex_usd",
            color_continuous_scale="RdYlBu_r",
            labels={"_capex_usd": "CAPEX (m$ USD)", label_col: "Project"},
            height=440,
        )
        fig2.update_traces(textposition="outside", textfont_size=10)
        fig2.update_layout(
            yaxis_title="CAPEX (m$ USD)",
            xaxis_tickangle=-35,
            coloraxis_showscale=False,
        )
        st.plotly_chart(fig2, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# CHART 3 — OPEX by project (General sheet, FX-converted to USD/year)
# ══════════════════════════════════════════════════════════════════════════════
st.header("OPEX by Project ($/year)")

_opex_col = next((c for c in df_gen.columns if c.strip().upper() == "OPEX"), None)

if _opex_col is None:
    st.info('No "OPEX" column found in the General sheet.')
else:
    df_opx = df_f[[label_col, _opex_col] + ([_curr_col] if _curr_col else [])].copy()
    df_opx["_opex_native"] = pd.to_numeric(df_opx[_opex_col], errors="coerce")
    df_opx["_fx"] = df_opx[_curr_col].apply(get_fx) if _curr_col else 1.0
    df_opx["_opex_usd"] = df_opx["_opex_native"] * df_opx["_fx"]
    df_opx = df_opx[df_opx["_opex_usd"].notna() & (df_opx["_opex_usd"] > 0)] \
               .sort_values("_opex_usd", ascending=False)

    if df_opx.empty:
        st.info("No OPEX data for the selected projects.")
    else:
        fig3 = px.bar(
            df_opx,
            x=label_col,
            y="_opex_usd",
            text_auto=".3s",
            color="_opex_usd",
            color_continuous_scale="RdYlBu_r",
            labels={"_opex_usd": "OPEX (USD/year)", label_col: "Project"},
            height=440,
        )
        fig3.update_traces(textposition="outside", textfont_size=10)
        fig3.update_layout(
            yaxis_title="OPEX (USD/year)",
            xaxis_tickangle=-35,
            coloraxis_showscale=False,
        )
        st.plotly_chart(fig3, use_container_width=True)
