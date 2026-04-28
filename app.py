import io
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

# Coerce columns that may have been read as object due to "-" or mixed entries
for _c in df_gen.columns:
    if df_gen[_c].dtype == object:
        _coerced = pd.to_numeric(df_gen[_c], errors="coerce")
        if _coerced.notna().any():
            df_gen[_c] = _coerced

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
    df["_fx"]     = pd.to_numeric(df[_cost_curr_col].apply(get_fx), errors="coerce").fillna(1.0) if _cost_curr_col else 1.0
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

# ── Column lookups shared across both tabs ────────────────────────────────────
_capex_col  = next((c for c in df_gen.columns if c.strip().upper() == "CAPEX"), None)
_opex_col   = next((c for c in df_gen.columns if c.strip().upper() == "OPEX"),  None)
_yr_col     = next((c for c in df_gen.columns if c.strip().lower() == "cod year"), None)
_epc_yr_col = next((c for c in df_gen.columns if "epc start" in c.strip().lower()), None)
_dc_gen_col = next((c for c in df_gen.columns
                    if "dc capacity" in c.lower() and "mwp" in c.lower()), None)
_comp_col   = next((c for c in df_gen.columns if "completeness" in c.lower()), None)
_lcoe_col   = next((c for c in df_gen.columns if "lcoe" in c.lower()), None)


# ── Trend data (Overview tab) ─────────────────────────────────────────────────

def build_trend_data(from_year):
    """
    Return long-format DataFrame[Project ID, Year, Series, _usd_mwp] where
    _usd_mwp is the cost normalised to $/MWp (or $/MWp/yr for OPEX).
    All unit bases (m$, $/kWp, $/year) are brought onto a single $/MWp axis.
    """
    _trend_yr = _epc_yr_col or _yr_col
    if _trend_yr is None or _dc_gen_col is None:
        return pd.DataFrame()

    yr_map = pd.to_numeric(df_gen.set_index("Project ID")[_trend_yr],   errors="coerce").dropna()
    dc_map = pd.to_numeric(df_gen.set_index("Project ID")[_dc_gen_col], errors="coerce").dropna()

    valid_ids = set(yr_map[yr_map >= from_year].index)
    if not valid_ids:
        return pd.DataFrame()

    rows = []

    # Cost-sheet detail
    if _c_pid and _c_phase and _c_val:
        df = df_cost[df_cost[_c_pid].isin(valid_ids)].copy()
        if _c_inc:
            df = df[df[_c_inc].astype(str).str.strip().str.upper() != "N"]

        if not df.empty:
            df["_native"] = pd.to_numeric(df[_c_val], errors="coerce")
            df["_fx"]     = df[_cost_curr_col].apply(get_fx) if _cost_curr_col else 1.0
            df["_dc_mwp"] = df[_c_pid].map(dc_map)
            units         = df[_c_unit].astype(str).str.strip() if _c_unit else pd.Series("", index=df.index)
            base          = df["_native"] * df["_fx"]

            df["_usd_mwp"] = np.where(
                units == "m$",    (base * 1e6) / df["_dc_mwp"].where(df["_dc_mwp"] > 0),
                np.where(
                    units == "$/kWp", base * 1000,
                    base / df["_dc_mwp"].where(df["_dc_mwp"] > 0),
                ),
            )
            df["Year"]   = df[_c_pid].map(yr_map)
            df["Series"] = df[_c_phase].astype(str) + " — " + df[_c_grp].astype(str)

            rows.append(
                df[df["_usd_mwp"].notna() & df["Year"].notna()][
                    [_c_pid, "Year", "Series", "_usd_mwp"]
                ].rename(columns={_c_pid: "Project ID"})
            )
            cost_ids = set(df[_c_pid].unique())
        else:
            cost_ids = set()
    else:
        cost_ids = set()

    # General-sheet fallback for projects with no Cost detail
    fallback_ids = valid_ids - cost_ids
    if fallback_ids and _capex_col and _curr_col:
        gen_fb = df_gen[df_gen["Project ID"].isin(fallback_ids)].copy()
        gen_fb["_yr"]     = gen_fb["Project ID"].map(yr_map)
        gen_fb["_dc_mwp"] = gen_fb["Project ID"].map(dc_map).replace(0, np.nan)
        gen_fb["_fx"]     = gen_fb[_curr_col].apply(get_fx)

        cap_mwp = (pd.to_numeric(gen_fb[_capex_col], errors="coerce")
                   * gen_fb["_fx"] * 1e6 / gen_fb["_dc_mwp"])
        fb_cap  = gen_fb[["Project ID", "_yr"]].copy()
        fb_cap["_usd_mwp"] = cap_mwp.values
        fb_cap["Series"]   = "CAPEX — Total"
        rows.append(fb_cap.rename(columns={"_yr": "Year"}).dropna(subset=["_usd_mwp"]))

        if _opex_col:
            opx_mwp = (pd.to_numeric(gen_fb[_opex_col], errors="coerce")
                       * gen_fb["_fx"] * 1e6 / gen_fb["_dc_mwp"])
            fb_opx  = gen_fb[["Project ID", "_yr"]].copy()
            fb_opx["_usd_mwp"] = opx_mwp.values
            fb_opx["Series"]   = "OPEX — Total"
            rows.append(fb_opx.rename(columns={"_yr": "Year"}).dropna(subset=["_usd_mwp"]))

    return pd.concat(rows, ignore_index=True) if rows else pd.DataFrame()


# ══════════════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════════════

tab1, tab2, tab3 = st.tabs(["📊 Overview", "🔍 Explorer", "📥 Import"])

# ──────────────────────────────────────────────────────────────────────────────
# TAB 1 — KPIs + cost trend
# ──────────────────────────────────────────────────────────────────────────────
with tab1:

    # KPI row
    n_proj      = len(df_gen)
    n_countries = df_gen["Country"].nunique() if "Country" in df_gen.columns else 0
    total_dc_gwp = (pd.to_numeric(df_gen[_dc_gen_col], errors="coerce").sum() / 1000
                    if _dc_gen_col else None)

    avg_comp = None
    if _comp_col:
        _cv = pd.to_numeric(df_gen[_comp_col], errors="coerce").dropna()
        if not _cv.empty:
            avg_comp = _cv.mean() * 100 if _cv.max() <= 1.0 else _cv.mean()

    avg_capex_kwp = None
    if _capex_col and _dc_gen_col and _curr_col:
        _cap = (pd.to_numeric(df_gen[_capex_col], errors="coerce")
                * df_gen[_curr_col].apply(get_fx) * 1e6)
        _dc  = (pd.to_numeric(df_gen[_dc_gen_col], errors="coerce").replace(0, np.nan) * 1000)
        _v   = (_cap / _dc).dropna()
        if not _v.empty:
            avg_capex_kwp = _v.mean()

    avg_opex_mwp = None
    if _opex_col and _dc_gen_col and _curr_col:
        _opx = (pd.to_numeric(df_gen[_opex_col], errors="coerce")
                * df_gen[_curr_col].apply(get_fx) * 1e6)
        _dc  = (pd.to_numeric(df_gen[_dc_gen_col], errors="coerce").replace(0, np.nan))
        _v   = (_opx / _dc).dropna()
        if not _v.empty:
            avg_opex_mwp = _v.mean()

    avg_lcoe = None
    if _lcoe_col:
        _lv = pd.to_numeric(df_gen[_lcoe_col], errors="coerce").dropna()
        if not _lv.empty:
            avg_lcoe = _lv.mean()

    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("Projects",        n_proj)
    k2.metric("Countries",       n_countries)
    k3.metric("Total DC",        f"{total_dc_gwp:.1f} GWp"   if total_dc_gwp   else "—")
    k4.metric("Avg CAPEX",       f"{avg_capex_kwp:,.0f} $/kWp" if avg_capex_kwp else "—")
    k5.metric("Avg LCOE",        f"{avg_lcoe:.2f} ¢/kWh"     if avg_lcoe       else "—")
    k6.metric("Avg Completeness",f"{avg_comp:.0f}%"           if avg_comp       else "—")

    # Year slider + trend chart
    _slider_yr_col = _epc_yr_col or _yr_col
    yr_vals = (pd.to_numeric(df_gen[_slider_yr_col], errors="coerce").dropna()
               if _slider_yr_col else pd.Series(dtype=float))

    if yr_vals.empty:
        st.info("No 'EPC Start Year' column found — trend chart unavailable.")
    else:
        yr_min, yr_max = int(yr_vals.min()), int(yr_vals.max())

        from_year = st.slider(
            "Show data from year",
            min_value=yr_min, max_value=yr_max, value=yr_min, step=1,
        )

        trend_raw = build_trend_data(from_year)

        if trend_raw.empty:
            st.info("No normalised cost data for the selected year range.")
        else:
            trend_agg = (
                trend_raw
                .groupby(["Year", "Series"], as_index=False)["_usd_mwp"]
                .mean()
            )
            trend_agg["Year"] = trend_agg["Year"].astype(int)

            fig_trend = px.line(
                trend_agg, x="Year", y="_usd_mwp",
                color="Series", markers=True,
                labels={"_usd_mwp": "Cost ($/MWp)", "Series": ""},
                color_discrete_sequence=COLOUR_SEQ,
                height=500,
            )
            fig_trend.update_traces(line_width=2, marker_size=8)
            fig_trend.update_layout(
                yaxis_title="Normalised Cost ($/MWp  ·  $/MWp/yr for OPEX)",
                xaxis=dict(dtick=1, tickangle=-30),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
                hovermode="x unified",
                margin=dict(t=80),
            )
            st.plotly_chart(fig_trend, use_container_width=True)
            st.caption(
                "Each point is the mean across all projects with that EPC Start year. "
                "Projects without Cost-sheet detail use General-sheet totals normalised by DC capacity."
            )

        # LCOE trend chart
        if _lcoe_col and _slider_yr_col:
            _lcoe_yr = df_gen[[_slider_yr_col, _lcoe_col]].copy()
            _lcoe_yr[_slider_yr_col] = pd.to_numeric(_lcoe_yr[_slider_yr_col], errors="coerce")
            _lcoe_yr[_lcoe_col]      = pd.to_numeric(_lcoe_yr[_lcoe_col],      errors="coerce")
            _lcoe_yr = (_lcoe_yr
                        .dropna()
                        .query(f"`{_slider_yr_col}` >= @from_year")
                        .groupby(_slider_yr_col, as_index=False)[_lcoe_col]
                        .mean()
                        .rename(columns={_slider_yr_col: "Year", _lcoe_col: "LCOE"}))
            _lcoe_yr["Year"] = _lcoe_yr["Year"].astype(int)

            if not _lcoe_yr.empty:
                st.subheader("LCOE by EPC Start Year")
                fig_lcoe = px.line(
                    _lcoe_yr, x="Year", y="LCOE",
                    markers=True,
                    labels={"LCOE": "LCOE (¢/kWh)", "Year": "EPC Start Year"},
                    color_discrete_sequence=[COLOUR_SEQ[5]],
                    height=360,
                )
                fig_lcoe.update_traces(line_width=2, marker_size=8)
                fig_lcoe.update_layout(
                    yaxis_title="LCOE (¢/kWh)",
                    xaxis=dict(dtick=1, tickangle=-30),
                    hovermode="x unified",
                )
                st.plotly_chart(fig_lcoe, use_container_width=True)
                st.caption(f"Mean LCOE across projects with data, per EPC Start year. "
                           f"{int(pd.to_numeric(df_gen[_lcoe_col], errors='coerce').notna().sum())} "
                           f"of {len(df_gen)} projects have LCOE data.")

# ──────────────────────────────────────────────────────────────────────────────
# TAB 2 — Explorer: bubble chart, CAPEX/OPEX bars, summary table
# ──────────────────────────────────────────────────────────────────────────────
with tab2:

    # ── CHART 1 — Bubble chart ────────────────────────────────────────────────
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

    df_f["_completeness"] = (df_f.notna().sum(axis=1) / len(df_f.columns) * 100).round(0).astype(int)

    # Unique display labels — computed now so df_plot inherits the column
    _base_label = "Project Name" if "Project Name" in df_f.columns else "Project ID"
    if "Project ID" in df_f.columns and _base_label == "Project Name":
        _dup_mask = df_f["Project Name"].astype(str).duplicated(keep=False)
        df_f["_proj_label"] = np.where(
            _dup_mask,
            df_f["Project Name"].astype(str) + " (" + df_f["Project ID"].astype(str) + ")",
            df_f["Project Name"].astype(str),
        )
        label_col = "_proj_label"
    else:
        label_col = _base_label

    # Warn about projects with no X-axis value (excluded before tag_filter is even built)
    _no_x_rows = df_gen[df_gen[x_col].isna()]
    if not _no_x_rows.empty:
        _id_col_x = next((c for c in ["Project Name", "Project ID"] if c in _no_x_rows.columns), None)
        _missing_x = _no_x_rows[_id_col_x].astype(str).tolist() if _id_col_x else _no_x_rows.index.astype(str).tolist()
        st.warning(
            f"**{len(_missing_x)} project(s) excluded — no value for '{x_col}':** "
            + ", ".join(_missing_x)
        )

    # Split: rows missing Y value are excluded from chart and listed separately
    _y_numeric = pd.to_numeric(df_f[y1_col], errors="coerce")
    _no_y_mask = _y_numeric.isna()
    df_no_y  = df_f[_no_y_mask].copy()
    df_plot  = df_f[~_no_y_mask].copy()

    if not df_no_y.empty:
        _id_col = next((c for c in ["Project Name", "Project ID"] if c in df_no_y.columns), None)
        _missing_names = df_no_y[_id_col].astype(str).tolist() if _id_col else df_no_y.index.astype(str).tolist()
        st.warning(
            f"**{len(_missing_names)} project(s) excluded — no value for '{y1_col}':** "
            + ", ".join(_missing_names)
        )

    hover_col = next((c for c in ["Project Name", "Project ID"] if c in df_plot.columns), None)

    _x_is_num = pd.api.types.is_numeric_dtype(df_gen[x_col])
    if _x_is_num:
        df_plot["_x_pos"] = pd.to_numeric(df_plot[x_col], errors="coerce")
        _cat_map = {}
    else:
        _cats = df_f[x_col].astype(str).unique().tolist()
        _cat_map = {c: i for i, c in enumerate(_cats)}
        rng = np.random.default_rng(42)
        df_plot["_x_pos"] = df_plot[x_col].astype(str).map(_cat_map) + rng.uniform(-0.3, 0.3, len(df_plot))

    _metric_col = size_col if size_col else y1_col
    _metric = pd.to_numeric(df_plot[_metric_col], errors="coerce").abs()
    _lo, _hi = _metric.dropna().min() if _metric.notna().any() else np.nan, \
               _metric.dropna().max() if _metric.notna().any() else np.nan

    if pd.notna(_hi) and pd.notna(_lo) and _hi > _lo:
        _pct = (_metric - _lo) / (_hi - _lo) * 100
        df_plot["_size"] = _pct.apply(lambda p: 16 if p <= 30 else (28 if p <= 60 else 44))
    else:
        df_plot["_size"] = 28
    df_plot["_size"] = df_plot["_size"].where(_metric.notna(), 16)
    df_plot["_has_metric"] = _metric.notna()

    def _fmt(s):
        return pd.to_numeric(s, errors="coerce").round(2).astype(str)

    def _make_hover(df):
        def _s(series):
            """Series → list of strings, NaN → '—'."""
            return [
                "—" if (v is None or (isinstance(v, float) and np.isnan(v))) else str(v)
                for v in series
            ]

        parts = []
        if hover_col:
            parts.append(_s(df[hover_col]))
        parts.append([f"<b>{x_col}:</b> {v}" for v in _s(df[x_col])])
        parts.append([f"<b>{y1_col}:</b> {v}" for v in _s(_fmt(df[y1_col]))])
        if size_col:
            parts.append([f"<b>{size_col}:</b> {v}" for v in _s(_fmt(df[size_col]))])
        parts.append([f"<b>Completeness:</b> {v}%" for v in df["_completeness"].astype(str)])
        return ["<br>".join(row) for row in zip(*parts)]

    fig1 = go.Figure()

    # Trace 1 — bubbles with a valid metric value (colored by colorscale)
    df_colored = df_plot[df_plot["_has_metric"]].copy()
    if not df_colored.empty:
        _cv = pd.to_numeric(df_colored[_metric_col], errors="coerce").abs()
        _labels_c = df_colored.apply(
            lambda r: f"{r['_completeness']}%" if r["_size"] >= 28 else "", axis=1
        ).tolist()
        fig1.add_trace(go.Scatter(
            x=df_colored["_x_pos"].tolist(),
            y=pd.to_numeric(df_colored[y1_col], errors="coerce").tolist(),
            mode="markers+text",
            name=_metric_col,
            text=_labels_c,
            textposition="middle center",
            textfont=dict(size=8, color="white", family="Arial Black"),
            hovertext=_make_hover(df_colored),
            hovertemplate="%{hovertext}<extra></extra>",
            marker=dict(
                size=df_colored["_size"].tolist(),
                sizemode="diameter",
                color=_cv.tolist(),
                colorscale="RdYlBu_r",
                showscale=True,
                colorbar=dict(title=dict(text=_metric_col, side="right"), thickness=14, len=0.8),
                cmin=float(_lo) if pd.notna(_lo) else None,
                cmax=float(_hi) if pd.notna(_hi) else None,
                opacity=0.85,
                line=dict(width=0.6, color="white"),
            ),
        ))

    # Trace 2 — bubbles with no metric value (gray)
    df_gray = df_plot[~df_plot["_has_metric"]].copy()
    if not df_gray.empty:
        _labels_g = df_gray.apply(
            lambda r: f"{r['_completeness']}%" if r["_size"] >= 28 else "", axis=1
        ).tolist()
        fig1.add_trace(go.Scatter(
            x=df_gray["_x_pos"].tolist(),
            y=pd.to_numeric(df_gray[y1_col], errors="coerce").tolist(),
            mode="markers+text",
            name=f"{_metric_col} (no data)",
            text=_labels_g,
            textposition="middle center",
            textfont=dict(size=8, color="white", family="Arial Black"),
            hovertext=_make_hover(df_gray),
            hovertemplate="%{hovertext}<extra></extra>",
            marker=dict(
                size=df_gray["_size"].tolist(),
                sizemode="diameter",
                color="#aaaaaa",
                opacity=0.55,
                line=dict(width=0.6, color="white"),
            ),
        ))

    if not _x_is_num and _cat_map:
        fig1.update_xaxes(tickvals=list(_cat_map.values()),
                          ticktext=list(_cat_map.keys()), tickangle=-30)

    fig1.update_layout(
        height=540, xaxis_title=x_col, yaxis_title=y1_col,
        hovermode="closest", margin=dict(r=120, b=80),
        showlegend=not df_gray.empty,
    )
    st.plotly_chart(fig1, use_container_width=True)

    with st.expander("Show filtered data"):
        st.dataframe(
            df_plot.drop(columns=["_x_pos", "_size", "_completeness", "_has_metric", "_proj_label"], errors="ignore")
                   .reset_index(drop=True),
            use_container_width=True,
        )

    # ── Shared for Charts 2, 3, table ─────────────────────────────────────────
    _proj_ids = set(df_plot["Project ID"].dropna()) if "Project ID" in df_plot.columns else set()

    def _gen_totals(value_col):
        """Return DataFrame[Project ID, label_col, _value] in m$ USD (General sheet)."""
        if value_col is None:
            return pd.DataFrame(columns=["Project ID", label_col, "_value"])
        cols = ["Project ID", label_col, value_col] + ([_curr_col] if _curr_col else [])
        df = df_plot[cols].copy()
        df["_value"] = (pd.to_numeric(df[value_col], errors="coerce")
                        * (df[_curr_col].apply(get_fx) if _curr_col else 1.0))
        return df[df["_value"].notna() & (df["_value"] > 0)][["Project ID", label_col, "_value"]]

    def _build_chart_df(gen_df, cost_phase_values):
        """
        Combine General-sheet totals with Cost-sheet breakdown.
        - Projects with Cost rows → stacked by Cost Group.
          Any gap vs the General total is added as 'Other Cost'.
        - Projects without Cost rows → single 'Total' bar from General.
        """
        detail = cost_totals(_proj_ids, cost_phase_values, to_m_usd=True)
        detail = detail.merge(
            df_plot[["Project ID", label_col]].drop_duplicates(), on="Project ID", how="left"
        )
        projects_with_detail = set(detail["Project ID"])

        gen_by_proj  = gen_df.groupby("Project ID")["_value"].sum()
        cost_by_proj = detail.groupby("Project ID")["_value"].sum()

        other_rows = []
        for pid in projects_with_detail:
            gen_val  = gen_by_proj.get(pid)
            if gen_val is None or pd.isna(gen_val):
                continue
            gap = float(gen_val) - float(cost_by_proj.get(pid, 0))
            if gap > 1e-9:
                lbl = df_plot.loc[df_plot["Project ID"] == pid, label_col]
                other_rows.append({
                    "Project ID": pid,
                    label_col:    lbl.iloc[0] if len(lbl) else pid,
                    "Cost Group": "Other Cost",
                    "_value":     gap,
                })

        fallback = gen_df[~gen_df["Project ID"].isin(projects_with_detail)].copy()
        fallback["Cost Group"] = "Total"

        parts = [detail[["Project ID", label_col, "Cost Group", "_value"]]]
        if other_rows:
            parts.append(pd.DataFrame(other_rows))
        parts.append(
            fallback[["Project ID", label_col, "_value"]]
            .assign(**{"Cost Group": fallback["Cost Group"]})
        )
        return pd.concat(parts, ignore_index=True)

    # ── CHART 2 — CAPEX ───────────────────────────────────────────────────────
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

    # ── CHART 3 — OPEX ────────────────────────────────────────────────────────
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

    # ── SUMMARY TABLE ─────────────────────────────────────────────────────────
    st.header("Project Cost Summary")

    def build_cost_summary():
        if not _proj_ids:
            return pd.DataFrame()

        df = df_cost[df_cost[_c_pid].isin(_proj_ids)].copy() if _c_pid else pd.DataFrame()
        if not df.empty and _c_inc:
            df = df[df[_c_inc].astype(str).str.strip().str.upper() != "N"]

        if not df.empty and _c_val:
            df["_native"] = pd.to_numeric(df[_c_val], errors="coerce")
            df["_fx"]     = pd.to_numeric(df[_cost_curr_col].apply(get_fx), errors="coerce").fillna(1.0) if _cost_curr_col else 1.0
            df["_dc_kwp"] = (pd.to_numeric(df[_c_dc], errors="coerce") * 1000
                             if _c_dc else pd.Series(0.0, index=df.index))
            units = df[_c_unit].astype(str).str.strip() if _c_unit else pd.Series("", index=df.index)
            base  = df["_native"] * df["_fx"]
            df["_usd"] = np.where(
                units == "m$",    base * 1e6,
                np.where(units == "$/kWp", base * df["_dc_kwp"].where(df["_dc_kwp"] > 0), base),
            )
            comp_col = "Component" if "Component" in df.columns else (_c_grp or "Component")
            df["_col"] = df[_c_phase].astype(str) + " | " + df[comp_col].astype(str)
            pivot = (
                df[df["_usd"].notna()]
                .groupby([_c_pid, "_col"])["_usd"].sum()
                .unstack(fill_value=np.nan).reset_index()
                .rename(columns={_c_pid: "Project ID"})
            )
        else:
            pivot = pd.DataFrame(columns=["Project ID"])

        cap_detail = (cost_totals(_proj_ids, ["CAPEX", "Hybrid Add-on"], to_m_usd=False)
                      .groupby("Project ID")["_value"].sum())
        cap_gen    = (_gen_totals(_capex_col).set_index("Project ID")["_value"] * 1e6
                      if _capex_col else pd.Series(dtype=float))
        total_cap  = cap_detail.combine_first(cap_gen).rename("Total CAPEX (USD)")

        opx_detail = (cost_totals(_proj_ids, ["OPEX"], to_m_usd=False)
                      .groupby("Project ID")["_value"].sum())
        opx_gen    = (_gen_totals(_opex_col).set_index("Project ID")["_value"] * 1e6
                      if _opex_col else pd.Series(dtype=float))
        total_opx  = opx_detail.combine_first(opx_gen).rename("Total OPEX (USD/year)")

        base_df = df_plot[["Project ID", label_col]].drop_duplicates()
        return (
            base_df
            .merge(pivot if not pivot.empty else pd.DataFrame(columns=["Project ID"]),
                   on="Project ID", how="left")
            .merge(total_cap.reset_index(), on="Project ID", how="left")
            .merge(total_opx.reset_index(), on="Project ID", how="left")
            .drop(columns=["Project ID"])
            .set_index(label_col)
        )

    tbl = build_cost_summary()
    if tbl.empty:
        st.info("No cost data available for the selected projects.")
    else:
        num_cols = tbl.select_dtypes("number").columns
        st.dataframe(
            tbl.style.format({c: "{:,.0f}" for c in num_cols}, na_rep="—"),
            use_container_width=True,
        )

# ──────────────────────────────────────────────────────────────────────────────
# TAB 3 — Import
# ──────────────────────────────────────────────────────────────────────────────
with tab3:
    st.header("Import New Data")
    st.caption(
        "Upload an Excel file with the same two-sheet structure (General + Cost). "
        "Rows whose Project ID already exists in the database are skipped."
    )

    @st.cache_data
    def load_import_file(f):
        xl = pd.ExcelFile(f)
        sheets = xl.sheet_names
        dg = pd.read_excel(xl, sheet_name="General") if "General" in sheets else pd.DataFrame()
        dc = pd.read_excel(xl, sheet_name="Cost")    if "Cost"    in sheets else pd.DataFrame()
        return dg, dc

    import_file = st.file_uploader(
        "Upload import file (.xlsx)", type=["xlsx", "xls"], key="import_uploader"
    )

    if import_file is None:
        st.info("No file uploaded yet.")
    else:
        try:
            df_imp_gen, df_imp_cost = load_import_file(import_file)
        except Exception as e:
            st.error(f"Could not read import file: {e}")
            st.stop()

        if df_imp_gen.empty and df_imp_cost.empty:
            st.error("Import file has no 'General' or 'Cost' sheet.")
            st.stop()

        # ── Identify new rows ─────────────────────────────────────────────────
        pid_col_gen  = next((c for c in df_imp_gen.columns  if c.strip() == "Project ID"), None)
        pid_col_cost = next((c for c in df_imp_cost.columns if c.strip() == "Project ID"), None)

        existing_gen_pids  = set(df_gen["Project ID"].dropna().astype(str))  if "Project ID" in df_gen.columns  else set()
        existing_cost_pids = set(df_cost[_c_pid].dropna().astype(str))       if _c_pid else set()

        if pid_col_gen and not df_imp_gen.empty:
            new_gen_mask = ~df_imp_gen[pid_col_gen].astype(str).isin(existing_gen_pids)
            df_add_gen   = df_imp_gen[new_gen_mask].copy()
            dup_gen_pids = df_imp_gen[pid_col_gen].astype(str)[~new_gen_mask].tolist()
        else:
            df_add_gen   = df_imp_gen.copy()
            dup_gen_pids = []

        new_pids = set(df_add_gen[pid_col_gen].astype(str)) if pid_col_gen and not df_add_gen.empty else set()

        if pid_col_cost and not df_imp_cost.empty:
            # Keep Cost rows only for projects being added (or not yet in Cost sheet)
            cost_mask    = ~df_imp_cost[pid_col_cost].astype(str).isin(existing_cost_pids)
            df_add_cost  = df_imp_cost[cost_mask & df_imp_cost[pid_col_cost].astype(str).isin(
                new_pids | (set(df_imp_cost[pid_col_cost].astype(str)) - existing_cost_pids)
            )].copy()
        else:
            df_add_cost = df_imp_cost.copy()

        if dup_gen_pids:
            st.warning(
                f"**{len(dup_gen_pids)} Project ID(s) already in database — skipped:** "
                + ", ".join(dup_gen_pids)
            )

        # ── KPIs ──────────────────────────────────────────────────────────────
        n_new   = len(df_add_gen)
        n_cost  = len(df_add_cost)

        new_countries = []
        if "Country" in df_add_gen.columns and "Country" in df_gen.columns:
            new_countries = sorted(
                set(df_add_gen["Country"].dropna().astype(str))
                - set(df_gen["Country"].dropna().astype(str))
            )

        dc_col_imp = next((c for c in df_add_gen.columns
                           if "dc capacity" in c.lower() and "mwp" in c.lower()), None)
        new_dc_mwp = (pd.to_numeric(df_add_gen[dc_col_imp], errors="coerce").sum()
                      if dc_col_imp else None)

        capex_col_imp = next((c for c in df_add_gen.columns if c.strip().upper() == "CAPEX"), None)
        curr_col_imp  = next((c for c in df_add_gen.columns if c.strip().lower() == "native currency"), None)
        new_avg_capex = None
        if capex_col_imp and dc_col_imp:
            _cap = pd.to_numeric(df_add_gen[capex_col_imp], errors="coerce")
            if curr_col_imp:
                _cap = _cap * df_add_gen[curr_col_imp].apply(get_fx) * 1e6
            _dc  = pd.to_numeric(df_add_gen[dc_col_imp], errors="coerce").replace(0, np.nan) * 1000
            _v   = (_cap / _dc).dropna()
            if not _v.empty:
                new_avg_capex = _v.mean()

        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("New projects",      n_new)
        k2.metric("New cost entries",  n_cost)
        k3.metric("New DC capacity",   f"{new_dc_mwp:,.0f} MWp" if new_dc_mwp else "—")
        k4.metric("New countries",     len(new_countries) if new_countries else 0)
        k5.metric("Avg CAPEX (new)",   f"{new_avg_capex:,.0f} $/kWp" if new_avg_capex else "—")

        if new_countries:
            st.caption("New countries: " + ", ".join(new_countries))

        # ── Preview tables ────────────────────────────────────────────────────
        st.subheader(f"General sheet — {n_new} new row(s)")
        if df_add_gen.empty:
            st.info("No new rows to add to the General sheet.")
        else:
            st.dataframe(df_add_gen.reset_index(drop=True), use_container_width=True)

        st.subheader(f"Cost sheet — {n_cost} new row(s)")
        if df_add_cost.empty:
            st.info("No new rows to add to the Cost sheet.")
        else:
            st.dataframe(df_add_cost.reset_index(drop=True), use_container_width=True)

        # ── Action buttons ────────────────────────────────────────────────────
        if df_add_gen.empty and df_add_cost.empty:
            st.info("Nothing new to import — all Project IDs already exist in the database.")
        else:
            merged_gen  = pd.concat([df_gen,  df_add_gen],  ignore_index=True)
            merged_cost = pd.concat([df_cost, df_add_cost], ignore_index=True)

            # Build merged xlsx in memory for download
            _buf = io.BytesIO()
            with pd.ExcelWriter(_buf, engine="openpyxl") as _writer:
                merged_gen.to_excel(_writer,  sheet_name="General", index=False)
                merged_cost.to_excel(_writer, sheet_name="Cost",    index=False)
            _buf.seek(0)

            btn_col1, btn_col2 = st.columns(2)

            btn_col2.download_button(
                label="⬇ Download merged database",
                data=_buf,
                file_name="Costdatabase_merged.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            # Write to disk only available when using the bundled local file
            _can_write = uploaded_file is None
            if btn_col1.button(
                "💾 Write to original database",
                disabled=not _can_write,
                help="Only available when using the bundled database (no file uploaded in the sidebar).",
            ):
                st.session_state["_confirm_write"] = True

            if st.session_state.get("_confirm_write"):
                st.warning(
                    "This will **permanently overwrite** `Costdatabase v.3.1.xlsx`. "
                    "Download a backup first if needed."
                )
                c_yes, c_no = st.columns(2)
                if c_yes.button("✅ Confirm — write now"):
                    with pd.ExcelWriter("Costdatabase v.3.1.xlsx", engine="openpyxl") as _w:
                        merged_gen.to_excel(_w,  sheet_name="General", index=False)
                        merged_cost.to_excel(_w, sheet_name="Cost",    index=False)
                    st.session_state.pop("_confirm_write", None)
                    st.cache_data.clear()
                    st.success("Database updated. Reloading…")
                    st.rerun()
                if c_no.button("❌ Cancel"):
                    st.session_state.pop("_confirm_write", None)
                    st.rerun()
