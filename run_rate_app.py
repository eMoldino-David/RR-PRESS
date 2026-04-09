import streamlit as st
import pandas as pd
import numpy as np
import warnings
import streamlit.components.v1 as components
from datetime import datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import run_rate_utils as rr_utils

# ==============================================================================
# --- 1. PAGE CONFIG & SETUP ---
# ==============================================================================

warnings.filterwarnings("ignore", category=FutureWarning)

try:
    st.set_page_config(layout="wide", page_title="Run Rate Analysis Dashboard")
except Exception:
    pass

# Multiselect chip colours — matching Capacity Risk app
st.markdown("""
<style>
    span[data-baseweb="tag"] {
        background-color: #34495e !important;
        color: #ecf0f1 !important;
    }
    span[data-baseweb="tag"] svg {
        fill: #ecf0f1 !important;
    }
</style>
""", unsafe_allow_html=True)


# ==============================================================================
# --- 2. LOGIN ---
# ==============================================================================

def check_password():
    """Returns True if the user has entered the correct password."""
    if st.session_state.get("password_correct", False):
        return True

    st.header("🔒 Protected Internal Tool")
    password_input = st.text_input("Enter Company Password", type="password")

    if password_input:
        if password_input == st.secrets["APP_PASSWORD"]:
            st.session_state["password_correct"] = True
            st.rerun()
        else:
            st.error("😕 Password incorrect")

    return False


if not check_password():
    st.stop()


# ==============================================================================
# --- 3. UI RENDERING FUNCTIONS ---
# ==============================================================================

def render_risk_tower(df_all_tools, run_interval_hours, min_shots_filter, tolerance, downtime_gap_tolerance):
    """Renders the Risk Tower tab."""
    st.title("Run Rate Risk Tower")
    st.info(
        "This tower analyses performance over the last 4 weeks, identifying tools "
        "that require attention. Tools with the lowest scores are at the highest risk."
    )

    with st.expander("ℹ️ How the Risk Tower Works"):
        st.markdown(f"""
        The Risk Tower evaluates each tool based on its performance over its own most recent
        4-week period of operation. Here's how the metrics are calculated:

        - **Analysis Period**: Shows the exact 4-week date range used for each tool's analysis.
        - **Data Filters**:
            - **Run Interval**: Gaps longer than {run_interval_hours} hours are treated as breaks between runs.
            - **Min Shots**: Production runs with fewer than {min_shots_filter} shots are excluded.
        - **Risk Score**: A performance indicator from 0–100.
            - Starts with the tool's overall **Stability Index (%)** for the period.
            - A **20-point penalty** is applied if stability shows a declining trend.
        - **Primary Risk Factor**: Identifies the main issue affecting performance:
            1. **Declining Trend** — if stability is worsening over time.
            2. **High MTTR** — if avg stop duration is significantly above peer average.
            3. **Frequent Stops** — if MTBF is significantly below peer average.
            4. **Low Stability** — if none of the above but stability is still low.
        - **Color Coding**:
            - <span style='background-color:#ff6961;color:black;padding:2px 5px;border-radius:5px;'>Red (0–50)</span>: High Risk
            - <span style='background-color:#ffb347;color:black;padding:2px 5px;border-radius:5px;'>Orange (51–70)</span>: Medium Risk
            - <span style='background-color:#77dd77;color:black;padding:2px 5px;border-radius:5px;'>Green (>70)</span>: Low Risk
        """, unsafe_allow_html=True)

    risk_df = rr_utils.calculate_risk_scores(df_all_tools, run_interval_hours, min_shots_filter, tolerance, downtime_gap_tolerance)

    if risk_df.empty:
        st.warning("Not enough data across multiple tools in the last 4 weeks to generate a risk tower.")
        return

    def style_risk(row):
        score = row['Risk Score']
        if score > 70:
            color = rr_utils.PASTEL_COLORS['green']
        elif score > 50:
            color = rr_utils.PASTEL_COLORS['orange']
        else:
            color = rr_utils.PASTEL_COLORS['red']
        return [f'background-color: {color}' for _ in row]

    cols_order = ['Tool ID', 'Analysis Period', 'Risk Score',
                  'Primary Risk Factor', 'Weekly Stability', 'Details']
    display_df = risk_df[[col for col in cols_order if col in risk_df.columns]]
    st.dataframe(
        display_df.style.apply(style_risk, axis=1).format({'Risk Score': '{:.0f}'}),
        width='stretch', hide_index=True
    )


def render_trends_tab(df_tool, tool_id_selection, tolerance, downtime_gap_tolerance,
                      run_interval_hours, min_shots_filter, key_prefix=''):
    """Renders the Trends Analysis tab."""
    _k = key_prefix  # short alias for prefixing widget keys
    st.header("Historical Performance Trends")
    st.info(
        f"Trends are calculated using 'Run-Based' logic. Gaps larger than "
        f"{run_interval_hours} hours are excluded from the timeline to provide "
        f"accurate stability metrics."
    )

    col_ctrl, _ = st.columns([1, 3])
    with col_ctrl:
        trend_freq = st.selectbox("Select Trend Frequency", ["Daily", "Weekly", "Monthly"],
                                  key=f"{_k}trend_freq_select")

    with st.expander("ℹ️ About Trends Metrics"):
        st.markdown("""
        - **Stability Index (%)**: Percentage of run time spent in production.
        - **Efficiency (%)**: Percentage of shots that were normal (non-stops).
        - **MTTR (min)**: Mean Time To Repair (avg stop duration).
        - **MTBF (min)**: Mean Time Between Failures (avg uptime between stops).
        - **Total Shots**: Total output for the period.
        - **Stop Events**: Number of times the machine stopped.
        """)

    trend_data = []

    if trend_freq == "Daily":
        period_name = "Date"
    elif trend_freq == "Weekly":
        period_name = "Week"
    else:
        period_name = "Month"

    # #9 fix: run ONE global processing pass on the full tool dataset so that
    # mode_ct, stop_flag and run boundaries are consistent across all periods.
    # df_tool is always raw (a slice of df_filtered) so we always need this pass.
    # Downstream calculate_run_summaries calls use pre_processed=True so they
    # slice the result without recomputing mode_ct from the period subset.
    _prep = rr_utils.RunRateCalculator(
        df_tool, tolerance, downtime_gap_tolerance,
        analysis_mode='aggregate', run_interval_hours=run_interval_hours
    )
    df_tool_proc = _prep.results.get("processed_df", df_tool)

    if trend_freq == "Daily":
        grouper = df_tool_proc.groupby(df_tool_proc['shot_time'].dt.date)
    elif trend_freq == "Weekly":
        grouper = df_tool_proc.groupby(df_tool_proc['shot_time'].dt.to_period('W'))
    else:
        grouper = df_tool_proc.groupby(df_tool_proc['shot_time'].dt.to_period('M'))

    for period, df_period in grouper:
        if df_period.empty:
            continue

        run_summaries = rr_utils.calculate_run_summaries(
            df_period, tolerance, downtime_gap_tolerance,
            run_interval_hours=run_interval_hours,
            pre_processed=True
        )
        if run_summaries.empty:
            continue

        run_summaries = run_summaries[run_summaries['total_shots'] >= min_shots_filter]
        if run_summaries.empty:
            continue

        total_runtime = run_summaries['total_runtime_sec'].sum()
        prod_time = run_summaries['production_time_sec'].sum()
        downtime = run_summaries['downtime_sec'].sum()
        stops = run_summaries['stops'].sum()
        total_shots = run_summaries['total_shots'].sum()
        normal_shots = run_summaries['normal_shots'].sum()

        stability = (prod_time / total_runtime * 100) if total_runtime > 0 else 0
        efficiency = (normal_shots / total_shots * 100) if total_shots > 0 else 0
        mttr = (downtime / 60 / stops) if stops > 0 else 0
        mtbf = (prod_time / 60 / stops) if stops > 0 else (prod_time / 60)

        if trend_freq == "Daily":
            label = period.strftime('%Y-%m-%d')
        elif trend_freq == "Weekly":
            label = f"W{period.week} {period.year}"
        else:
            label = period.strftime('%B %Y')

        trend_data.append({
            period_name: label,
            'SortKey': period if trend_freq == "Daily" else period.start_time,
            'Stability Index (%)': stability,
            'Efficiency (%)': efficiency,
            'MTTR (min)': mttr,
            'MTBF (min)': mtbf,
            'Total Shots': total_shots,
            'Normal Shots': normal_shots,  # #11 fix: was missing, caused blank in PPTX
            'Stop Events': stops,
            'Production Time (h)': prod_time / 3600,
            'Downtime (h)': downtime / 3600,
        })

    if not trend_data:
        st.warning("No data found for the selected tool to generate trends.")
        return

    df_trends = (pd.DataFrame(trend_data)
                 .sort_values('SortKey', ascending=True)
                 .drop(columns=['SortKey']))

    st.dataframe(
        df_trends.style.format({
            'Stability Index (%)': '{:.1f}', 'Efficiency (%)': '{:.1f}',
            'MTTR (min)': '{:.1f}', 'MTBF (min)': '{:.1f}',
            'Total Shots': '{:,.0f}', 'Stop Events': '{:,.0f}',
            'Production Time (h)': '{:.1f}', 'Downtime (h)': '{:.1f}',
        }).background_gradient(subset=['Stability Index (%)'],
                               cmap='RdYlGn', vmin=0, vmax=100),
        width='stretch'
    )

    # ── Weekly comparison report download ────────────────────────────────────
    if trend_freq == "Weekly" and not df_trends.empty:
        st.markdown("---")
        col_dl, col_info = st.columns([1, 3])
        with col_dl:
            try:
                pptx_bytes = rr_utils.generate_weekly_comparison_pptx(
                    df_trends, tool_id_selection
                )
                st.download_button(
                    label="📊 Download Weekly Report (.pptx)",
                    data=pptx_bytes,
                    file_name=(
                        f"Weekly_Report_{tool_id_selection.replace(' ', '_')}_"
                        f"{pd.Timestamp.now().strftime('%Y%m%d')}.pptx"
                    ),
                    mime="application/vnd.openxmlformats-officedocument"
                          ".presentationml.presentation",
                    width='stretch',
                )
            except Exception as e:
                st.error(f"Could not generate report: {e}")
        with col_info:
            st.caption(
                "Generates a single-slide PowerPoint with a week-on-week "
                "comparison table including delta % vs the prior week."
            )
        st.markdown("---")

    st.subheader("Visual Trend")
    metric_to_plot = st.selectbox(
        "Select Metric to Visualize",
        ['Stability Index (%)', 'Efficiency (%)', 'MTTR (min)', 'MTBF (min)', 'Total Shots'],
        key=f"{_k}trend_viz_select"
    )

    fig = px.line(df_trends.sort_index(ascending=True), x=period_name,
                  y=metric_to_plot, markers=True,
                  title=f"{metric_to_plot} Trend ({trend_freq})")

    if '%)' in metric_to_plot:
        for y0, y1, c in [(0, 50, rr_utils.PASTEL_COLORS['red']),
                          (50, 70, rr_utils.PASTEL_COLORS['orange']),
                          (70, 100, rr_utils.PASTEL_COLORS['green'])]:
            fig.add_shape(type="rect", xref="paper", x0=0, x1=1, y0=y0, y1=y1,
                          fillcolor=c, opacity=0.1, layer="below", line_width=0)
        fig.update_yaxes(range=[0, 105])

    st.plotly_chart(fig, width='stretch')


def render_dashboard(df_tool, tool_id_selection, tolerance, downtime_gap_tolerance,
                     run_interval_hours, show_approved_ct, min_shots_filter,
                     key_prefix=''):
    """Renders the main Run Rate Dashboard tab."""
    _k = key_prefix  # short alias for prefixing widget keys

    analysis_level = st.radio(
        "Select Analysis Level",
        options=["Daily (by Run)", "Weekly (by Run)", "Monthly (by Run)", "Custom Period (by Run)"],
        horizontal=True,
        key=f"{_k}rr_analysis_level"
    )

    # Press / stamping mode — auto-detected from tooling_type, overridable by toggle
    _press_auto = (
        df_tool['tooling_type'].str.lower().str.contains('press|stamp', na=False).any()
        if 'tooling_type' in df_tool.columns else False
    )
    press_mode = st.toggle(
        "Press / Stamping Mode",
        value=bool(_press_auto),
        key=f"{_k}rr_press_mode",
        help="Enables stroke rate charts (SPM/SPH) for press and stamping tools."
    )

    if press_mode:
        stroke_unit = st.radio(
            "Mode Display Unit",
            options=["SPM", "SPH", "CT"],
            index=0,
            horizontal=True,
            key=f"{_k}rr_stroke_unit",
            help="SPM = Strokes Per Minute  |  SPH = Strokes Per Hour  |  CT = Cycle Time (sec)"
        )
    else:
        stroke_unit = "CT"

    st.markdown("---")

    # ------------------------------------------------------------------
    # FIX: get_processed_data now takes the user's tolerance params so
    # that mode_ct, lower_limit, and upper_limit columns in df_processed
    # are always computed with the correct slider values.
    # The cache key includes tolerance + downtime_gap_tolerance so any
    # slider change automatically triggers a fresh computation.
    # ------------------------------------------------------------------
    @st.cache_data(show_spinner="Performing initial data processing...")
    def get_processed_data(df, interval_hours, tolerance, downtime_gap_tolerance,
                           _schema_version=2):
        """
        Single authoritative processing pass over the FULL tool dataset.
        _schema_version: bump this when internal column names change to bust the cache.
        v2: lower_limit/upper_limit → mode_lower/mode_upper, ACTUAL CT → actual_ct
        """
        base_calc = rr_utils.RunRateCalculator(
            df, tolerance, downtime_gap_tolerance,
            analysis_mode='aggregate', run_interval_hours=interval_hours
        )
        df_processed = base_calc.results.get("processed_df", pd.DataFrame())
        if not df_processed.empty:
            df_processed['week'] = df_processed['shot_time'].dt.isocalendar().week
            df_processed['year'] = df_processed['shot_time'].dt.isocalendar().year
            df_processed['date'] = df_processed['shot_time'].dt.date
            df_processed['month'] = df_processed['shot_time'].dt.to_period('M')
        return df_processed

    df_processed = get_processed_data(
        df_tool, run_interval_hours, tolerance, downtime_gap_tolerance
    )

    detailed_view = st.toggle("Show Detailed Analysis", value=True, key=f"{_k}rr_detailed_view")

    if df_processed.empty:
        st.error(f"Could not process data for {tool_id_selection}. "
                 f"Check file format or data range.")
        st.stop()

    st.markdown(f"### {tool_id_selection} Overview")

    mode = 'by_run'
    df_view = pd.DataFrame()
    info_placeholder = None
    info_base_text = ""

    # ------------------------------------------------------------------
    # Date / period selection
    # ------------------------------------------------------------------
    if "Daily" in analysis_level:
        min_date = df_processed['date'].min()
        max_date = df_processed['date'].max()

        col_sel, col_info = st.columns([1, 2])
        with col_sel:
            selected_date = st.date_input(
                "Select Date", value=max_date,
                min_value=min_date, max_value=max_date,
                key=f"{_k}rr_daily_select"
            )
        with col_info:
            info_placeholder = st.empty()
            info_base_text = f"**Viewing Date:** {selected_date.strftime('%A, %d %b %Y')}"

        df_view = df_processed[df_processed["date"] == selected_date]
        if df_view.empty:
            st.warning(f"No data available for {selected_date.strftime('%d %b %Y')}.")
        sub_header = f"Summary for {selected_date.strftime('%d %b %Y')}"

    elif "Weekly" in analysis_level:
        available_years = sorted(df_processed['year'].unique())
        col_w_sel, col_w_info = st.columns([1, 2])
        with col_w_sel:
            c_yr, c_wk = st.columns(2)
            with c_yr:
                selected_year = st.selectbox(
                    "Select Year", options=available_years,
                    index=len(available_years) - 1, key=f"{_k}rr_year_week_select"
                )
            weeks_in_year = df_processed[df_processed['year'] == selected_year]['week'].unique()
            sorted_weeks = sorted(weeks_in_year)
            with c_wk:
                selected_week = st.selectbox(
                    "Select Week", options=sorted_weeks,
                    index=len(sorted_weeks) - 1,
                    format_func=lambda w: f"Week {w}",
                    key=f"{_k}rr_week_select"
                )
        try:
            start_of_week = datetime.strptime(
                f'{selected_year}-W{int(selected_week):02d}-1', "%G-W%V-%u"
            )
        except Exception:
            start_of_week = (datetime(selected_year, 1, 1)
                             + timedelta(weeks=int(selected_week)))
        end_of_week = start_of_week + timedelta(days=6)

        with col_w_info:
            info_placeholder = st.empty()
            info_base_text = (
                f"**Viewing Week {selected_week}, {selected_year}**\n\n"
                f"({start_of_week.strftime('%d %b')} – {end_of_week.strftime('%d %b %Y')})"
            )

        df_view = df_processed[
            (df_processed["week"] == selected_week)
            & (df_processed["year"] == selected_year)
        ]
        sub_header = f"Summary for Week {selected_week} ({selected_year})"

    elif "Monthly" in analysis_level:
        df_processed['year_cal'] = df_processed['shot_time'].dt.year
        available_years = sorted(df_processed['year_cal'].unique())
        col_m_sel, col_m_info = st.columns([1, 2])
        with col_m_sel:
            c_yr, c_mo = st.columns(2)
            with c_yr:
                selected_year = st.selectbox(
                    "Select Year", options=available_years,
                    index=len(available_years) - 1, key=f"{_k}rr_year_select"
                )
            months_in_year = df_processed[
                df_processed['year_cal'] == selected_year
            ]['month'].unique()
            sorted_months = sorted(months_in_year)
            with c_mo:
                selected_month_period = st.selectbox(
                    "Select Month", options=sorted_months,
                    index=len(sorted_months) - 1,
                    format_func=lambda p: p.strftime('%B'),
                    key=f"{_k}rr_month_select"
                )
        with col_m_info:
            info_placeholder = st.empty()
            info_base_text = f"**Viewing Month:** {selected_month_period.strftime('%B %Y')}"

        df_view = df_processed[df_processed["month"] == selected_month_period]
        sub_header = f"Summary for {selected_month_period.strftime('%B %Y')}"

    elif "Custom Period" in analysis_level:
        min_date = df_processed['date'].min()
        max_date = df_processed['date'].max()
        col_c_sel, col_c_info = st.columns([1, 2])
        with col_c_sel:
            c1, c2 = st.columns(2)
            with c1:
                start_date = st.date_input("Start date", min_date,
                                           min_value=min_date, max_value=max_date,
                                           key=f"{_k}rr_custom_start")
            with c2:
                end_date = st.date_input("End date", max_date,
                                         min_value=start_date, max_value=max_date,
                                         key=f"{_k}rr_custom_end")
        with col_c_info:
            info_placeholder = st.empty()
            info_base_text = (
                f"**Viewing Period:** {start_date.strftime('%d %b %Y')} "
                f"to {end_date.strftime('%d %b %Y')}"
                if start_date and end_date else "**Viewing Period:** Select dates"
            )

        if start_date and end_date:
            mask = (df_processed['date'] >= start_date) & (df_processed['date'] <= end_date)
            df_view = df_processed[mask]
            sub_header = (f"Summary for {start_date.strftime('%d %b %Y')} "
                          f"to {end_date.strftime('%d %b %Y')}")

    # ------------------------------------------------------------------
    # Run labelling on the slice
    # ------------------------------------------------------------------
    if not df_view.empty:
        df_view = df_view.copy()
        if 'run_id' in df_view.columns:
            # #18 fix: sort by first shot_time of each run before labelling
            # so Run 001 is always the earliest run, matching the Excel export
            run_first_shot = (df_view.groupby('run_id')['shot_time']
                              .min().sort_values())
            run_label_map = {rid: f"Run {i+1:03d}"
                             for i, rid in enumerate(run_first_shot.index)}
            df_view['run_label'] = df_view['run_id'].map(run_label_map)

    # Min-shots filter
    run_count = 0
    if 'by Run' in analysis_level and not df_view.empty:
        run_shot_counts = df_view.groupby('run_label')['run_label'].transform('count')
        df_view = df_view[run_shot_counts >= min_shots_filter].copy()
        # Re-label surviving runs consecutively in chronological order.
        # Without this, run_label has gaps (e.g. Run 001,002,003,005,...) while
        # calculate_run_summaries below renumbers consecutively (Run 001..00M),
        # causing the bucket-trend pivot reindex to drop one row to all zeros.
        if not df_view.empty and 'run_id' in df_view.columns:
            _first_shot = (df_view.groupby('run_id')['shot_time']
                           .min().sort_values())
            _relabel = {rid: f"Run {i+1:03d}"
                        for i, rid in enumerate(_first_shot.index)}
            df_view['run_label'] = df_view['run_id'].map(_relabel)
        run_count = df_view['run_label'].nunique() if not df_view.empty else 0
    elif not df_view.empty and 'run_label' in df_view.columns:
        run_count = df_view['run_label'].nunique()

    if info_placeholder:
        info_placeholder.info(
            f"{info_base_text}\n\n**Number of Production Runs:** {run_count}"
        )

    if df_view.empty:
        st.warning("No data for the selected period (or all runs were filtered out).")
        return

    # ------------------------------------------------------------------
    # KPI computation
    # ------------------------------------------------------------------
    # Run summaries (for MTTR, MTBF, stability, etc.)
    run_summary_df_for_totals = rr_utils.calculate_run_summaries(
        df_view, tolerance, downtime_gap_tolerance, pre_processed=True
    )

    summary_metrics = {}
    if not run_summary_df_for_totals.empty:
        total_runtime_sec = run_summary_df_for_totals['total_runtime_sec'].sum()
        production_time_sec = run_summary_df_for_totals['production_time_sec'].sum()
        downtime_sec = run_summary_df_for_totals['downtime_sec'].sum()
        total_shots = run_summary_df_for_totals['total_shots'].sum()
        normal_shots = run_summary_df_for_totals['normal_shots'].sum()
        stop_events = run_summary_df_for_totals['stops'].sum()

        summary_metrics = {
            'total_runtime_sec': total_runtime_sec,
            'production_time_sec': production_time_sec,
            'downtime_sec': downtime_sec,
            'total_shots': total_shots,
            'normal_shots': normal_shots,
            'stop_events': stop_events,
            'mttr_min': (downtime_sec / 60 / stop_events) if stop_events > 0 else 0,
            'mtbf_min': ((production_time_sec / 60 / stop_events)
                         if stop_events > 0 else (production_time_sec / 60)),
            'stability_index': ((production_time_sec / total_runtime_sec * 100)
                                if total_runtime_sec > 0 else 100.0),
            'efficiency': (normal_shots / total_shots) if total_shots > 0 else 0,
        }
        sub_header = sub_header.replace("Summary for", "Summary for (Combined Runs)")

    # ------------------------------------------------------------------
    # FIX: CT display metrics read from df_view's pre-computed columns.
    #
    # df_view is a slice of df_processed, which was built by the single
    # authoritative RunRateCalculator call in get_processed_data().
    # That call operated on the FULL tool dataset, so mode_ct/limits are
    # computed from the correct run boundaries — NOT from the day slice.
    #
    # Previously a fresh RunRateCalculator(df_view) was called here,
    # which recalculated mode_ct from only the subset of shots in the
    # selected period, causing the 0.1 discrepancy (e.g. 101.6 vs 101.5).
    # ------------------------------------------------------------------
    if 'mode_ct' in df_view.columns:
        summary_metrics['min_mode_ct'] = df_view['mode_ct'].min()
        summary_metrics['max_mode_ct'] = df_view['mode_ct'].max()
    else:
        summary_metrics['min_mode_ct'] = 0
        summary_metrics['max_mode_ct'] = 0

    if 'mode_lower' in df_view.columns:
        summary_metrics['min_lower_limit'] = df_view['mode_lower'].min()
        summary_metrics['max_lower_limit'] = df_view['mode_lower'].max()
    else:
        summary_metrics['min_lower_limit'] = 0
        summary_metrics['max_lower_limit'] = 0

    if 'mode_upper' in df_view.columns:
        summary_metrics['min_upper_limit'] = df_view['mode_upper'].min()
        summary_metrics['max_upper_limit'] = df_view['mode_upper'].max()
    else:
        summary_metrics['min_upper_limit'] = 0
        summary_metrics['max_upper_limit'] = 0

    if 'approved_ct' in df_view.columns:
        valid_app = df_view['approved_ct'].dropna()
        summary_metrics['min_approved_ct'] = valid_app.min() if not valid_app.empty else np.nan
        summary_metrics['max_approved_ct'] = valid_app.max() if not valid_app.empty else np.nan
    else:
        summary_metrics['min_approved_ct'] = np.nan
        summary_metrics['max_approved_ct'] = np.nan

    # We still need a results dict for the bar chart, run_durations, hourly_summary etc.
    # build_display_results uses df_view's pre-computed columns so no mode
    # recomputation occurs on the day/week slice.
    results = rr_utils.build_display_results(df_view, run_interval_hours)

    # ------------------------------------------------------------------
    # Trend / run summary (for charts and tables)
    # ------------------------------------------------------------------
    trend_summary_df = None
    run_summary_df = None
    if "by Run" in analysis_level:
        trend_summary_df = rr_utils.calculate_run_summaries(
            df_view, tolerance, downtime_gap_tolerance, pre_processed=True
        )
        if trend_summary_df is not None and not trend_summary_df.empty:
            trend_summary_df.rename(columns={
                'run_label': 'RUN ID', 'stability_index': 'STABILITY %',
                'stops': 'STOPS', 'mttr_min': 'MTTR (min)',
                'mtbf_min': 'MTBF (min)', 'total_shots': 'Total Shots',
                'approved_ct': 'Approved CT'
            }, inplace=True)
        run_summary_df = trend_summary_df

    # ------------------------------------------------------------------
    # Header + export button
    # ------------------------------------------------------------------
    col1, col2 = st.columns([3, 1])
    with col1:
        st.subheader(sub_header)
    with col2:
        st.download_button(
            label="📥 Export Run-Based Report",
            data=rr_utils.prepare_and_generate_run_based_excel(
                df_view.copy(), tolerance, downtime_gap_tolerance,
                run_interval_hours, tool_id_selection
            ),
            file_name=(
                f"Run_Based_Report_"
                f"{tool_id_selection.replace(' / ', '_').replace(' ', '_')}_"
                f"{analysis_level.replace(' ', '_')}_"
                f"{datetime.now():%Y%m%d}.xlsx"
            ),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width='stretch'
        )

    # ------------------------------------------------------------------
    # 1. KPI metrics
    # ------------------------------------------------------------------
    with st.container(border=True):
        col1, col2, col3, col4, col5 = st.columns(5)
        total_d = summary_metrics.get('total_runtime_sec', 0)
        prod_t = summary_metrics.get('production_time_sec', 0)
        down_t = summary_metrics.get('downtime_sec', 0)
        prod_p = (prod_t / total_d * 100) if total_d > 0 else 0
        down_p = (down_t / total_d * 100) if total_d > 0 else 0
        mttr_display = f"{summary_metrics.get('mttr_min', 0):.1f} min"
        mtbf_display = f"{summary_metrics.get('mtbf_min', 0):.1f} min"

        with col1:
            st.metric("Run Rate MTTR", mttr_display,
                      help="Mean Time To Repair: average duration of a stop event.\n\n"
                           "Formula: Total Downtime / Stop Events")
        with col2:
            st.metric("Run Rate MTBF", mtbf_display,
                      help="Mean Time Between Failures: average duration of stable "
                           "operation between stop events.\n\n"
                           "Formula: Total Production Time / Stop Events")
        with col3:
            st.metric("Total Run Duration", rr_utils.format_duration(total_d),
                      help="Sum of all individual production run durations. "
                           "Gaps between runs (> Run Interval Threshold) are excluded.\n\n"
                           "Formula (per run): (Last Shot Time − First Shot Time) + Last Shot CT")
        with col4:
            st.metric("Production Time", rr_utils.format_duration(prod_t),
                      help="Sum of Actual CT for all normal (non-stop) shots.")
            st.markdown(
                f'<span style="background-color:{rr_utils.PASTEL_COLORS["green"]};'
                f'color:#0E1117;padding:3px 8px;border-radius:10px;'
                f'font-size:0.8rem;font-weight:bold;">{prod_p:.1f}%</span>',
                unsafe_allow_html=True
            )
        with col5:
            st.metric("Run Rate Downtime", rr_utils.format_duration(down_t),
                      help="Total Run Duration − Total Production Time.")
            st.markdown(
                f'<span style="background-color:{rr_utils.PASTEL_COLORS["red"]};'
                f'color:#0E1117;padding:3px 8px;border-radius:10px;'
                f'font-size:0.8rem;font-weight:bold;">{down_p:.1f}%</span>',
                unsafe_allow_html=True
            )

    with st.container(border=True):
        c1, c2 = st.columns(2)
        c1.plotly_chart(
            rr_utils.create_gauge(
                summary_metrics.get('efficiency', 0) * 100, "Run Rate Shot Efficiency (%)"
            ),
            width='stretch'
        )
        steps = [
            {'range': [0, 50], 'color': rr_utils.PASTEL_COLORS['red']},
            {'range': [50, 70], 'color': rr_utils.PASTEL_COLORS['orange']},
            {'range': [70, 100], 'color': rr_utils.PASTEL_COLORS['green']},
        ]
        c2.plotly_chart(
            rr_utils.create_gauge(
                summary_metrics.get('stability_index', 0),
                "Run Rate Time Stability (%)", steps=steps
            ),
            width='stretch'
        )

    with st.expander("ℹ️ What do these metrics mean?"):
        st.markdown("""
        **Run Rate Shot Efficiency (%)**
        > Percentage of shots that were 'Normal' (stop_flag = 0).
        > - *Formula: Normal Shots / Total Shots*

        **Run Rate Time Stability (%)**
        > Percentage of total run time spent in normal production.
        > - *Formula: Total Production Time / Total Run Duration*

        **Run Rate MTTR (min)**
        > Average duration of a single stop event.
        > - *Formula: Total Downtime / Stop Events*

        **Run Rate MTBF (min)**
        > Average duration of stable operation between stop events.
        > - *Formula: Total Production Time / Stop Events*
        """)

    with st.container(border=True):
        c1, c2, c3 = st.columns(3)
        t_s = summary_metrics.get('total_shots', 0)
        n_s = summary_metrics.get('normal_shots', 0)
        s_s = t_s - n_s
        n_p = (n_s / t_s * 100) if t_s > 0 else 0
        s_p = (s_s / t_s * 100) if t_s > 0 else 0
        with c1:
            st.metric("Total Shots", f"{t_s:,}")
        with c2:
            st.metric("Normal Shots", f"{n_s:,}")
            st.markdown(
                f'<span style="background-color:{rr_utils.PASTEL_COLORS["green"]};'
                f'color:#0E1117;padding:3px 8px;border-radius:10px;'
                f'font-size:0.8rem;font-weight:bold;">{n_p:.1f}% of Total</span>',
                unsafe_allow_html=True
            )
        with c3:
            st.metric("Stop Events", f"{summary_metrics.get('stop_events', 0)}")
            st.markdown(
                f'<span style="background-color:{rr_utils.PASTEL_COLORS["red"]};'
                f'color:#0E1117;padding:3px 8px;border-radius:10px;'
                f'font-size:0.8rem;font-weight:bold;">{s_p:.1f}% Stopped Shots</span>',
                unsafe_allow_html=True
            )

    # ------------------------------------------------------------------
    # 2. Cycle time / mode display (unit-aware: CT, SPM or SPH)
    # ------------------------------------------------------------------
    def fmt_metric(min_val, max_val, to_spm=False):
        if pd.isna(min_val) or pd.isna(max_val):
            return "N/A"
        if to_spm:
            # Limits invert under reciprocal transform: high CT → low SPM
            _lo = rr_utils.ct_to_stroke_rate(max_val, stroke_unit)
            _hi = rr_utils.ct_to_stroke_rate(min_val, stroke_unit)
            min_val, max_val = float(np.atleast_1d(_lo)[0]), float(np.atleast_1d(_hi)[0])
            if pd.isna(min_val) or pd.isna(max_val):
                return "N/A"
        if abs(min_val - max_val) < 0.005:
            return f"{min_val:.2f}"
        return f"{min_val:.2f} – {max_val:.2f}"

    _convert = press_mode and stroke_unit != "CT"

    if stroke_unit == "CT" or not press_mode:
        _lbl_mode  = "Mode Cycle Time (sec)"
        _lbl_c1    = "Lower Limit (sec)"
        _lbl_c3    = "Upper Limit (sec)"
        _lbl_app   = "Approved CT (sec)"
    else:
        _lbl_mode  = f"Mode {stroke_unit}"
        _lbl_c1    = f"Upper {stroke_unit} Limit"   # limits invert under reciprocal
        _lbl_c3    = f"Lower {stroke_unit} Limit"
        _lbl_app   = f"Approved {stroke_unit}"

    _mode_lo = summary_metrics.get('min_mode_ct', 0)
    _mode_hi = summary_metrics.get('max_mode_ct', 0)
    _low_lo  = summary_metrics.get('min_lower_limit', 0)
    _low_hi  = summary_metrics.get('max_lower_limit', 0)
    _up_lo   = summary_metrics.get('min_upper_limit', 0)
    _up_hi   = summary_metrics.get('max_upper_limit', 0)
    _app_lo  = summary_metrics.get('min_approved_ct', np.nan)
    _app_hi  = summary_metrics.get('max_approved_ct', np.nan)

    _fmt_mode  = fmt_metric(_mode_lo, _mode_hi, to_spm=_convert)
    _fmt_lower = fmt_metric(_low_lo,  _low_hi,  to_spm=_convert)
    _fmt_upper = fmt_metric(_up_lo,   _up_hi,   to_spm=_convert)
    _fmt_app   = fmt_metric(_app_lo,  _app_hi,  to_spm=_convert)

    if show_approved_ct:
        c_main, c_app = st.columns([3, 1])
        with c_main:
            with st.container(border=True):
                c1, c2, c3 = st.columns(3)
                c1.metric(_lbl_c1, _fmt_lower)
                with c2:
                    with st.container(border=True):
                        st.metric(_lbl_mode, _fmt_mode)
                c3.metric(_lbl_c3, _fmt_upper)
        with c_app:
            with st.container(border=True):
                st.metric(_lbl_app, _fmt_app)
    else:
        with st.container(border=True):
            c1, c2, c3 = st.columns(3)
            c1.metric(_lbl_c1, _fmt_lower)
            with c2:
                with st.container(border=True):
                    st.metric(_lbl_mode, _fmt_mode)
            c3.metric(_lbl_c3, _fmt_upper)

    # Always show raw CT reference when displaying in rate units
    if _convert:
        _raw_mode  = fmt_metric(_mode_lo, _mode_hi, to_spm=False)
        _raw_lower = fmt_metric(_low_lo,  _low_hi,  to_spm=False)
        _raw_upper = fmt_metric(_up_lo,   _up_hi,   to_spm=False)
        st.caption(
            f"ℹ️ Calculation reference (CT seconds) — "
            f"Lower: **{_raw_lower}s** · Mode: **{_raw_mode}s** · Upper: **{_raw_upper}s**"
        )

    # ------------------------------------------------------------------
    # Automated analysis expander
    # ------------------------------------------------------------------
    if detailed_view:
        st.markdown("---")
        with st.expander("🤖 View Automated Analysis Summary", expanded=False):
            analysis_df = pd.DataFrame()
            if trend_summary_df is not None and not trend_summary_df.empty:
                analysis_df = trend_summary_df.copy()
                rename_map = {
                    'RUN ID': 'period', 'STABILITY %': 'stability',
                    'STOPS': 'stops', 'MTTR (min)': 'mttr'
                }
                analysis_df.rename(columns=rename_map, inplace=True)

            # #10 fix: validate required columns exist before calling analysis
            _required = {'period', 'stability', 'stops', 'mttr'}
            if not analysis_df.empty and not _required.issubset(analysis_df.columns):
                analysis_df = pd.DataFrame()  # fall through to "not enough data"

            insights = rr_utils.generate_detailed_analysis(
                analysis_df,
                summary_metrics.get('stability_index', 0),
                summary_metrics.get('mttr_min', 0),
                summary_metrics.get('mtbf_min', 0),
                analysis_level
            )

            if "error" in insights:
                st.error(insights["error"])
            else:
                components.html(
                    f"""<div style="border:1px solid #333;border-radius:0.5rem;padding:1.5rem;
                    margin-top:1rem;font-family:sans-serif;line-height:1.6;
                    background-color:#0E1117;">
                    <h4 style="margin-top:0;color:#FAFAFA;">Automated Analysis Summary</h4>
                    <p style="color:#FAFAFA;"><strong>Overall Assessment:</strong> {insights['overall']}</p>
                    <p style="color:#FAFAFA;"><strong>Predictive Trend:</strong> {insights['predictive']}</p>
                    <p style="color:#FAFAFA;"><strong>Performance Variance:</strong> {insights['best_worst']}</p>
                    {'<p style="color:#FAFAFA;"><strong>Identified Patterns:</strong> ' + insights['patterns'] + '</p>' if insights['patterns'] else ''}
                    <p style="margin-top:1rem;color:#FAFAFA;background-color:#262730;
                    padding:1rem;border-radius:0.5rem;">
                    <strong>Key Recommendation:</strong> {insights['recommendation']}</p>
                    </div>""",
                    height=400, scrolling=True
                )

    st.markdown("---")

    # ------------------------------------------------------------------
    # 3. Shot / stroke charts
    # ------------------------------------------------------------------
    time_agg = ('hourly' if "Daily" in analysis_level
                else 'daily' if 'Weekly' in analysis_level
                else 'weekly')

    # stroke_unit for rate charts — CT mode falls back to SPM for bucketed chart
    _chart_su = stroke_unit if stroke_unit != "CT" else "SPM"

    if press_mode:
        if stroke_unit != "CT":
            with st.expander("ℹ️ How to read this chart — Bucketed Stroke Rate", expanded=False):
                st.markdown(f"""
                **What it shows:** Actual stroke counts aggregated into
                {'1-minute' if _chart_su == 'SPM' else '1-hour'} time buckets.
                The count in each bucket *is* the {_chart_su} for that period.

                **Stacked bars:** Blue = normal strokes, Red = stopped strokes.
                **Mode line** (dotted grey) = proven rhythm from mode CT.
                **Run boundaries** (purple dashed) show where distinct runs start.
                """)
            rr_utils.plot_stroke_rate_chart(
                results['processed_df'], results.get('mode_ct'),
                stroke_unit=_chart_su, show_approved_ct=show_approved_ct
            )
            st.markdown("---")

        with st.expander("ℹ️ How to read this chart — Raw Cycle Time per Shot", expanded=False):
            st.markdown("""
            **What it shows:** Every individual stroke plotted at its actual cycle time
            in seconds. This is the raw signal all metrics are derived from.

            **Green band** = tolerance window (mode CT ± tolerance %).
            Blue = normal strokes, Red = stopped/out-of-tolerance strokes.
            **Hard-stop shots** (CT = 999.9s) appear as very tall red bars.
            """)
        rr_utils.plot_shot_bar_chart(
            results['processed_df'],
            results.get('mode_lower'), results.get('mode_upper'),
            results.get('mode_ct'), time_agg=time_agg,
            show_approved_ct=show_approved_ct, press_mode=False, stroke_unit=_chart_su
        )
        st.markdown("---")
    else:
        rr_utils.plot_shot_bar_chart(
            results['processed_df'],
            results.get('mode_lower'),
            results.get('mode_upper'),
            results.get('mode_ct'),
            time_agg=time_agg,
            show_approved_ct=show_approved_ct,
            press_mode=False,
            stroke_unit=stroke_unit
        )

    # CT Histogram — all tool types, collapsed by default
    # CT Histogram — commented out pending further development
    # with st.expander("📊 Cycle Time Distribution", expanded=False):
    #     rr_utils.plot_ct_histogram(results['processed_df'])

    with st.expander("View Shot Data Table", expanded=False):
        _df_src = results['processed_df']
        _pm = press_mode

        # Build column list — hierarchy first, then shot detail
        _shot_cols  = []
        _shot_names = {}

        # Hierarchy columns (if present in data)
        for _col, _lbl in [
            ('tool_id',       'Tooling ID'),
            ('supplier_id', 'Supplier'),
            ('tooling_type',  'Tooling Type'),
            ('part_id',       'Part(s)'),
        ]:
            if _col in _df_src.columns:
                _shot_cols.append(_col)
                _shot_names[_col] = _lbl

        # Run ID
        if 'run_label' in _df_src.columns:
            _shot_cols.append('run_label')
            _shot_names['run_label'] = 'Run ID'

        # Core shot columns
        _shot_cols  += ['shot_time', 'mode_ct', 'actual_ct', 'adj_ct_sec']
        _shot_names.update({
            'shot_time':  'Date / Time',
            'mode_ct':    'Mode CT (sec)',
            'actual_ct':  'Actual CT (sec)',
            'adj_ct_sec': 'Adjusted CT (sec)',
        })

        if show_approved_ct and 'approved_ct' in _df_src.columns:
            _shot_cols.append('approved_ct')
            _shot_names['approved_ct'] = 'Approved CT (sec)'

        _shot_cols += ['time_diff_sec', 'stop_flag', 'stop_event']
        _shot_names.update({
            'time_diff_sec': 'Time Difference (sec)',
            'stop_flag':     'Stop Flag',
            'stop_event':    'Stop Event',
        })

        _existing = [c for c in _shot_cols if c in _df_src.columns]
        df_shot_data = _df_src[_existing].copy()

        df_shot_data.rename(columns=_shot_names, inplace=True)

        # Format datetime to milliseconds only (3dp), not microseconds
        if 'Date / Time' in df_shot_data.columns:
            df_shot_data['Date / Time'] = pd.to_datetime(
                df_shot_data['Date / Time']
            ).dt.strftime('%Y-%m-%d %H:%M:%S.%f').str[:-3]

        # UI display: format to 2dp consistently
        _fmt_shot = {c: '{:.2f}' for c in ['Actual CT (sec)', 'Adjusted CT (sec)',
                                             'Approved CT (sec)', 'Mode CT (sec)',
                                             'Time Difference (sec)']
                     if c in df_shot_data.columns}

        st.dataframe(df_shot_data.style.format(_fmt_shot, na_rep='—'), use_container_width=True)

        # CSV download — full precision, no formatting
        st.download_button(
            "📥 Download Shot Data (CSV)",
            data=df_shot_data.to_csv(index=False),
            file_name=f"shot_data_{tool_id_selection.replace(' ', '_')}.csv",
            mime="text/csv",
            key=f"{_k}rr_shot_csv"
        )

    st.markdown("---")

    # ------------------------------------------------------------------
    # 4. Detailed analysis section
    # ------------------------------------------------------------------
    analysis_view_mode = "Run"
    if analysis_level == "Daily (by Run)":
        c_head, c_view = st.columns([3, 1])
        with c_head:
            st.header("Detailed Analysis")
        with c_view:
            analysis_view_mode = st.selectbox("Group By", ["Run", "Hour"],
                                              key=f"{_k}rr_view_mode")
    else:
        st.header("Run-Based Analysis")

    # Shared data for charts
    run_durations = results.get("run_durations", pd.DataFrame())
    processed_df = results.get('processed_df', pd.DataFrame())
    stop_events_df = processed_df.loc[processed_df['stop_event']].copy()
    complete_runs = pd.DataFrame()
    if not stop_events_df.empty:
        stop_events_df['terminated_run_group'] = stop_events_df['run_group'] - 1
        # Keep only the first stop event per run_group to guarantee a unique
        # index — pandas 2.x raises InvalidIndexError on duplicate index in .map()
        end_time_map = (stop_events_df
                        .drop_duplicates(subset='terminated_run_group', keep='first')
                        .set_index('terminated_run_group')['shot_time'])
        run_durations['run_end_time'] = run_durations['run_group'].map(end_time_map)
        complete_runs = run_durations.dropna(subset=['run_end_time']).copy()

    # ------------------------------------------------------------------
    # Option A: Run-based view
    # ------------------------------------------------------------------
    if analysis_view_mode == "Run":

        with st.expander("View Run Breakdown Table", expanded=True):
            if run_summary_df is not None and not run_summary_df.empty:
                d_df = run_summary_df.copy()
                _pm = press_mode
                _shot_lbl   = 'Strokes' if _pm else 'Shots'

                d_df["Period (date/time from to)"] = d_df.apply(
                    lambda r: (f"{r['start_time'].strftime('%Y-%m-%d %H:%M')} to "
                               f"{r['end_time'].strftime('%Y-%m-%d %H:%M')}"),
                    axis=1
                )

                total_shots_col = ('Total Shots' if 'Total Shots' in d_df.columns
                                   else 'total_shots')
                # Store raw numeric series before overwriting with formatted string
                _raw_total = d_df[total_shots_col]
                d_df[f"Total {_shot_lbl}"] = _raw_total.apply(lambda x: f"{x:,}")

                d_df[f"Normal {_shot_lbl}"] = d_df.apply(
                    lambda r: (
                        f"{r['normal_shots']:,} "
                        f"({r['normal_shots'] / _raw_total.loc[r.name] * 100:.1f}%)"
                        if _raw_total.loc[r.name] > 0 else "0 (0.0%)"
                    ), axis=1
                )

                if 'stopped_shots' not in d_df.columns:
                    d_df['stopped_shots'] = _raw_total - d_df['normal_shots']

                stops_col = 'STOPS' if 'STOPS' in d_df.columns else 'stops'
                d_df["Stop Events"] = d_df.apply(
                    lambda r: (
                        f"{r[stops_col]} "
                        f"({r['stopped_shots'] / _raw_total.loc[r.name] * 100:.1f}%)"
                        if _raw_total.loc[r.name] > 0 else "0 (0.0%)"
                    ), axis=1
                )

                d_df["Total Run duration (d/h/m)"] = d_df['total_runtime_sec'].apply(
                    rr_utils.format_duration
                )
                d_df["Production Time (d/h/m)"] = d_df.apply(
                    lambda r: (
                        f"{rr_utils.format_duration(r['production_time_sec'])} "
                        f"({r['production_time_sec'] / r['total_runtime_sec'] * 100:.1f}%)"
                        if r['total_runtime_sec'] > 0 else "0m (0.0%)"
                    ), axis=1
                )
                d_df["Downtime (d/h/m)"] = d_df.apply(
                    lambda r: (
                        f"{rr_utils.format_duration(r['downtime_sec'])} "
                        f"({r['downtime_sec'] / r['total_runtime_sec'] * 100:.1f}%)"
                        if r['total_runtime_sec'] > 0 else "0m (0.0%)"
                    ), axis=1
                )

                col_rename = {
                    'run_label': 'RUN ID', 'mode_ct': 'Mode CT (sec)',
                    'mode_lower': 'Lower Limit (sec)',
                    'mode_upper': 'Upper Limit (sec)',
                    'mttr_min': 'MTTR (min)', 'mtbf_min': 'MTBF (min)',
                    'stability_index': 'Run Rate Time Stability (%)', 'stops': 'STOPS',
                    'MTTR (min)': 'MTTR (min)', 'MTBF (min)': 'MTBF (min)',
                    'STABILITY %': 'Run Rate Time Stability (%)', 'STOPS': 'STOPS',
                }
                approved_key = ('Approved CT' if 'Approved CT' in d_df.columns
                                else 'approved_ct')
                col_rename[approved_key] = 'Approved CT (sec)'
                d_df.rename(columns=col_rename, inplace=True)

                # Add Mode SPM / SPH derived from Mode CT — only meaningful at run level
                if 'Mode CT (sec)' in d_df.columns:
                    _mct = pd.to_numeric(d_df['Mode CT (sec)'], errors='coerce')
                    d_df['Mode SPM'] = (60.0   / _mct).round(2)
                    d_df['Mode SPH'] = (3600.0 / _mct).round(2)

                final_cols = [
                    'RUN ID', 'Period (date/time from to)',
                    f'Total {_shot_lbl}', f'Normal {_shot_lbl}', 'Stop Events',
                    'Mode CT (sec)', 'Mode SPM', 'Mode SPH',
                    'Approved CT (sec)',
                    'Lower Limit (sec)', 'Upper Limit (sec)',
                    'Total Run duration (d/h/m)', 'Production Time (d/h/m)',
                    'Downtime (d/h/m)', 'MTTR (min)', 'MTBF (min)', 'Run Rate Time Stability (%)'
                ]
                if not show_approved_ct and 'Approved CT (sec)' in final_cols:
                    final_cols.remove('Approved CT (sec)')
                final_cols = [c for c in final_cols if c in d_df.columns]

                # Item 4: consistent 2dp for all numeric columns
                fmt = {c: '{:.2f}' for c in [
                    'Mode CT (sec)', 'Mode SPM', 'Mode SPH',
                    'Approved CT (sec)',
                    'Lower Limit (sec)', 'Upper Limit (sec)',
                    'MTTR (min)', 'MTBF (min)', 'Run Rate Time Stability (%)'
                ] if c in d_df.columns}

                st.dataframe(d_df[final_cols].style.format(fmt, na_rep='—'),
                             use_container_width=True)

        c1, c2 = st.columns(2)
        with c1:
            st.subheader("Total Bucket Analysis")
            if not complete_runs.empty and "time_bucket" in complete_runs.columns:
                b_counts = (complete_runs["time_bucket"]
                            .value_counts()
                            .reindex(results["bucket_labels"], fill_value=0))
                fig_b = (px.bar(b_counts,
                                title="Total Time Bucket Analysis",
                                labels={"index": "Duration (min)", "value": "Occurrences"},
                                text_auto=True,
                                color=b_counts.index,
                                color_discrete_map=results["bucket_color_map"])
                         .update_layout(legend_title_text='Duration'))
                fig_b.update_xaxes(title_text="Duration (min)")
                fig_b.update_yaxes(title_text="Occurrences")
                st.plotly_chart(fig_b, width='stretch')
                with st.expander("View Bucket Data Table", expanded=False):
                    cols_bucket = ['run_group', 'duration_min', 'time_bucket',
                                   'run_end_time', 'run_label']
                    df_bucket_data = complete_runs[
                        [c for c in cols_bucket if c in complete_runs.columns]
                    ].rename(columns={
                        'run_group': 'Run Group', 'duration_min': 'Duration (min)',
                        'time_bucket': 'Time Bucket', 'run_end_time': 'Run End Date/Time',
                        'run_label': 'Run ID'
                    })
                    st.dataframe(df_bucket_data)
            else:
                st.info("No complete runs.")

        with c2:
            st.subheader("Stability per Production Run")
            if run_summary_df is not None and not run_summary_df.empty:
                rr_utils.plot_trend_chart(
                    run_summary_df, 'RUN ID', 'STABILITY %',
                    "Stability per Run", "Run ID", "Run Rate Time Stability (%)", is_stability=True
                )
                with st.expander("View Stability Data Table", expanded=False):
                    df_renamed = rr_utils.get_renamed_summary_df(run_summary_df)
                    if not show_approved_ct and 'Approved CT' in df_renamed.columns:
                        df_renamed = df_renamed.drop(columns=['Approved CT'])
                    st.dataframe(df_renamed)
            else:
                st.info("No runs to analyse.")

        st.subheader("Bucket Trend per Production Run")
        if (not complete_runs.empty
                and run_summary_df is not None
                and not run_summary_df.empty):
            run_group_to_label_map = (processed_df.drop_duplicates('run_group')
                                      [['run_group', 'run_label']]
                                      .set_index('run_group')['run_label'])
            complete_runs['run_label'] = complete_runs['run_group'].map(run_group_to_label_map)
            pivot_df = pd.crosstab(
                index=complete_runs['run_label'],
                columns=complete_runs['time_bucket'].astype('category')
                        .cat.set_categories(results["bucket_labels"])
            )
            pivot_df = pivot_df.reindex(run_summary_df['RUN ID'], fill_value=0)

            fig_bucket_trend = make_subplots(specs=[[{"secondary_y": True}]])
            for col in pivot_df.columns:
                fig_bucket_trend.add_trace(
                    go.Bar(name=col, x=pivot_df.index, y=pivot_df[col],
                           marker_color=results["bucket_color_map"].get(col)),
                    secondary_y=False
                )
            fig_bucket_trend.add_trace(
                go.Scatter(
                    name='Total Shots', x=run_summary_df['RUN ID'],
                    y=run_summary_df['Total Shots'], mode='lines+markers+text',
                    text=run_summary_df['Total Shots'], textposition='top center',
                    line=dict(color='blue')
                ),
                secondary_y=True
            )
            fig_bucket_trend.update_layout(
                barmode='stack',
                title_text='Distribution of Run Durations per Run vs. Shot Count',
                xaxis_title='Run ID', yaxis_title='Number of Runs',
                yaxis2_title='Total Shots', legend_title_text='Run Duration (min)'
            )
            st.plotly_chart(fig_bucket_trend, width='stretch')
            with st.expander("View Bucket Trend Data Table & Analysis", expanded=False):
                st.dataframe(pivot_df)
                if detailed_view:
                    st.markdown(
                        rr_utils.generate_bucket_analysis(
                            complete_runs, results["bucket_labels"]
                        ),
                        unsafe_allow_html=True
                    )

        st.subheader("MTTR & MTBF per Production Run")
        if (run_summary_df is not None
                and not run_summary_df.empty
                and run_summary_df['STOPS'].sum() > 0):
            rr_utils.plot_mttr_mtbf_chart(
                df=run_summary_df, x_col='RUN ID',
                mttr_col='MTTR (min)', mtbf_col='MTBF (min)',
                shots_col='Total Shots',
                title="MTTR, MTBF & Shot Count per Run"
            )
            with st.expander("View MTTR/MTBF Data Table & Correlation Analysis",
                             expanded=False):
                df_renamed = rr_utils.get_renamed_summary_df(run_summary_df)
                if not show_approved_ct and 'Approved CT' in df_renamed.columns:
                    df_renamed = df_renamed.drop(columns=['Approved CT'])
                st.dataframe(df_renamed)
                if detailed_view:
                    analysis_df = pd.DataFrame()
                    if trend_summary_df is not None and not trend_summary_df.empty:
                        analysis_df = trend_summary_df.copy()
                        rm = {}
                        if 'RUN ID' in analysis_df.columns:
                            rm = {'RUN ID': 'period', 'STABILITY %': 'stability',
                                  'STOPS': 'stops', 'MTTR (min)': 'mttr'}
                        analysis_df.rename(columns=rm, inplace=True)
                    st.markdown(
                        rr_utils.generate_mttr_mtbf_analysis(analysis_df, analysis_level),
                        unsafe_allow_html=True
                    )

    # ------------------------------------------------------------------
    # Option B: Hourly view (Daily only)
    # ------------------------------------------------------------------
    elif analysis_view_mode == "Hour":
        hourly_summary_df = results.get('hourly_summary', pd.DataFrame())

        with st.expander("View Hourly Breakdown Table", expanded=True):
            if not hourly_summary_df.empty:
                df_renamed = rr_utils.get_renamed_summary_df(hourly_summary_df)
                if not show_approved_ct and 'Approved CT' in df_renamed.columns:
                    df_renamed = df_renamed.drop(columns=['Approved CT'])
                st.dataframe(df_renamed, width='stretch')
            else:
                st.info("No hourly data available.")

        c1, c2 = st.columns(2)
        with c1:
            st.subheader("Hourly Bucket Trend")
            if not complete_runs.empty:
                complete_runs['hour'] = complete_runs['run_end_time'].dt.hour
                pivot_df = pd.crosstab(
                    index=complete_runs['hour'],
                    columns=complete_runs['time_bucket'].astype('category')
                            .cat.set_categories(results["bucket_labels"])
                )
                pivot_df = pivot_df.reindex(pd.Index(range(24), name='hour'), fill_value=0)
                fig_hourly_bucket = px.bar(
                    pivot_df, x=pivot_df.index, y=pivot_df.columns,
                    title='Hourly Distribution of Run Durations', barmode='stack',
                    color_discrete_map=results["bucket_color_map"],
                    labels={'hour': 'Hour', 'value': 'Number of Buckets',
                            'variable': 'Run Duration (min)'}
                )
                st.plotly_chart(fig_hourly_bucket, width='stretch')
                with st.expander("View Bucket Trend Data", expanded=False):
                    st.dataframe(pivot_df)
            else:
                st.info("No completed runs to chart by hour.")

        with c2:
            st.subheader("Hourly Stability Trend")
            if not hourly_summary_df.empty:
                rr_utils.plot_trend_chart(
                    hourly_summary_df, 'hour', 'stability_index',
                    "Hourly Stability Trend", "Hour of Day", "Run Rate Time Stability (%)",
                    is_stability=True
                )
                with st.expander("View Stability Data", expanded=False):
                    df_renamed = rr_utils.get_renamed_summary_df(hourly_summary_df)
                    if not show_approved_ct and 'Approved CT' in df_renamed.columns:
                        df_renamed = df_renamed.drop(columns=['Approved CT'])
                    st.dataframe(df_renamed)
            else:
                st.info("No hourly stability data.")

        st.subheader("Hourly MTTR & MTBF Trend")
        if not hourly_summary_df.empty and hourly_summary_df['stops'].sum() > 0:
            rr_utils.plot_mttr_mtbf_chart(
                df=hourly_summary_df, x_col='hour',
                mttr_col='mttr_min', mtbf_col='mtbf_min', shots_col='total_shots',
                title="Hourly MTTR & MTBF Trend"
            )
            with st.expander("View MTTR/MTBF Data", expanded=False):
                df_renamed = rr_utils.get_renamed_summary_df(hourly_summary_df)
                if not show_approved_ct and 'Approved CT' in df_renamed.columns:
                    df_renamed = df_renamed.drop(columns=['Approved CT'])
                st.dataframe(df_renamed)
            if detailed_view:
                with st.expander("🤖 View MTTR/MTBF Correlation Analysis", expanded=False):
                    st.info("Automated correlation analysis is best viewed in 'Run' mode.")
        else:
            st.info("No hourly stop data for MTTR/MTBF charts.")


# ==============================================================================
# --- 4. MAIN APP ENTRY POINT ---
# ==============================================================================

APP_VERSION = "v3.53"

def run_run_rate_ui():

    # Version badge — always visible at top of sidebar
    st.sidebar.markdown(
        f"<div style='text-align:left;padding:4px 0 10px 0;margin:0;"
        f"font-size:0.78rem;color:var(--text-color);opacity:0.55;"
        f"display:block;width:100%;'>"
        f"Run Rate Analysis &nbsp;|&nbsp; <strong>{APP_VERSION}</strong></div>",
        unsafe_allow_html=True
    )

    st.sidebar.title("File Upload")
    uploaded_files = st.sidebar.file_uploader(
        "Upload one or more Run Rate files (Excel / CSV)",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        key="rr_file_uploader"
    )

    if not uploaded_files:
        st.info("👈 Upload one or more production data files to begin.")
        st.stop()

    df_all = rr_utils.load_all_data(uploaded_files, _cache_version=APP_VERSION)

    id_col = "tool_id"
    if id_col not in df_all.columns:
        st.error(
            "None of the uploaded files contain an 'EQUIPMENT_CODE', "
            "'TOOLING ID' or 'EQUIPMENT CODE' column."
        )
        st.stop()

    df_all.dropna(subset=[id_col], inplace=True)
    df_all[id_col] = df_all[id_col].astype(str)

    # ------------------------------------------------------------------
    # Sidebar: Global Filters (CR-aligned)
    # Order: Date Range → Project → Material → Part → Supplier → Plant → Tooling Type
    # Uses get_options_multi / apply_filter pattern matching cr_CG_utils.
    # Unknown values sort to bottom; empty selection = show all.
    # ------------------------------------------------------------------
    def get_options_multi(df, col):
        if col not in df.columns:
            return []
        # Strip whitespace before deduplication so "Supplier " and "Supplier"
        # don't appear as two separate options when multiple files are uploaded.
        raw = sorted(set(
            str(x).strip() for x in df[col].unique()
            if str(x).strip().lower() not in ["nan", "none", "", "nat"]
        ))
        known   = [x for x in raw if x.lower() != "unknown"]
        unknown = [x for x in raw if x.lower() == "unknown"]
        return known + unknown

    def apply_filter(df, col, sel):
        if not sel or col not in df.columns:
            return df
        sel_lower = [str(s).lower() for s in sel]
        keep_unknown = "unknown" in sel_lower
        mask = df[col].astype(str).isin(sel)
        if keep_unknown:
            mask = mask | (df[col].astype(str).str.lower() == "unknown")
        return df[mask]

    st.sidebar.markdown("### Global Filters")

    # 0. Date Range
    _data_min = df_all['shot_time'].min().date() if not df_all.empty else datetime.now().date()
    _data_max = df_all['shot_time'].max().date() if not df_all.empty else datetime.now().date()
    _range = st.sidebar.date_input(
        "Date Range", value=[_data_min, _data_max],
        min_value=_data_min, max_value=_data_max,
        key="rr_global_date_range"
    )
    if isinstance(_range, (list, tuple)) and len(_range) == 2:
        start_d, end_d = _range
        df_all = df_all[
            (df_all['shot_time'].dt.date >= start_d) &
            (df_all['shot_time'].dt.date <= end_d)
        ]
    else:
        st.sidebar.warning("Please select both a start and end date.")
        st.stop()

    if df_all.empty:
        st.sidebar.warning("No data for the selected date range.")
        st.stop()

    # 1–6. Cascading hierarchy filters
    opts_proj = get_options_multi(df_all, 'project_id')
    sel_proj  = st.sidebar.multiselect("Project",      opts_proj, default=opts_proj, key="rr_f_project")
    df_f1     = apply_filter(df_all, 'project_id', sel_proj)

    opts_mat  = get_options_multi(df_f1, 'material')
    sel_mat   = st.sidebar.multiselect("Material",     opts_mat,  default=opts_mat,  key="rr_f_material")
    df_f2     = apply_filter(df_f1, 'material', sel_mat)

    opts_part = get_options_multi(df_f2, 'part_id')
    sel_part  = st.sidebar.multiselect("Part",         opts_part, default=opts_part, key="rr_f_part")
    df_f3     = apply_filter(df_f2, 'part_id', sel_part)

    opts_sup  = get_options_multi(df_f3, 'supplier_id')
    sel_sup   = st.sidebar.multiselect("Supplier",     opts_sup,  default=opts_sup,  key="rr_f_supplier")
    df_f4     = apply_filter(df_f3, 'supplier_id', sel_sup)

    opts_plt  = get_options_multi(df_f4, 'plant_id')
    sel_plt   = st.sidebar.multiselect("Plant",        opts_plt,  default=opts_plt,  key="rr_f_plant")
    df_f5     = apply_filter(df_f4, 'plant_id', sel_plt)

    opts_tt   = get_options_multi(df_f5, 'tooling_type')
    sel_tt    = st.sidebar.multiselect("Tooling Type", opts_tt,   default=opts_tt,   key="rr_f_tooling_type")
    df_filtered = apply_filter(df_f5, 'tooling_type', sel_tt)

    if df_filtered.empty:
        st.sidebar.warning("No data matches the current filters.")
        st.stop()

    # ------------------------------------------------------------------
    # Analysis parameters (sidebar)
    # ------------------------------------------------------------------
    st.sidebar.markdown("---")
    st.sidebar.markdown("### Analysis Parameters ⚙️")
    with st.sidebar.expander("Configure Metrics", expanded=True):
        tolerance = st.slider(
            "Tolerance Band", 0.01, 0.50, 0.05, 0.01,
            key="rr_tolerance",
            help="±% around Mode CT used to classify normal vs stopped shots."
        )
        downtime_gap_tolerance = st.slider(
            "Downtime Gap (sec)", 0.0, 5.0, 2.0, 0.5,
            key="rr_downtime_gap",
            help="Minimum idle time between shots to be considered a stop."
        )
        run_interval_hours = st.slider(
            "Run Interval (hours)", 1, 24, 8, 1,
            key="rr_run_interval",
            help="Max hours between shots before a new Production Run is identified."
        )
        enable_min_shots = st.checkbox(
            "Filter Small Production Runs", value=False, key="rr_filter_enable"
        )
        min_shots_filter = (
            st.slider("Min Shots per Run", 1, 500, 10, 1, key="rr_min_shots_global")
            if enable_min_shots else 1
        )
        show_approved_ct = st.checkbox(
            "Show Approved CT", value=False, key="rr_show_approved_ct",
            help="Displays the Approved CT column in tables and metrics."
        )

    # ------------------------------------------------------------------
    # In-page tool selector (CR-aligned)
    # Sidebar is global scope. Tool selection lives in-page so it is
    # visible, contextual, and doesn't compete with filters for space.
    # ------------------------------------------------------------------
    tool_ids = sorted([
        str(x) for x in df_filtered[id_col].unique()
        if str(x).lower() not in ["nan", "unknown", "none"]
    ])

    if not tool_ids:
        st.warning("No tools found for the current filter scope.")
        st.stop()

    st.markdown("### Tool Selection")
    st.caption(
        f"{len(tool_ids)} tool{'s' if len(tool_ids) != 1 else ''} in current filter scope. "
        "Risk Tower uses all tools. Dashboard & Trends use the selection below."
    )

    col_sel, col_mode = st.columns([3, 1])
    with col_sel:
        selected_tools = st.multiselect(
            "Select tool(s) for Dashboard & Trends",
            options=tool_ids,
            default=tool_ids[:1] if tool_ids else [],
            key="rr_tool_select_inline"
        )
    with col_mode:
        if len(selected_tools) >= 2:
            view_mode = st.radio(
                "View mode", ["Rolled-Up", "Side-by-Side"],
                horizontal=True, key="rr_view_mode_inline"
            )
            if view_mode == "Side-by-Side" and len(selected_tools) > 5:
                st.caption("⚠️ Side-by-Side limited to 5 tools.")
                selected_tools = selected_tools[:5]
        else:
            view_mode = "Rolled-Up"

    st.markdown("---")

    if not selected_tools:
        st.info("Select at least one tool above to begin analysis.")
        st.stop()

    df_tool_scope = df_filtered[df_filtered[id_col].isin(selected_tools)]
    tool_name_display = (selected_tools[0] if len(selected_tools) == 1
                         else f"{len(selected_tools)} tools: {', '.join(selected_tools)}")

    # ------------------------------------------------------------------
    # Helper: render side-by-side columns
    # ------------------------------------------------------------------
    def _render_side_by_side(render_fn, *args, **kwargs):
        cols = st.columns(len(selected_tools))
        for i, t_id in enumerate(selected_tools):
            with cols[i]:
                st.markdown(
                    f"<h3 style='text-align:center;color:#3498DB;'>Tool: {t_id}</h3>",
                    unsafe_allow_html=True
                )
                t_df = df_tool_scope[df_tool_scope[id_col] == t_id]
                if not t_df.empty:
                    render_fn(t_df, t_id, *args, key_prefix=f"{t_id}_", **kwargs)
                else:
                    st.warning(f"No data for {t_id}")

    # ------------------------------------------------------------------
    # Tabs
    # ------------------------------------------------------------------
    tab1, tab2, tab3 = st.tabs(["Risk Tower", "Run Rate Dashboard", "Trends"])

    with tab1:
        render_risk_tower(df_filtered, run_interval_hours, min_shots_filter,
                          tolerance, downtime_gap_tolerance)

    with tab2:
        if view_mode == "Side-by-Side":
            _render_side_by_side(
                render_dashboard,
                tolerance, downtime_gap_tolerance, run_interval_hours,
                show_approved_ct, min_shots_filter
            )
        else:
            render_dashboard(
                df_tool_scope, tool_name_display,
                tolerance, downtime_gap_tolerance, run_interval_hours,
                show_approved_ct, min_shots_filter
            )

    with tab3:
        if view_mode == "Side-by-Side":
            _render_side_by_side(
                render_trends_tab,
                tolerance, downtime_gap_tolerance,
                run_interval_hours, min_shots_filter
            )
        else:
            render_trends_tab(
                df_tool_scope, tool_name_display,
                tolerance, downtime_gap_tolerance,
                run_interval_hours, min_shots_filter
            )


if __name__ == "__main__":
    run_run_rate_ui()
