"""EU Research Funding Pitch Dashboard."""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="EU Research Funding Dashboard", layout="wide")

st.markdown("""
<style>
.main > div {max-width: 1150px; margin: 0 auto;}
h1, h2, h3 {color: #1F3864;}
.stat-card {background:#f5f7fa; border-radius:10px; padding:14px 10px; text-align:center; height:100%;}
.stat-label {font-size:0.78rem; color:#555; margin-bottom:4px;}
.stat-value {font-size:1.55rem; font-weight:700; color:#1F3864; line-height:1.25; word-wrap:break-word; white-space:normal;}
.stat-sub {font-size:0.72rem; color:#777; margin-top:2px;}
.link-card {background:#eef2f8; border-radius:10px; padding:18px; text-align:center;}
.link-card a {font-weight:600; color:#1F3864; text-decoration:none; font-size:1.05rem;}
.tagcloud {display:flex; flex-wrap:wrap; gap:10px 16px; align-items:baseline; padding:10px 4px;}
.tagcloud span {color:#1F3864;}
.stPlotlyChart, .js-plotly-plot, .plotly {width: 100% !important;}
</style>
""", unsafe_allow_html=True)


def stat_card(col, label, value, sub=None):
    # Single-line HTML on purpose: multi-line indented HTML passed to
    # st.markdown(unsafe_allow_html=True) can have its closing tags treated
    # as literal text instead of markup.
    sub_html = f'<div class="stat-sub">{sub}</div>' if sub else ""
    html = f'<div class="stat-card"><div class="stat-label">{label}</div><div class="stat-value">{value}</div>{sub_html}</div>'
    col.markdown(html, unsafe_allow_html=True)


st.title("🎓 EU Research Funding — Pitch Dashboard")
st.caption("Upload the DataKey.xlsx export to generate slide-ready visuals. "
           "Every chart is interactive — zoom, pan, toggle legend items, and export as PNG from its own toolbar "
           "(hover top-right corner of a chart to see it).")
uploaded = st.file_uploader("Upload DataKey.xlsx", type=["xlsx"])
if not uploaded:
    st.info("Waiting for a file.")
    st.stop()

raw = pd.read_excel(uploaded, sheet_name=None, engine="openpyxl")


def find(cols_required):
    for df in raw.values():
        if all(any(str(c).lower() == col.lower() for c in df.columns) for col in cols_required):
            return df.copy()
    return None


def get_col(df, name):
    return next((c for c in df.columns if str(c).lower() == name.lower()), None)


def eu_num(s):
    if isinstance(s, (int, float)):
        return float(s)
    if not isinstance(s, str):
        return None
    t = s.strip().replace("%", "")
    if t in ("", "-"):
        return None
    try:
        return float(t.replace(".", "").replace(",", "."))
    except ValueError:
        return None


def parse_label_value(cell):
    if not isinstance(cell, str):
        return cell
    cell = cell.replace("_x000D_", "").strip()
    return cell.split(":", 1)[1].strip() if ":" in cell else cell


# ---------- Locate sheets ----------
df_status = find(["Project Status", "Signed Grants"])
df_pillar = find(["Pillar", "Net EU Contribution (EUR)"])
df_contrib_total = find(["Net EU Contribution", "of total"])
df_fp = find(["Framework Programme", "Participation"])
df_partic_total = find(["Participation", "of total"])
df_thematic = find(["Participation", "Net EU Contribution (EUR)", "Participant Cost (EUR)"])
df_rank = find(["Participation rank"])
df_eic = find(["EIC Participation"])
df_years = find(["Signature Year", "Participation", "Cumulative Participation"])
df_collab = find(["Collaborative organisation legal name", "Collaboration links"])
df_grants_total = find(["Signed grants", "of total"])
df_role = find(["[Partner Role]", "Participation"])
df_erc = find(["ERC Principal Investigators"])
df_dept = find(["Department Name", "Pro rata Department EU Contribution (EUR)"])
df_msca = find(["MSCA Participation"])
df_keywords = find(["Keywords", "Number of Keywords"])
df_org = find(["Organization PIC", "VAT Number"])
df_keyfig = find(["Indicator", "Organisation in HE"])

# ================= Exact-match department classification (computed once, reused everywhere) =================
dept_agg = pd.DataFrame(columns=["Group", "value"])
central_total = unknown_total = 0.0
unknown_rows = new_to_file = pd.DataFrame()
name_col = val_col = None

if df_dept is not None:
    name_col = get_col(df_dept, "Department Name")
    val_col = get_col(df_dept, "Pro rata Department EU Contribution (EUR)")
    d = df_dept[[name_col, val_col]].dropna()
    d = d[d[name_col].str.strip().str.lower() != "totals"]

    FACULTY, CENTRAL, UNKNOWN = "faculty", "Central & university-wide units", "Unspecified in source data"

    DEPT_MAP = {
        # International Institute of Social Studies
        "International Institute of Social Studies": ("International Institute of Social Studies (ISS)", FACULTY),
        "ISS - International Institute of Social Studies": ("International Institute of Social Studies (ISS)", FACULTY),
        "Institute for Social Studies": ("International Institute of Social Studies (ISS)", FACULTY),
        "International Institute For Social Studies": ("International Institute of Social Studies (ISS)", FACULTY),
        "International Institute Of Social Studies": ("International Institute of Social Studies (ISS)", FACULTY),
        "International Institute of Social Studies (ISS)": ("International Institute of Social Studies (ISS)", FACULTY),
        # Erasmus School of Social and Behavioural Sciences (incl. DPAS, DPECS, Sociology, DRIFT)
        "Erasmus School of Social and Behavioural Sciences": ("Erasmus School of Social and Behavioural Sciences (ESSB)", FACULTY),
        "Department of Public Administration and Sociology": ("Erasmus School of Social and Behavioural Sciences (ESSB)", FACULTY),
        "Sociology / Faculty of Social Sciences": ("Erasmus School of Social and Behavioural Sciences (ESSB)", FACULTY),
        "Department of Psychology, Education, and Child Studies [DPECS]": ("Erasmus School of Social and Behavioural Sciences (ESSB)", FACULTY),
        "Faculty of Social Sciences": ("Erasmus School of Social and Behavioural Sciences (ESSB)", FACULTY),
        "Faculty of Social Sciences, Department of Sociology": ("Erasmus School of Social and Behavioural Sciences (ESSB)", FACULTY),
        "Erasmus University Rotterdam, Faculty of Social Sciences": ("Erasmus School of Social and Behavioural Sciences (ESSB)", FACULTY),
        "Faculty of Social Sciences, Department of Psychology, Education & Child Studies": ("Erasmus School of Social and Behavioural Sciences (ESSB)", FACULTY),
        "Erasmus School of Social Sciences and Behaviour": ("Erasmus School of Social and Behavioural Sciences (ESSB)", FACULTY),
        "Faculty of Social Sciences, Dept. of Public Administration": ("Erasmus School of Social and Behavioural Sciences (ESSB)", FACULTY),
        "Department of Psychology, Education and Child Studies": ("Erasmus School of Social and Behavioural Sciences (ESSB)", FACULTY),
        "Department Of Sociology": ("Erasmus School of Social and Behavioural Sciences (ESSB)", FACULTY),
        "Faculty of Social Sciences, Dept. of Public Admin & Sociology": ("Erasmus School of Social and Behavioural Sciences (ESSB)", FACULTY),
        "SSB Department of Public Administration and Sociology": ("Erasmus School of Social and Behavioural Sciences (ESSB)", FACULTY),
        "Erasmus School of Social and Behaviour Sciences (ESSB)": ("Erasmus School of Social and Behavioural Sciences (ESSB)", FACULTY),
        "Department of Public Administration": ("Erasmus School of Social and Behavioural Sciences (ESSB)", FACULTY),
        "Department Of Public Administration": ("Erasmus School of Social and Behavioural Sciences (ESSB)", FACULTY),
        "Public Administration and Sociology": ("Erasmus School of Social and Behavioural Sciences (ESSB)", FACULTY),
        "Faculty of Social Sciences, Department of Public Administration": ("Erasmus School of Social and Behavioural Sciences (ESSB)", FACULTY),
        "Drift": ("Erasmus School of Social and Behavioural Sciences (ESSB)", FACULTY),
        "Dutch Research Institute For Transitions – DRIFT": ("Erasmus School of Social and Behavioural Sciences (ESSB)", FACULTY),
        # Erasmus School of Health Policy & Management (incl. legacy "iBMG", iMTA, HTA)
        "Erasmus School of Health Policy & Management": ("Erasmus School of Health Policy & Management (ESHPM)", FACULTY),
        "health technology assessment": ("Erasmus School of Health Policy & Management (ESHPM)", FACULTY),
        "Erasmus School of Health Policy and Management": ("Erasmus School of Health Policy & Management (ESHPM)", FACULTY),
        "Institute of Health Policy and Management (iBMG)": ("Erasmus School of Health Policy & Management (ESHPM)", FACULTY),
        "Institute for Medical Technology Assessment (iMTA)": ("Erasmus School of Health Policy & Management (ESHPM)", FACULTY),
        "Institute Of Health Policy And Management": ("Erasmus School of Health Policy & Management (ESHPM)", FACULTY),
        "Health Policy And Management": ("Erasmus School of Health Policy & Management (ESHPM)", FACULTY),
        "Institute for Health Policy and Management": ("Erasmus School of Health Policy & Management (ESHPM)", FACULTY),
        "Institute of Health Policy and Management": ("Erasmus School of Health Policy & Management (ESHPM)", FACULTY),
        "Health Economics": ("Erasmus School of Health Policy & Management (ESHPM)", FACULTY),
        "Health Policy & Management (iBMG)": ("Erasmus School of Health Policy & Management (ESHPM)", FACULTY),
        "Erasmus School of Health Policy & Management, section HE": ("Erasmus School of Health Policy & Management (ESHPM)", FACULTY),
        "ESHPM (Erasmus School of Health Policy and Management)": ("Erasmus School of Health Policy & Management (ESHPM)", FACULTY),
        "The Institute of Health Policy & Management (iBMG)": ("Erasmus School of Health Policy & Management (ESHPM)", FACULTY),
        "Institute for Medical Technology Assessment": ("Erasmus School of Health Policy & Management (ESHPM)", FACULTY),
        "Erasmus School of Health Policy and Law": ("Erasmus School of Health Policy & Management (ESHPM)", FACULTY),
        "Health Technology Assessment (HTA)": ("Erasmus School of Health Policy & Management (ESHPM)", FACULTY),
        "Institute of Health Policy & Manangement": ("Erasmus School of Health Policy & Management (ESHPM)", FACULTY),
        "ESHPM": ("Erasmus School of Health Policy & Management (ESHPM)", FACULTY),
        "Erasmus School of Health Policy and Management (ESHPM)": ("Erasmus School of Health Policy & Management (ESHPM)", FACULTY),
        "Health Technology Assessment": ("Erasmus School of Health Policy & Management (ESHPM)", FACULTY),
        "Institute Of Health Policy And Management Erasmus University Rotterdam": ("Erasmus School of Health Policy & Management (ESHPM)", FACULTY),
        # Rotterdam School of Management, Erasmus University — official name confirmed via eur.nl regulations page
        "Rotterdam School of Management": ("Rotterdam School of Management, Erasmus University (RSM)", FACULTY),
        "Rotterdam School of Management, Erasmus University": ("Rotterdam School of Management, Erasmus University (RSM)", FACULTY),
        "Erasmus University Rotterdam, Rotterdam School of Management": ("Rotterdam School of Management, Erasmus University (RSM)", FACULTY),
        "Faculty of Management": ("Rotterdam School of Management, Erasmus University (RSM)", FACULTY),
        "Rotterdam School of Management (RSM)": ("Rotterdam School of Management, Erasmus University (RSM)", FACULTY),
        "Rotterdam School of Management /Business-Society Management dept": ("Rotterdam School of Management, Erasmus University (RSM)", FACULTY),
        "Department of Organisation and Personnel Management": ("Rotterdam School of Management, Erasmus University (RSM)", FACULTY),
        "Rotterdam School of Managem": ("Rotterdam School of Management, Erasmus University (RSM)", FACULTY),
        # Erasmus School of Economics
        "Erasmus School of Economics": ("Erasmus School of Economics (ESE)", FACULTY),
        "Erasmus School of Economics - Department of Business Economics": ("Erasmus School of Economics (ESE)", FACULTY),
        "Department of Economics": ("Erasmus School of Economics (ESE)", FACULTY),
        "Applied Economics": ("Erasmus School of Economics (ESE)", FACULTY),
        "Erasmus School of Economics - Business Economics": ("Erasmus School of Economics (ESE)", FACULTY),
        "Department of Economics, Erasmus School of Economics": ("Erasmus School of Economics (ESE)", FACULTY),
        "Economics Department, Erasmus School of Economics": ("Erasmus School of Economics (ESE)", FACULTY),
        "Erasmus School Of Economics": ("Erasmus School of Economics (ESE)", FACULTY),
        "Econometric Institute": ("Erasmus School of Economics (ESE)", FACULTY),
        "Erasmus School Of Economics, Marketing Department": ("Erasmus School of Economics (ESE)", FACULTY),
        "Erasmus School Of Economics, Marketing Research Group": ("Erasmus School of Economics (ESE)", FACULTY),
        # Erasmus School of Law
        "Erasmus School of Law/Private International and Comparative Law": ("Erasmus School of Law (ESL)", FACULTY),
        "Erasmus School of Law": ("Erasmus School of Law (ESL)", FACULTY),
        "ERASMUS SCHOOL OF LAW": ("Erasmus School of Law (ESL)", FACULTY),
        "Criminology Department, Erasmus School of Law": ("Erasmus School of Law (ESL)", FACULTY),
        # Erasmus School of History, Culture and Communication — name confirmed current (2024 dean announcement)
        "Arts and Culture Studies": ("Erasmus School of History, Culture and Communication (ESHCC)", FACULTY),
        "Media and Communication": ("Erasmus School of History, Culture and Communication (ESHCC)", FACULTY),
        "Erasmus School of History, Culture and Communication": ("Erasmus School of History, Culture and Communication (ESHCC)", FACULTY),
        "Department of Media and Communication": ("Erasmus School of History, Culture and Communication (ESHCC)", FACULTY),
        "Erasmus School of History , Culture and Communication": ("Erasmus School of History, Culture and Communication (ESHCC)", FACULTY),
        "Erasmus School for History, Culture and Communication": ("Erasmus School of History, Culture and Communication (ESHCC)", FACULTY),
        "Erasmus Research Centre for Media, Communication and Culture": ("Erasmus School of History, Culture and Communication (ESHCC)", FACULTY),
        "Erasmus School Of History, Art And Communication": ("Erasmus School of History, Culture and Communication (ESHCC)", FACULTY),
        "ERASMUS SCHOOL OF HISTORY CULTURE AND COMMUNICATION": ("Erasmus School of History, Culture and Communication (ESHCC)", FACULTY),
        "Department of media and communication": ("Erasmus School of History, Culture and Communication (ESHCC)", FACULTY),
        "Department of Arts and Culture Studies": ("Erasmus School of History, Culture and Communication (ESHCC)", FACULTY),
        "Erasmus School of History, Culture and Communication, Department": ("Erasmus School of History, Culture and Communication (ESHCC)", FACULTY),
        "ESHCC, Arts and Culture Studies": ("Erasmus School of History, Culture and Communication (ESHCC)", FACULTY),
        # Erasmus School of Philosophy
        "Erasmus School of Philosophy": ("Erasmus School of Philosophy (ESPhil)", FACULTY),
        "Erasmus School of Philosophy, Erasmus University Rotterdam": ("Erasmus School of Philosophy (ESPhil)", FACULTY),
        "Faculty of Philosophy": ("Erasmus School of Philosophy (ESPhil)", FACULTY),
        "Erasmus School of Philosophy (ESPhil)": ("Erasmus School of Philosophy (ESPhil)", FACULTY),
        # Central & university-wide units (not a faculty)
        "CVB": (CENTRAL, CENTRAL),
        "General Governance of EUR - UNIC": (CENTRAL, CENTRAL),
        "Erasmus University Rotterdam": (CENTRAL, CENTRAL),
        "Marketing & Communications Department": (CENTRAL, CENTRAL),
        "Erasmus Centre for Data Analytics (ECDA)": (CENTRAL, CENTRAL),
        "Marketing and Communications": (CENTRAL, CENTRAL),
        "Erasmus Research Services": (CENTRAL, CENTRAL),
        "University Library": (CENTRAL, CENTRAL),
        # Genuinely unspecified in the source data itself — not guessed
        "Missing": (UNKNOWN, UNKNOWN),
        "new department": (UNKNOWN, UNKNOWN),
        "HSMO": (UNKNOWN, UNKNOWN),
        "Staff Group 2": (UNKNOWN, UNKNOWN),
        "Political Ecology Research Group": (UNKNOWN, UNKNOWN),
        "Society, Youth and Neuroscience Connected (SYNC) lab": (UNKNOWN, UNKNOWN),
    }

    def classify(name):
        key = name.strip()
        return DEPT_MAP.get(key, (f"Not yet classified: {key}", "new"))

    mapped = d[name_col].apply(classify)
    d["Group"] = mapped.apply(lambda t: t[0])
    d["Kind"] = mapped.apply(lambda t: t[1])

    faculty_rows = d[d["Kind"] == FACULTY]
    central_rows = d[d["Kind"] == CENTRAL]
    unknown_rows = d[d["Kind"] == UNKNOWN]
    new_to_file = d[d["Kind"] == "new"]
    central_total = central_rows[val_col].sum()
    unknown_total = unknown_rows[val_col].sum()
    dept_agg = faculty_rows.groupby("Group", as_index=False)[val_col].sum().sort_values(val_col, ascending=False)

# ================= Thematic-priority averages (computed once, reused) =================
thematic_avg = pd.DataFrame()
if df_thematic is not None:
    th = df_thematic.copy()
    th.columns = ["Thematic Priority", "Participation", "Net EU Contribution (EUR)", "Participant Cost (EUR)"]
    th_robust = th[th["Participation"] >= 3].copy()
    th_robust["Avg award (EUR)"] = th_robust["Net EU Contribution (EUR)"] / th_robust["Participation"]
    thematic_avg = th_robust.sort_values("Avg award (EUR)", ascending=False)

# ================= YoY growth (computed once) =================
yoy_df = pd.DataFrame()
if df_years is not None:
    yc, pc = df_years.columns[0], df_years.columns[1]
    yoy_df = df_years[[yc, pc]].copy()
    yoy_df["YoY change (%)"] = yoy_df[pc].pct_change() * 100

# ================= Sidebar: TOC, filters, tips =================
TOC = [
    ("Headline numbers", "headline-numbers"),
    ("Executive summary", "exec-summary"),
    ("A — Recreating the organisation dashboard", "section-a"),
    ("Organisation details", "org-details"),
    ("Evolution of participation", "evolution"),
    ("Year-over-year growth", "yoy"),
    ("Participation per Framework Programme", "fp-programme"),
    ("Net EU contribution by pillar", "pillar"),
    ("Funding by faculty", "faculty-funding"),
    ("Project keywords", "keywords"),
    ("Top collaboration partners", "top-partners"),
    ("Full partner network", "full-partners"),
    ("Collaboration map", "collab-map"),
    ("Key figures vs national totals", "key-figures"),
    ("C — Funding vs. participation efficiency", "section-c"),
    ("B — Average award size", "section-b"),
]
with st.sidebar:
    st.header("Contents")
    for title, anchor in TOC:
        st.markdown(f'<a href="#{anchor}" style="text-decoration:none;color:#1F3864;">{title}</a>',
                     unsafe_allow_html=True)
    st.divider()

    st.header("Filters")
    fp_options = df_fp[df_fp.columns[0]].tolist() if df_fp is not None else []
    selected_fp = st.multiselect("Framework Programme", fp_options, default=fp_options) if fp_options else []
    if df_years is not None:
        yc = df_years.columns[0]
        ymin, ymax = int(df_years[yc].min()), int(df_years[yc].max())
        year_range = st.slider("Year range", ymin, ymax, (ymin, ymax))
    else:
        year_range = None
    st.caption("These two filters only apply to the **Evolution/Year-over-year** and **Framework Programme** "
               "sections, plus the columns shown in the **Key figures** table. Everything else in this file "
               "(faculty, pillar, thematic priority, keywords, partners) is exported as a single all-time total "
               "with no year or programme breakdown, so it can't be filtered without inventing numbers — it stays "
               "fixed regardless of the filters above.")
    st.divider()
    st.caption("Tip: click-drag on any chart to zoom into a region, double-click to reset. "
               "Hover over the top-right of a chart for a toolbar with a camera icon — "
               "that downloads the chart as a PNG, ready to drop into a slide. "
               "Click legend entries to show/hide a series.")

# ================= HEADLINE KPIs =================
st.header("Headline numbers", anchor="headline-numbers")
row1 = st.columns(4)
if df_contrib_total is not None:
    stat_card(row1[0], "Net EU Contribution", f"€{df_contrib_total.iloc[0, 0] / 1e6:.1f}M")
if df_partic_total is not None:
    stat_card(row1[1], "Participations", int(df_partic_total.iloc[0, 0]))
if df_grants_total is not None:
    stat_card(row1[2], "Signed Grants", int(df_grants_total.iloc[0, 0]))
if df_rank is not None:
    stat_card(row1[3], "National Rank", df_rank.iloc[0, 0])

row2_items = []
if df_erc is not None:
    row2_items.append(("ERC Principal Investigators", int(df_erc.iloc[0, 0])))
if df_msca is not None:
    row2_items.append(("MSCA Participation", int(df_msca.iloc[0, 0])))
if df_eic is not None and int(df_eic.iloc[0, 0]) > 0:
    row2_items.append(("EIC Participation", int(df_eic.iloc[0, 0])))
if row2_items:
    row2 = st.columns(len(row2_items))
    for c, (label, val) in zip(row2, row2_items):
        stat_card(c, label, val)

# ================= EXECUTIVE SUMMARY =================
st.divider()
st.header("Executive summary", anchor="exec-summary")
st.caption("All-time snapshot — not affected by the sidebar filters (those apply to the detailed sections below).")

ex1, ex2, ex3, ex4 = st.columns(4)
if df_contrib_total is not None:
    stat_card(ex1, "Total EU Contribution", f"€{df_contrib_total.iloc[0, 0] / 1e6:.1f}M")
if df_fp is not None:
    fp_col, fp_part = df_fp.columns[:2]
    fp_vals = dict(zip(df_fp[fp_col], df_fp[fp_part]))
    fp7 = fp_vals.get("FP7")
    he = fp_vals.get("HORIZON EUROPE") or fp_vals.get("HE") or list(fp_vals.values())[-1]
    if fp7:
        stat_card(ex2, "Participation growth since FP7", f"{fp7} → {list(fp_vals.values())[1]} → {he}",
                   "FP7 → H2020 → Horizon Europe (raw counts, unequal period lengths)")
if not dept_agg.empty:
    top_fac = dept_agg.iloc[0]
    stat_card(ex3, "Largest faculty", top_fac["Group"].split(" (")[0], f"€{top_fac[val_col]:,.0f}")
if not thematic_avg.empty:
    top_them = thematic_avg.iloc[0]
    stat_card(ex4, "Largest thematic priority by avg. award", top_them["Thematic Priority"],
               f"€{top_them['Avg award (EUR)']:,.0f}/project (n={int(top_them['Participation'])})")

ex5, ex6, ex7 = st.columns(3)
if df_collab is not None:
    org_col, link_col = df_collab.columns[:2]
    top_partner = df_collab.sort_values(link_col, ascending=False).iloc[0]
    stat_card(ex5, "Most important collaborator", top_partner[org_col], f"{int(top_partner[link_col])} shared project links")
if df_keyfig is not None:
    part_row = df_keyfig[df_keyfig["Indicator"] == "Participation"]
    if not part_row.empty:
        latest_pct = None
        for prog in ["HE", "H2020", "FP7"]:
            v = eu_num(part_row.iloc[0].get(f"% in {prog}"))
            if v is not None:
                latest_pct = (prog, v)
                break
        if latest_pct:
            stat_card(ex6, "EUR's share of Dutch total (participation)", f"{latest_pct[1]}%".replace(".", ","),
                       f"most recent programme: {latest_pct[0]}")
if not yoy_df.empty and yoy_df.shape[0] >= 4:
    recent3 = yoy_df[pc].iloc[-3:].mean()
    prior3 = yoy_df[pc].iloc[-6:-3].mean() if yoy_df.shape[0] >= 6 else yoy_df[pc].iloc[:-3].mean()
    if prior3:
        growth = (recent3 - prior3) / prior3 * 100
        stat_card(ex7, "Recent growth rate", f"{growth:+.0f}%",
                   "avg. new participations/year, last 3 years vs the 3 before that")

# ================= SECTION A =================
st.divider()
st.header("A — Recreating the organisation dashboard", anchor="section-a")
st.caption("Matches the panels in the reference PDF, using the current export.")

if df_org is not None:
    st.markdown('<div id="org-details"></div>', unsafe_allow_html=True)
    with st.expander("Organisation details (legal/registry info — not usually slide material, kept for completeness)"):
        vals = {c: parse_label_value(df_org[c].iloc[0]) for c in df_org.columns}
        cols = st.columns(4)
        for i, (label, value) in enumerate(vals.items()):
            stat_card(cols[i % 4], label, value if value else "—")

c1, c2 = st.columns(2)
if df_status is not None:
    st_col, cnt_col = df_status.columns[:2]
    fig = px.pie(df_status, names=st_col, values=cnt_col, hole=0.55,
                 title="Project Status", color_discrete_sequence=["#C00000", "#2E75B6"])
    fig.update_layout(margin=dict(t=40, b=10))
    c1.plotly_chart(fig, use_container_width=True)
if df_role is not None:
    role_col, part_col = df_role.columns[:2]
    total = df_role[part_col].sum()
    fig = px.pie(df_role, names=role_col, values=part_col, hole=0.55,
                 title="Role in Projects", color_discrete_sequence=["#1F3864", "#8FAADC"])
    fig.update_layout(margin=dict(t=40, b=10))
    c2.plotly_chart(fig, use_container_width=True)
    for _, r in df_role.iterrows():
        c2.caption(f"{r[role_col]}: {r[part_col] / total * 100:.1f}% ({int(r[part_col])} of {int(total)})")

# ---- Evolution timeline, filtered by year range, with EU programme era markers ----
if df_years is not None:
    st.subheader("Evolution of participation", anchor="evolution")
    year_col, part_col, cum_col = df_years.columns[:3]
    dfy = df_years[(df_years[year_col] >= year_range[0]) & (df_years[year_col] <= year_range[1])] if year_range else df_years
    fig = go.Figure()
    fig.add_bar(x=dfy[year_col], y=dfy[part_col], name="New participations that year", marker_color="#8FAADC")
    fig.add_scatter(x=dfy[year_col], y=dfy[cum_col], name="Cumulative participation (running total since 2007)",
                     line=dict(color="#C00000", width=3), yaxis="y2")
    for yr, label in [(2014, "H2020 starts"), (2021, "Horizon Europe starts")]:
        if dfy[year_col].min() <= yr <= dfy[year_col].max():
            fig.add_vline(x=yr, line_dash="dot", line_color="#999")
            fig.add_annotation(x=yr, y=1.05, yref="paper", text=label, showarrow=False,
                                font=dict(size=10, color="#777"))
    fig.update_layout(yaxis=dict(title="New participations per year"),
                       yaxis2=dict(title="Cumulative participation (running total)", overlaying="y", side="right"),
                       legend=dict(orientation="h", y=1.18), margin=dict(t=50))
    st.plotly_chart(fig, use_container_width=True)
    recent = df_years[df_years[year_col] >= 2022][part_col].sum()
    st.caption(f"{int(recent)} of {int(df_years[cum_col].iloc[-1])} total participations to date came from 2022 onward. "
               "\"Cumulative participation\" is the running total of new participations added year by year since 2007 "
               "(right-hand axis, values taken directly from the source, not recalculated). Programme start years "
               "shown are the EU's official framework programme calendar (public knowledge, not from this file).")

    st.subheader("Year-over-year growth", anchor="yoy")
    yoy_view = yoy_df[(yoy_df[year_col] >= year_range[0]) & (yoy_df[year_col] <= year_range[1])] if year_range else yoy_df
    yoy_view = yoy_view.dropna(subset=["YoY change (%)"])
    fig = px.bar(yoy_view, x=year_col, y="YoY change (%)",
                 color=yoy_view["YoY change (%)"] > 0,
                 color_discrete_map={True: "#2E7D32", False: "#C00000"})
    fig.update_layout(showlegend=False, margin=dict(t=10), yaxis_title="YoY change in new participations (%)")
    st.plotly_chart(fig, use_container_width=True)
    st.caption("Year-over-year % change in new participations (not cumulative) — a single low or high year can "
               "swing this a lot, which is exactly why the 'recent growth rate' in the executive summary above "
               "uses a 3-year average instead.")

if df_fp is not None:
    st.subheader("Participation per Framework Programme", anchor="fp-programme")
    fp_col, part_col = df_fp.columns[:2]
    dfp = df_fp[df_fp[fp_col].isin(selected_fp)] if selected_fp else df_fp
    fig = px.bar(dfp, x=fp_col, y=part_col, text=part_col, color=fp_col,
                 color_discrete_sequence=px.colors.sequential.Blues_r[:3])
    fig.update_traces(textposition="outside", cliponaxis=False)
    fig.update_layout(showlegend=False, margin=dict(t=10))
    fig.update_yaxes(automargin=True)
    st.plotly_chart(fig, use_container_width=True)
    st.caption("FP7 (2007–2013), H2020 (2014–2020) and Horizon Europe (2021–2027, still running) cover different "
               "numbers of years, so these totals aren't a like-for-like comparison — roughly 8–9/year under FP7, "
               "~14/year under H2020, ~14/year so far under Horizon Europe (using public EU programme calendar years).")

if df_pillar is not None:
    st.subheader("Net EU contribution by pillar", anchor="pillar")
    pil_col, pval_col = df_pillar.columns[:2]
    pil = df_pillar.sort_values(pval_col)
    fig = px.bar(pil, x=pval_col, y=pil_col, orientation="h", color_discrete_sequence=["#1F3864"])
    fig.update_layout(margin=dict(t=10), xaxis_title="Net EU Contribution (EUR)", yaxis_title="")
    fig.update_yaxes(automargin=True)
    st.plotly_chart(fig, use_container_width=True)
    st.caption("Pillar names are shown as exported; categories from different framework programmes are not merged. "
               "Not filterable by year/programme — this sheet only has an all-time total per pillar.")

# ---- Departments (uses the classification computed once, above) ----
if df_dept is not None:
    st.subheader("Funding by faculty", anchor="faculty-funding")
    fig = px.bar(dept_agg.sort_values(val_col), x=val_col, y="Group", orientation="h",
                 color_discrete_sequence=["#1F3864"])
    fig.update_layout(margin=dict(t=10), xaxis_title="Pro-rata EU Contribution (EUR)", yaxis_title="")
    fig.update_yaxes(automargin=True)
    st.plotly_chart(fig, use_container_width=True)

    st.caption(f"Every one of the {d[name_col].nunique()} distinct raw department-name strings in this file was "
               f"hand-matched, against EUR's official list of 8 faculties/schools (verified against eur.nl), to an "
               f"exact lookup table — not a fuzzy match. €{central_total:,.0f} sits with central/university-wide "
               f"units (not a faculty), and €{unknown_total:,.0f} across {unknown_rows.shape[0]} rows is genuinely "
               f"unlabelled in the source data itself (e.g. literally listed as \"Missing\"). "
               "Not filterable by year/programme — this sheet only has an all-time total per department.")
    if (d[name_col].str.contains("drift", case=False, na=False)).any():
        st.caption("Note: the R&I export shows no funding recorded for DRIFT (Dutch Research Institute For "
                   "Transitions) — an explicit €0 in the source data, not a gap in this app's processing.")
    if not new_to_file.empty:
        st.warning(f"{new_to_file.shape[0]} department name(s) in this file were not in the lookup table used to "
                    "build this app — likely a different export. Listed below for manual classification:")
        st.dataframe(new_to_file[[name_col, val_col]], use_container_width=True)

# ---- Keywords as a tag cloud, not a bar chart ----
if df_keywords is not None:
    st.subheader("Project keywords", anchor="keywords")
    kw_col, n_col = df_keywords.columns[:2]
    top = df_keywords.nlargest(30, n_col)
    mn, mx = top[n_col].min(), top[n_col].max()

    def size_for(n):
        if mx == mn:
            return 18
        return 14 + (n - mn) / (mx - mn) * 26  # 14px .. 40px

    spans = "".join(
        f'<span style="font-size:{size_for(r[n_col]):.0f}px;font-weight:{600 if r[n_col] >= (mn+mx)/2 else 400};" '
        f'title="{int(r[n_col])} projects">{r[kw_col]}</span>'
        for _, r in top.iterrows()
    )
    st.markdown(f'<div class="tagcloud">{spans}</div>', unsafe_allow_html=True)
    st.caption(f"Top {top.shape[0]} of {df_keywords.shape[0]} keywords, sized by number of tagged projects "
               "(hover for exact count). Not filterable by year/programme.")

# ---- Collaborations: leaderboard + full picture + full table ----
if df_collab is not None:
    org_col, link_col = df_collab.columns[:2]

    st.subheader("Top collaboration partners", anchor="top-partners")
    top_collab = df_collab.nlargest(15, link_col).sort_values(link_col)
    fig = px.bar(top_collab, x=link_col, y=org_col, orientation="h", color_discrete_sequence=["#2E75B6"])
    fig.update_layout(margin=dict(t=10), xaxis_title="Shared project links", yaxis_title="")
    fig.update_yaxes(automargin=True)
    st.plotly_chart(fig, use_container_width=True)
    st.caption(f"Top 15 of {df_collab.shape[0]} distinct partner organisations. Not filterable by year/programme.")

    st.subheader("Full partner network", anchor="full-partners")
    fig2 = px.treemap(df_collab, path=[org_col], values=link_col, color=link_col,
                       color_continuous_scale="Blues")
    fig2.update_layout(margin=dict(t=10, l=0, r=0, b=0))
    st.plotly_chart(fig2, use_container_width=True)
    st.caption(f"All {df_collab.shape[0]} partner organisations — box size = number of shared project links.")

    with st.expander(f"Browse or search all {df_collab.shape[0]} partner organisations"):
        query = st.text_input("Filter by organisation name", "")
        table = df_collab.sort_values(link_col, ascending=False)
        if query:
            table = table[table[org_col].str.contains(query, case=False, na=False)]
        st.dataframe(table, use_container_width=True, height=400, hide_index=True)

st.subheader("Collaboration map", anchor="collab-map")
st.markdown(
    '<div class="link-card">View the interactive collaboration map on the EU R&I Profiles dashboard:<br>'
    '<a href="https://dashboard.tech.ec.europa.eu/qs_digit_dashboard_mt/public/sense/app/dc5f6f40-c9de-4c40-8648-015d6ff21342" target="_blank">Open collaboration map ↗</a></div>',
    unsafe_allow_html=True
)
st.caption("Opens in the EU's own dashboard (it can't be reliably embedded here).")

if df_keyfig is not None:
    st.subheader("Key figures — EUR vs national totals, by Framework Programme", anchor="key-figures")
    progs_to_show = [p for p in ["HE", "H2020", "FP7"] if p in selected_fp] or ["HE", "H2020", "FP7"]
    all_indicators = df_keyfig["Indicator"].tolist()
    shown_indicators = st.multiselect("Rows to show", all_indicators, default=all_indicators, key="keyfig_rows")
    rows = []
    for _, r in df_keyfig.iterrows():
        indicator = r["Indicator"]
        if indicator not in shown_indicators:
            continue
        is_eur_amt = "(EUR)" in indicator
        entry = {"Indicator": indicator}
        for prog in progs_to_show:
            org = eu_num(r.get(f"Organisation in {prog}"))
            pct = r.get(f"% in {prog}")
            if org is None:
                entry[prog] = "—"
            else:
                amount = f"€{org:,.0f}" if is_eur_amt else f"{org:,.0f}"
                entry[prog] = f"{amount} ({pct} of NL total)" if isinstance(pct, str) and pct != "-" else amount
        rows.append(entry)
    if rows:
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    else:
        st.caption("No rows selected.")
    st.caption("Columns follow the Framework Programme filter in the sidebar; rows can be toggled above — e.g. "
               "hide \"Net EU Contribution (EUR)\" if you'd rather not show its blank FP7 cell on a slide.")

    share_rows = []
    for _, r in df_keyfig.iterrows():
        if r["Indicator"] in ("Net EU Contribution (EUR)", "Participation"):
            for prog in ["FP7", "H2020", "HE"]:
                pct = eu_num(r.get(f"% in {prog}"))
                if pct is not None:
                    share_rows.append({"Programme": prog, "Indicator": r["Indicator"], "Share of NL total (%)": pct})
    if share_rows:
        share_df = pd.DataFrame(share_rows)
        share_df["Programme"] = pd.Categorical(share_df["Programme"], ["FP7", "H2020", "HE"], ordered=True)
        fig = px.line(share_df.sort_values("Programme"), x="Programme", y="Share of NL total (%)",
                      color="Indicator", markers=True, color_discrete_sequence=["#1F3864", "#C00000"])
        fig.update_layout(margin=dict(t=10))
        st.plotly_chart(fig, use_container_width=True)
        st.caption("EUR's share of the Dutch national total, across successive framework programmes — a real "
                   "upward trend, not present as a chart in the source dashboard.")

with st.expander("Not included — needs data this export doesn't contain"):
    st.markdown("""
- **Project-level list** — this export contains pre-aggregated summaries rather than one row per project.
- **True per-project cost-vs-contribution scatter** — needs per-project figures. Section C below gives the closest
  honest substitute available in this export: the same three metrics (participation, EU contribution, cost),
  aggregated at thematic-priority level instead of per project.
""")

# ================= SECTION C: efficiency view =================
st.divider()
st.header("C — Funding vs. participation efficiency", anchor="section-c")
if not thematic_avg.empty:
    fig = px.scatter(th, x="Participation", y="Net EU Contribution (EUR)", size="Participant Cost (EUR)",
                      color="Thematic Priority", hover_name="Thematic Priority",
                      size_max=45)
    fig.update_layout(margin=dict(t=10), showlegend=False,
                       xaxis_title="Number of participations", yaxis_title="Net EU Contribution (EUR)")
    st.plotly_chart(fig, use_container_width=True)
    st.caption("Each bubble is one thematic priority (hover for its name) — x = how many projects, y = total EU "
               "contribution, bubble size = total participant cost. Upper-left bubbles get a lot of funding from "
               "few projects (high per-project value); lower-right bubbles spread funding across many projects. "
               "This is the closest honest equivalent to the PDF's missing cost-vs-contribution scatter, at "
               "thematic-priority granularity rather than per-project.")
    st.warning(
        "A faculty-level version of this view isn't possible from this export: the department sheet only has a "
        "pro-rata EU contribution per department — it has no participation count and no participant cost per "
        "department. Building those numbers would mean guessing/estimating them, which isn't something this app "
        "does. If a future export adds participation and cost broken down by department, this chart is a direct "
        "drop-in."
    )
else:
    st.info("Thematic priority data not found in this export — efficiency view unavailable.")

# ================= SECTION B =================
st.divider()
st.header("B — Average award size by thematic priority", anchor="section-b")
st.caption("Contribution ÷ participation — a ratio not shown anywhere in the source dashboard.")
if not thematic_avg.empty:
    top_avg = thematic_avg.head(8).sort_values("Avg award (EUR)")
    top_avg["label"] = top_avg["Thematic Priority"] + " (n=" + top_avg["Participation"].astype(int).astype(str) + ")"
    fig = px.bar(top_avg, x="Avg award (EUR)", y="label", orientation="h", color_discrete_sequence=["#C00000"])
    fig.update_layout(margin=dict(t=10), yaxis_title="", xaxis_title="Average EU contribution per project (EUR)")
    fig.update_yaxes(automargin=True)
    st.plotly_chart(fig, use_container_width=True)
    st.caption("Restricted to priorities with at least 3 projects so a single large grant doesn't distort the average.")

st.divider()
st.caption("Right-click any chart to save as an image for slides.")
