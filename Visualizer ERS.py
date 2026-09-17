"""EU Research Funding Pitch Dashboard."""

import re
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit.components.v1 as components

st.set_page_config(page_title="EU Research Funding Dashboard", layout="wide")

st.markdown("""
<style>
.main > div {max-width: 1150px; margin: 0 auto;}
h1, h2, h3 {color: #1F3864;}
.stat-card {background:#f5f7fa; border-radius:10px; padding:14px 10px; text-align:center; height:100%;}
.stat-label {font-size:0.78rem; color:#555; margin-bottom:4px;}
.stat-value {font-size:1.55rem; font-weight:700; color:#1F3864; line-height:1.25; word-wrap:break-word; white-space:normal;}
.stat-sub {font-size:0.72rem; color:#888; margin-top:2px;}
.stPlotlyChart, .js-plotly-plot, .plotly {width: 100% !important;}
</style>
""", unsafe_allow_html=True)


def stat_card(col, label, value, sub=None):
    """Custom card, not st.metric — st.metric clips long values/labels with an
    ellipsis and doesn't wrap, which is exactly the bug we're avoiding here."""
    col.markdown(f"""<div class="stat-card">
        <div class="stat-label">{label}</div>
        <div class="stat-value">{value}</div>
        {f'<div class="stat-sub">{sub}</div>' if sub else ''}
    </div>""", unsafe_allow_html=True)


st.title("🎓 EU Research Funding — Pitch Dashboard")
st.caption("Upload the DataKey.xlsx export to generate slide-ready visuals.")
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

# ================= HEADLINE KPIs =================
st.header("Headline numbers")
row1 = st.columns(4)
if df_contrib_total is not None:
    stat_card(row1[0], "Net EU Contribution", f"€{df_contrib_total.iloc[0, 0] / 1e6:.1f}M")
if df_partic_total is not None:
    stat_card(row1[1], "Participations", int(df_partic_total.iloc[0, 0]))
if df_grants_total is not None:
    stat_card(row1[2], "Signed Grants", int(df_grants_total.iloc[0, 0]))
if df_rank is not None:
    stat_card(row1[3], "National Rank", df_rank.iloc[0, 0])

row2 = st.columns(3)
if df_erc is not None:
    stat_card(row2[0], "ERC Principal Investigators", int(df_erc.iloc[0, 0]))
if df_msca is not None:
    stat_card(row2[1], "MSCA Participation", int(df_msca.iloc[0, 0]))
if df_eic is not None:
    eic_val = int(df_eic.iloc[0, 0])
    if eic_val > 0:
        stat_card(row2[2], "EIC Participation", eic_val)
    else:
        row2[2].caption("No EIC-funded projects yet.")

st.divider()
st.header("A — Recreating the organisation dashboard")
st.caption("Matches the panels in the reference PDF, using the current export.")

if df_org is not None:
    st.subheader("Organisation details")
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

if df_years is not None:
    st.subheader("Evolution of participation")
    year_col, part_col, cum_col = df_years.columns[:3]
    fig = go.Figure()
    fig.add_bar(x=df_years[year_col], y=df_years[part_col], name="New participations", marker_color="#8FAADC")
    fig.add_scatter(x=df_years[year_col], y=df_years[cum_col], name="Cumulative",
                    line=dict(color="#C00000", width=3), yaxis="y2")
    fig.update_layout(yaxis=dict(title="New per year"),
                      yaxis2=dict(title="Cumulative", overlaying="y", side="right"),
                      legend=dict(orientation="h", y=1.15), margin=dict(t=30))
    st.plotly_chart(fig, use_container_width=True)
    recent = df_years[df_years[year_col] >= 2022][part_col].sum()
    st.caption(f"{int(recent)} of {int(df_years[cum_col].iloc[-1])} total participations came from 2022 onward.")

if df_fp is not None:
    st.subheader("Participation per Framework Programme")
    fp_col, part_col = df_fp.columns[:2]
    fig = px.bar(df_fp, x=fp_col, y=part_col, text=part_col, color=fp_col,
                 color_discrete_sequence=px.colors.sequential.Blues_r[:3])
    fig.update_traces(textposition="outside", cliponaxis=False)
    fig.update_layout(showlegend=False, margin=dict(t=10))
    fig.update_yaxes(automargin=True)
    st.plotly_chart(fig, use_container_width=True)

if df_pillar is not None:
    st.subheader("Net EU contribution by pillar")
    pil_col, val_col = df_pillar.columns[:2]
    pil = df_pillar.sort_values(val_col)
    fig = px.bar(pil, x=val_col, y=pil_col, orientation="h", color_discrete_sequence=["#1F3864"])
    fig.update_layout(margin=dict(t=10), xaxis_title="Net EU Contribution (EUR)", yaxis_title="")
    fig.update_yaxes(automargin=True)
    st.plotly_chart(fig, use_container_width=True)
    st.caption("Pillar names are shown as exported; categories from different framework programmes are not merged.")

if df_dept is not None:
    st.subheader("Funding by faculty (cleaned)")
    name_col = get_col(df_dept, "Department Name")
    val_col = get_col(df_dept, "Pro rata Department EU Contribution (EUR)")
    d = df_dept[[name_col, val_col]].dropna()
    d = d[d[name_col].str.lower() != "totals"]
    canonical = [
        (r"international institute (of|for) social studies|\biss\b|institute (of|for) social studies", "International Institute of Social Studies (ISS)"),
        (r"health polic|ibmg|eshpm|health econom|health technology assessment|medical technology assessment|\bhta\b", "Erasmus School of Health Policy & Management (ESHPM)"),
        (r"social and behav|faculty of social sciences|public admin|sociology|psychology,? education|essb", "Erasmus School of Social and Behavioural Sciences (ESSB)"),
        (r"rotterdam school of manage|\brsm\b|faculty of management", "Rotterdam School of Management (RSM)"),
        (r"school of economics|econometric|business economics|\bese\b", "Erasmus School of Economics (ESE)"),
        (r"school of law|criminolog|\besl\b", "Erasmus School of Law (ESL)"),
        (r"history,?\s*(culture|art)|eshcc|arts? and culture|media and communication", "Erasmus School of History, Culture and Communication (ESHCC)"),
        (r"philosoph|esphil", "Erasmus School of Philosophy (ESPhil)"),
        (r"drift|transitions", "DRIFT"),
        (r"data analytics|research services|university library|executive board|\bcvb\b", "Central research support / governance"),
    ]

    def canon(name):
        for pattern, label in canonical:
            if re.search(pattern, name.lower()):
                return label
        return "Unmatched (see audit below)"

    d["Faculty"] = d[name_col].apply(canon)
    matched = d[d["Faculty"] != "Unmatched (see audit below)"]
    unmatched = d[d["Faculty"] == "Unmatched (see audit below)"]
    agg = matched.groupby("Faculty", as_index=False)[val_col].sum().sort_values(val_col)
    fig = px.bar(agg, x=val_col, y="Faculty", orientation="h", color_discrete_sequence=["#1F3864"])
    fig.update_layout(margin=dict(t=10), xaxis_title="Pro-rata EU Contribution (EUR)", yaxis_title="")
    fig.update_yaxes(automargin=True)
    st.plotly_chart(fig, use_container_width=True)
    st.caption(f"Consolidated {matched[name_col].nunique()} raw name variants into {agg.shape[0]} faculties. "
               f"€{unmatched[val_col].sum():,.0f} across {unmatched.shape[0]} unmatched rows is excluded.")
    if not unmatched.empty:
        with st.expander("Audit: unmatched department names (not shown above)"):
            st.dataframe(unmatched.sort_values(val_col, ascending=False), use_container_width=True)

if df_keywords is not None:
    st.subheader("Project keywords")
    kw_col, n_col = df_keywords.columns[:2]
    top_kw = df_keywords.nlargest(15, n_col).sort_values(n_col)
    fig = px.bar(top_kw, x=n_col, y=kw_col, orientation="h", color_discrete_sequence=["#548235"])
    fig.update_layout(margin=dict(t=10), xaxis_title="Number of projects tagged", yaxis_title="")
    fig.update_yaxes(automargin=True)
    st.plotly_chart(fig, use_container_width=True)

if df_collab is not None:
    st.subheader("Top collaboration partners")
    org_col, link_col = df_collab.columns[:2]
    top_collab = df_collab.nlargest(15, link_col).sort_values(link_col)
    fig = px.bar(top_collab, x=link_col, y=org_col, orientation="h", color_discrete_sequence=["#2E75B6"])
    fig.update_layout(margin=dict(t=10), xaxis_title="Shared project links", yaxis_title="")
    fig.update_yaxes(automargin=True)
    st.plotly_chart(fig, use_container_width=True)
    st.caption(f"Top 15 of {df_collab.shape[0]} distinct partner organisations.")

st.subheader("Collaboration map")
components.html("""
<div style="width:100%;height:520px;">
<iframe src="https://dashboard.tech.ec.europa.eu/qs_digit_dashboard_mt/public/single/?appid=dc5f6f40-c9de-4c40-8648-015d6ff21342&obj=EVcQAd&theme=card&opt=ctxmenu,currsel&select=$::Signature%20Year,2007,2008,2009,2010,2011,2012,2013,2014,2015,2016,2017,2018,2019,2020,2021,2022,2023,2024,2025,2026&select=$::Organisation%20Name,ERASMUS%20UNIVERSITEIT%20ROTTERDAM"
style="border:none;width:100%;height:100%;"></iframe></div>
""", height=530)
st.caption("If this shows blank instead of a map, the EC's dashboard is likely blocking embeds from this domain "
           "(a server-side X-Frame-Options/CSP restriction, not something fixable from this app). "
           "Fallback link: https://dashboard.tech.ec.europa.eu/qs_digit_dashboard_mt/public/sense/app/dc5f6f40-c9de-4c40-8648-015d6ff21342")

if df_keyfig is not None:
    st.subheader("Key figures — EUR vs national totals, by Framework Programme")
    rows = []
    for _, r in df_keyfig.iterrows():
        indicator = r["Indicator"]
        is_eur_amt = "(EUR)" in indicator
        entry = {"Indicator": indicator}
        for prog in ["HE", "H2020", "FP7"]:
            org = eu_num(r.get(f"Organisation in {prog}"))
            pct = r.get(f"% in {prog}")
            if org is None:
                entry[prog] = "—"
            else:
                amount = f"€{org:,.0f}" if is_eur_amt else f"{org:,.0f}"
                entry[prog] = f"{amount} ({pct} of NL total)" if isinstance(pct, str) and pct != "-" else amount
        rows.append(entry)
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

with st.expander("Not included — needs data this export doesn't contain"):
    st.markdown("""
- **Project-level list** — this export contains pre-aggregated summaries rather than one row per project.
- **Contribution vs. Total Cost scatter** — this needs per-project cost and funding figures.
""")

st.divider()
st.header("B — Two things not in the source dashboard")
st.caption("Genuinely new, derived from the raw figures — not just a reformat of an existing PDF panel.")

if df_thematic is not None:
    st.subheader("Average award size by thematic priority")
    th = df_thematic.copy()
    th.columns = ["Thematic Priority", "Participation", "Net EU Contribution (EUR)", "Participant Cost (EUR)"]
    th = th[th["Participation"] >= 3]
    th["Avg award (EUR)"] = th["Net EU Contribution (EUR)"] / th["Participation"]
    top_avg = th.nlargest(8, "Avg award (EUR)").sort_values("Avg award (EUR)")
    top_avg["label"] = top_avg["Thematic Priority"] + " (n=" + top_avg["Participation"].astype(int).astype(str) + ")"
    fig = px.bar(top_avg, x="Avg award (EUR)", y="label", orientation="h", color_discrete_sequence=["#C00000"])
    fig.update_layout(margin=dict(t=10), yaxis_title="", xaxis_title="Average EU contribution per project (EUR)")
    fig.update_yaxes(automargin=True)
    st.plotly_chart(fig, use_container_width=True)
    st.caption("Contribution ÷ participation, restricted to priorities with at least 3 projects.")

st.caption("The cleaned faculty rollup in Section A is the other genuinely new piece of analysis.")
st.divider()
st.caption("Right-click any chart to save as an image for slides.")
