import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
from io import StringIO
import uuid
"""
EU Research Funding Pitch Dashboard
------------------------------------
Upload the "DataKey.xlsx" export (EU Funding & Tenders Portal organisation
dashboard) and get a set of pitch-ready visuals for slides.

Run locally:   streamlit run eur_funding_dashboard.py
Deploy:        push to GitHub, then deploy on streamlit.io Community Cloud.
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import re

st.set_page_config(page_title="EU Research Funding Dashboard", layout="wide")

st.markdown("""
    <style>
        .main > div {max-width: 1150px; margin: 0 auto;}
        h1, h2, h3 {color: #1F3864;}
        .stMetric {background:#f5f7fa; padding:10px; border-radius:10px;}
    </style>
""", unsafe_allow_html=True)

st.title("EU Research Funding Dashboard")
st.caption("Upload the DataKey.xlsx export from the EU Funding & Tenders Portal "
           "organisation dashboard to generate slide-ready visuals.")

uploaded = st.file_uploader("Upload DataKey.xlsx", type=["xlsx"])

if not uploaded:
    st.info("Waiting for a file. Sheet order/naming can vary — this app "
            "detects each table by its column headers, not by sheet name.")
    st.stop()

# ---------- Load every sheet ----------
raw = pd.read_excel(uploaded, sheet_name=None, engine="openpyxl")


def find(cols_required):
    """Return the first sheet whose columns are a superset of cols_required."""
    for df in raw.values():
        if all(any(c.lower() == col.lower() for col in df.columns) for c in cols_required):
            return df.copy()
    return None


def get_col(df, name):
    for c in df.columns:
        if c.lower() == name.lower():
            return c
    return None


# ---------- Locate each known table ----------
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

# ---------- KPI row ----------
st.header("Headline numbers")
c1, c2, c3, c4, c5 = st.columns(5)
if df_contrib_total is not None:
    c1.metric("Net EU Contribution", f"€{df_contrib_total.iloc[0,0]/1e6:.1f}M")
if df_partic_total is not None:
    c2.metric("Participations", f"{int(df_partic_total.iloc[0,0])}")
if df_grants_total is not None:
    c3.metric("Signed Grants", f"{int(df_grants_total.iloc[0,0])}")
if df_rank is not None:
    c4.metric("National Rank", df_rank.iloc[0, 0])
if df_erc is not None:
    c5.metric("ERC Principal Investigators", int(df_erc.iloc[0, 0]))

# ---------- Leadership: Coordinator vs Participant ----------
st.header("1. Leadership in projects")
if df_role is not None:
    role_col, part_col = df_role.columns[0], df_role.columns[1]
    total = df_role[part_col].sum()
    coord = df_role.loc[df_role[role_col].str.upper() == "COORDINATOR", part_col].sum()
    pct = coord / total * 100
    colA, colB = st.columns([1, 2])
    colA.metric("Share as Coordinator (not just participant)", f"{pct:.1f}%")
    fig = px.pie(df_role, names=role_col, values=part_col, hole=0.55,
                 color_discrete_sequence=["#1F3864", "#8FAADC"])
    fig.update_layout(margin=dict(t=10, b=10))
    colB.plotly_chart(fig, use_container_width=True)

# ---------- Growth trajectory ----------
st.header("2. Growth trajectory")
if df_years is not None:
    year_col, part_col, cum_col = df_years.columns[0], df_years.columns[1], df_years.columns[2]
    fig = go.Figure()
    fig.add_bar(x=df_years[year_col], y=df_years[part_col], name="New participations",
                marker_color="#8FAADC")
    fig.add_scatter(x=df_years[year_col], y=df_years[cum_col], name="Cumulative",
                     line=dict(color="#C00000", width=3), yaxis="y2")
    fig.update_layout(
        yaxis=dict(title="New participations per year"),
        yaxis2=dict(title="Cumulative", overlaying="y", side="right"),
        legend=dict(orientation="h", y=1.1), margin=dict(t=30)
    )
    st.plotly_chart(fig, use_container_width=True)
    recent = df_years[df_years[year_col] >= 2022][part_col].sum()
    st.caption(f"**{recent}** of {int(df_years[cum_col].iloc[-1])} total participations "
               f"(2022 onward) came in the last five years — momentum is accelerating.")

# ---------- Framework programme comparison ----------
st.header("3. Framework programme evolution")
if df_fp is not None:
    fp_col, part_col = df_fp.columns[0], df_fp.columns[1]
    fig = px.bar(df_fp, x=fp_col, y=part_col, text=part_col,
                 color=fp_col, color_discrete_sequence=px.colors.sequential.Blues_r[:3])
    fig.update_layout(showlegend=False, margin=dict(t=10))
    st.plotly_chart(fig, use_container_width=True)

# ---------- Pillar funding + avg award size (efficiency insight) ----------
st.header("4. Where the money comes from — and average award size")
if df_pillar is not None and df_thematic is not None:
    pil_col, val_col = df_pillar.columns[0], df_pillar.columns[1]
    top_pillar = df_pillar.nlargest(8, val_col)
    col1, col2 = st.columns(2)
    fig1 = px.bar(top_pillar, x=val_col, y=pil_col, orientation="h",
                  color_discrete_sequence=["#1F3864"])
    fig1.update_layout(yaxis=dict(autorange="reversed"), margin=dict(t=10),
                        xaxis_title="Net EU Contribution (EUR)", yaxis_title="")
    col1.plotly_chart(fig1, use_container_width=True)

    th = df_thematic.copy()
    th.columns = ["Thematic Priority", "Participation", "Net EU Contribution (EUR)", "Participant Cost (EUR)"]
    th = th[th["Participation"] > 0]
    th["Avg award (EUR)"] = th["Net EU Contribution (EUR)"] / th["Participation"]
    top_avg = th.nlargest(8, "Avg award (EUR)")
    fig2 = px.bar(top_avg, x="Avg award (EUR)", y="Thematic Priority", orientation="h",
                  color_discrete_sequence=["#C00000"])
    fig2.update_layout(yaxis=dict(autorange="reversed"), margin=dict(t=10), yaxis_title="")
    col2.plotly_chart(fig2, use_container_width=True)
    st.caption("Left: total funding by pillar. Right: which thematic priorities carry "
               "the biggest average grant per project (funding concentration, not just volume).")

# ---------- Department / Faculty cleanup ----------
st.header("5. Funding by faculty (cleaned)")
if df_dept is not None:
    name_col = get_col(df_dept, "Department Name")
    val_col = get_col(df_dept, "Pro rata Department EU Contribution (EUR)")
    d = df_dept[[name_col, val_col]].dropna()
    d = d[d[name_col].str.lower() != "totals"]

    CANONICAL = [
        (r"international institute of social studies|\biss\b", "International Institute of Social Studies (ISS)"),
        (r"health policy|ibmg|eshpm|health economics|health technology assessment|medical technology assessment|hta", "Erasmus School of Health Policy & Management (ESHPM)"),
        (r"social and behav|faculty of social sciences|public administration|sociology|psychology, education", "Erasmus School of Social and Behavioural Sciences (ESSB)"),
        (r"rotterdam school of management|\brsm\b", "Rotterdam School of Management (RSM)"),
        (r"school of economics|econometric|business economics", "Erasmus School of Economics (ESE)"),
        (r"school of law|criminology", "Erasmus School of Law (ESL)"),
        (r"history, culture and communication|media and communication|arts and culture", "Erasmus School of History, Culture and Communication (ESHCC)"),
        (r"philosophy", "Erasmus School of Philosophy (ESPhil)"),
        (r"drift|transitions", "DRIFT"),
        (r"data analytics|research services|university library", "Central research support"),
    ]

    def canon(name):
        n = name.lower()
        for pat, label in CANONICAL:
            if re.search(pat, n):
                return label
        return "Other / central units"

    d["Faculty"] = d[name_col].apply(canon)
    agg = d.groupby("Faculty", as_index=False)[val_col].sum().sort_values(val_col, ascending=False)
    fig = px.bar(agg, x=val_col, y="Faculty", orientation="h", color_discrete_sequence=["#1F3864"])
    fig.update_layout(yaxis=dict(autorange="reversed"), margin=dict(t=10),
                       xaxis_title="Pro-rata EU Contribution (EUR)", yaxis_title="")
    st.plotly_chart(fig, use_container_width=True)
    st.caption(f"Consolidated {d[name_col].nunique()} raw department-name variants into "
               f"{agg.shape[0]} recognisable faculties.")

# ---------- Top collaboration partners ----------
st.header("6. Closest research partners across Europe")
if df_collab is not None:
    org_col, link_col = df_collab.columns[0], df_collab.columns[1]
    top_collab = df_collab.nlargest(15, link_col).sort_values(link_col)
    fig = px.bar(top_collab, x=link_col, y=org_col, orientation="h",
                 color_discrete_sequence=["#2E75B6"])
    fig.update_layout(margin=dict(t=10), xaxis_title="Shared project links", yaxis_title="")
    st.plotly_chart(fig, use_container_width=True)
    st.caption(f"Out of {df_collab.shape[0]} distinct partner organisations across all projects.")

# ---------- Research identity via keywords ----------
st.header("7. Research identity — recurring topics")
if df_keywords is not None:
    kw_col, n_col = df_keywords.columns[0], df_keywords.columns[1]
    top_kw = df_keywords.nlargest(15, n_col).sort_values(n_col)
    fig = px.bar(top_kw, x=n_col, y=kw_col, orientation="h",
                 color_discrete_sequence=["#548235"])
    fig.update_layout(margin=dict(t=10), xaxis_title="Number of projects tagged", yaxis_title="")
    st.plotly_chart(fig, use_container_width=True)

# ---------- Project status & MSCA/EIC context ----------
st.header("8. Portfolio status")
colA, colB, colC = st.columns(3)
if df_status is not None:
    st_col, cnt_col = df_status.columns[0], df_status.columns[1]
    fig = px.pie(df_status, names=st_col, values=cnt_col, hole=0.55,
                 color_discrete_sequence=["#C00000", "#2E75B6"])
    fig.update_layout(margin=dict(t=10, b=10), showlegend=True)
    colA.plotly_chart(fig, use_container_width=True)
if df_msca is not None:
    colB.metric("MSCA Participation", int(df_msca.iloc[0, 0]))
if df_eic is not None:
    colC.metric("EIC Participation", int(df_eic.iloc[0, 0]))

st.divider()
st.caption("Built for slide export — right-click any chart to save as an image, "
           "or use Streamlit's camera icon in the top-right of each chart.")