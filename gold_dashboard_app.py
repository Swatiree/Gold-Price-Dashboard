import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import os

st.set_page_config(page_title="Gold Price Dashboard", page_icon="🥇", layout="wide")

st.markdown("""
<style>
    body { background-color: #0e1117; }
    .metric-card {
        background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
        border: 1px solid #f4c430; border-radius: 12px;
        padding: 16px 20px; text-align: center;
    }
    .metric-label { font-size: 13px; color: #aaa; margin-bottom: 4px; }
    .metric-value { font-size: 26px; font-weight: 700; color: #f4c430; }
    .metric-sub   { font-size: 12px; color: #888; margin-top: 2px; }
    .section-title {
        font-size: 20px; font-weight: 700; color: #f4c430;
        border-bottom: 1px solid #333; padding-bottom: 6px; margin-bottom: 14px;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("## 🥇 Gold Price Analytics Dashboard")
st.markdown("*Daily data from August 2000 – March 2026 · 6,399 trading sessions*")

# ─── Locate Excel File ────────────────────────────────────────────────────────
EXCEL_FILENAME = "gold_price_edited2.xlsx"
script_dir = os.path.dirname(os.path.abspath(__file__))
auto_path = os.path.join(script_dir, EXCEL_FILENAME)

if os.path.exists(auto_path):
    file_source = auto_path
else:
    st.warning(f"⚠️ **{EXCEL_FILENAME}** not found next to the script. Please upload it below.")
    uploaded = st.file_uploader("Upload gold_price_edited2.xlsx", type=["xlsx"])
    if uploaded is None:
        st.info("👆 Upload the Excel file to load the dashboard.")
        st.stop()
    file_source = uploaded

# ─── Load Data ────────────────────────────────────────────────────────────────
@st.cache_data
def load_data(src):
    xl = pd.ExcelFile(src)

    master = pd.read_excel(xl, sheet_name="MASTER DATA", parse_dates=["Date"])
    master["Year"] = master["Date"].dt.year

    extracted = pd.read_excel(xl, sheet_name="EXTRACTED GOLD PRICE TABLE", parse_dates=["Date"])
    extracted["Year"] = extracted["Date"].dt.year

    ma_raw = pd.read_excel(xl, sheet_name="200-DAY MA VS 50-DAY MA", header=None)
    ma = ma_raw.iloc[4:31].copy()
    ma.columns = ["Year", "50_Day_MA", "200_Day_MA"]
    ma = ma[ma["Year"] != "<30-08-2000"]
    ma["Year"] = pd.to_numeric(ma["Year"], errors="coerce")
    ma["50_Day_MA"] = pd.to_numeric(ma["50_Day_MA"], errors="coerce")
    ma["200_Day_MA"] = pd.to_numeric(ma["200_Day_MA"], errors="coerce")
    ma = ma.dropna()

    avg_close_raw = pd.read_excel(xl, sheet_name="Average Close Price", header=None)
    avg_close = avg_close_raw.iloc[3:30].copy()
    avg_close.columns = ["Year", "Avg_Close"]
    avg_close = avg_close[avg_close["Year"] != "<30-08-2000"]
    avg_close["Year"] = pd.to_numeric(avg_close["Year"], errors="coerce")
    avg_close["Avg_Close"] = pd.to_numeric(avg_close["Avg_Close"], errors="coerce")
    avg_close = avg_close.dropna()

    adr_raw = pd.read_excel(xl, sheet_name="Avg. Daily Return", header=None)
    adr = adr_raw.iloc[3:30].copy()
    adr.columns = ["Year", "Avg_Daily_Return"]
    adr = adr[adr["Year"] != "<30-08-2000"]
    adr["Year"] = pd.to_numeric(adr["Year"], errors="coerce")
    adr["Avg_Daily_Return"] = pd.to_numeric(adr["Avg_Daily_Return"], errors="coerce")
    adr = adr.dropna()

    dr_raw = pd.read_excel(xl, sheet_name="Daily Return", header=None)
    dr = dr_raw.iloc[3:10].copy()
    dr.columns = ["Range", "Count"]
    dr["Count"] = pd.to_numeric(dr["Count"], errors="coerce")
    dr = dr.dropna()

    vol = pd.read_excel(xl, sheet_name="year wise volume")
    vol.columns = ["Year", "Trading_Days"]
    vol = vol[vol["Year"] != "Grand Total"]
    vol["Year"] = pd.to_numeric(vol["Year"], errors="coerce")
    vol = vol.dropna()

    oc_raw = pd.read_excel(xl, sheet_name=" open vs close Price", header=None)
    oc = oc_raw.iloc[3:30].copy()
    oc.columns = ["Year", "Avg_Open", "Avg_Close"]
    oc = oc[oc["Year"] != "<30-08-2000"]
    oc["Year"] = pd.to_numeric(oc["Year"], errors="coerce")
    oc["Avg_Open"] = pd.to_numeric(oc["Avg_Open"], errors="coerce")
    oc["Avg_Close"] = pd.to_numeric(oc["Avg_Close"], errors="coerce")
    oc = oc.dropna()

    return master, extracted, ma, avg_close, adr, dr, vol, oc

master, extracted, ma, avg_close, adr, dr, vol, oc = load_data(file_source)

GOLD = "#f4c430"
BG   = "#0e1117"

# ─── KPI Metrics ──────────────────────────────────────────────────────────────
high_row = master.loc[master["Close"].idxmax()]
low_row  = master.loc[master["Close"].idxmin()]
latest   = master.sort_values("Date").iloc[-1]
avg_ret  = extracted["Daily Return"].mean()

kpi_data = [
    ("All-Time High",    f"${high_row['Close']:,.0f}",  high_row["Date"].strftime("%b %Y")),
    ("All-Time Low",     f"${low_row['Close']:,.0f}",   low_row["Date"].strftime("%b %Y")),
    ("Latest Close",     f"${latest['Close']:,.0f}",    latest["Date"].strftime("%b %d, %Y")),
    ("Avg Daily Return", f"{avg_ret*100:.4f}%",         "2000–2026"),
    ("Total Growth",     f"{((latest['Close']-master['Close'].iloc[0])/master['Close'].iloc[0]*100):.0f}%", "Since Aug 2000"),
    ("Bullish Days",     f"{(extracted['Daily Return']>0).sum():,}", "of 6,399 sessions"),
]

cols = st.columns(6)
for col, (label, value, sub) in zip(cols, kpi_data):
    col.markdown(f"""
    <div class="metric-card">
      <div class="metric-label">{label}</div>
      <div class="metric-value">{value}</div>
      <div class="metric-sub">{sub}</div>
    </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

tabs = st.tabs(["📈 Price History", "📊 Moving Averages", "🔁 Returns Analysis", "📦 Volume", "🔍 Data Explorer"])

with tabs[0]:
    st.markdown('<div class="section-title">Gold Close Price (2000–2026)</div>', unsafe_allow_html=True)
    year_range = st.slider("Filter by Year", 2000, 2026, (2000, 2026), key="yr1")
    filtered = master[(master["Year"] >= year_range[0]) & (master["Year"] <= year_range[1])]
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=filtered["Date"], y=filtered["Close"], mode="lines",
                             name="Close Price", line=dict(color=GOLD, width=1.5),
                             fill="tozeroy", fillcolor="rgba(244,196,48,0.07)"))
    fig.update_layout(paper_bgcolor=BG, plot_bgcolor=BG, font_color="#ccc", height=420,
                      xaxis=dict(showgrid=False, color="#666"),
                      yaxis=dict(showgrid=True, gridcolor="#222", tickprefix="$"),
                      margin=dict(l=10, r=10, t=20, b=10), hovermode="x unified")
    st.plotly_chart(fig, use_container_width=True)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown('<div class="section-title">Average Annual Close Price</div>', unsafe_allow_html=True)
        fig2 = px.bar(avg_close, x="Year", y="Avg_Close", color_discrete_sequence=[GOLD])
        fig2.update_layout(paper_bgcolor=BG, plot_bgcolor=BG, font_color="#ccc",
                           height=300, yaxis_tickprefix="$", xaxis_title="",
                           margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(fig2, use_container_width=True)
    with col2:
        st.markdown('<div class="section-title">Open vs Close Price (Yearly Avg)</div>', unsafe_allow_html=True)
        fig3 = go.Figure()
        fig3.add_trace(go.Scatter(x=oc["Year"], y=oc["Avg_Open"], mode="lines+markers",
                                  name="Avg Open", line=dict(color="#4fc3f7")))
        fig3.add_trace(go.Scatter(x=oc["Year"], y=oc["Avg_Close"], mode="lines+markers",
                                  name="Avg Close", line=dict(color=GOLD)))
        fig3.update_layout(paper_bgcolor=BG, plot_bgcolor=BG, font_color="#ccc",
                           height=300, yaxis_tickprefix="$",
                           margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(fig3, use_container_width=True)

with tabs[1]:
    st.markdown('<div class="section-title">50-Day vs 200-Day Moving Average (Yearly Avg)</div>', unsafe_allow_html=True)
    fig4 = go.Figure()
    fig4.add_trace(go.Scatter(x=ma["Year"], y=ma["50_Day_MA"], mode="lines+markers",
                              name="50-Day MA", line=dict(color="#ff7043", width=2)))
    fig4.add_trace(go.Scatter(x=ma["Year"], y=ma["200_Day_MA"], mode="lines+markers",
                              name="200-Day MA", line=dict(color=GOLD, width=2)))
    fig4.update_layout(paper_bgcolor=BG, plot_bgcolor=BG, font_color="#ccc", height=420,
                       yaxis_tickprefix="$", xaxis=dict(showgrid=False),
                       yaxis=dict(gridcolor="#222"),
                       margin=dict(l=10, r=10, t=20, b=10), hovermode="x unified")
    st.plotly_chart(fig4, use_container_width=True)

    st.markdown('<div class="section-title">MA Spread (200-Day − 50-Day)</div>', unsafe_allow_html=True)
    ma["Spread"] = ma["200_Day_MA"] - ma["50_Day_MA"]
    fig5 = px.bar(ma, x="Year", y="Spread", color="Spread",
                  color_continuous_scale=["#ff5252", "#333", GOLD])
    fig5.update_layout(paper_bgcolor=BG, plot_bgcolor=BG, font_color="#ccc",
                       height=280, yaxis_tickprefix="$",
                       margin=dict(l=10, r=10, t=10, b=10), coloraxis_showscale=False)
    st.plotly_chart(fig5, use_container_width=True)
    st.info("💡 **Golden Cross**: 50-Day MA crosses above 200-Day MA = bullish. **Death Cross** = opposite.")

with tabs[2]:
    col1, col2 = st.columns([3, 2])
    with col1:
        st.markdown('<div class="section-title">Daily Return Distribution</div>', unsafe_allow_html=True)
        order = ["< -5%", "-5% to -2%", "-2% to -0.5%", "-0.5% to 0%",
                 "0% to 0.5%", "0.5% to 2%", "> 2%"]
        dr_sorted = dr.set_index("Range").reindex(order).reset_index().dropna()
        colors = ["#d32f2f","#e53935","#ef9a9a","#ffcdd2","#a5d6a7","#66bb6a","#2e7d32"]
        fig6 = px.bar(dr_sorted, x="Range", y="Count", color="Range",
                      color_discrete_sequence=colors, text="Count")
        fig6.update_traces(textfont_size=11)
        fig6.update_layout(paper_bgcolor=BG, plot_bgcolor=BG, font_color="#ccc",
                           height=380, showlegend=False, margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(fig6, use_container_width=True)
    with col2:
        st.markdown('<div class="section-title">Return Buckets</div>', unsafe_allow_html=True)
        fig7 = px.pie(dr_sorted, names="Range", values="Count",
                      color_discrete_sequence=px.colors.diverging.RdYlGn)
        fig7.update_layout(paper_bgcolor=BG, font_color="#ccc",
                           height=380, margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(fig7, use_container_width=True)

    st.markdown('<div class="section-title">Average Daily Return by Year</div>', unsafe_allow_html=True)
    adr["Color"] = adr["Avg_Daily_Return"].apply(lambda x: GOLD if x >= 0 else "#ef5350")
    fig8 = go.Figure(go.Bar(x=adr["Year"], y=adr["Avg_Daily_Return"] * 100,
                            marker_color=adr["Color"],
                            text=(adr["Avg_Daily_Return"] * 100).round(3).astype(str) + "%",
                            textfont=dict(size=9)))
    fig8.add_hline(y=0, line_color="#555", line_dash="dot")
    fig8.update_layout(paper_bgcolor=BG, plot_bgcolor=BG, font_color="#ccc",
                       height=320, yaxis_ticksuffix="%",
                       margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig8, use_container_width=True)

with tabs[3]:
    st.markdown('<div class="section-title">Trading Days per Year</div>', unsafe_allow_html=True)
    vol_plot = vol[vol["Year"] < 2026]
    fig9 = px.bar(vol_plot, x="Year", y="Trading_Days",
                  color_discrete_sequence=[GOLD], text="Trading_Days")
    fig9.update_traces(textfont_size=9, textangle=0)
    fig9.update_layout(paper_bgcolor=BG, plot_bgcolor=BG, font_color="#ccc",
                       height=400, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig9, use_container_width=True)
    st.info("📌 2026 shows only 41 trading days as data runs to March 3, 2026.")

with tabs[4]:
    st.markdown('<div class="section-title">Raw Master Data Explorer</div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        yr = st.slider("Year range", 2000, 2026, (2020, 2026), key="yr_exp")
    with col2:
        search_col = st.selectbox("Sort by", ["Date", "Close", "Volume", "High", "Low"])
    exp = master[(master["Year"] >= yr[0]) & (master["Year"] <= yr[1])].sort_values(search_col, ascending=False)
    st.dataframe(exp[["Date", "Open", "High", "Low", "Close", "Volume"]].reset_index(drop=True),
                 use_container_width=True, height=420)
    st.download_button("⬇️ Download Filtered Data as CSV", data=exp.to_csv(index=False),
                       file_name="gold_filtered.csv", mime="text/csv")

st.markdown("---")
st.markdown("<center style='color:#555;font-size:12px'>Gold Price Analytics · Data covers Aug 2000 – Mar 2026 · Built with Streamlit & Plotly</center>",
            unsafe_allow_html=True)
