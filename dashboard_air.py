"""
Dashboard Analisis Konsumsi Air
================================
Cara hosting di Streamlit Cloud:
1. Upload dashboard_air.py + data_air.xlsx ke GitHub
2. Connect repo di share.streamlit.io

Cara jalankan lokal:
    pip install streamlit plotly pandas openpyxl
    streamlit run dashboard_air.py
"""

import os
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Dashboard Konsumsi Air",
    page_icon="💧",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

/* Global */
html, body { font-family: 'Inter', sans-serif; }
[data-testid="stAppViewContainer"] { background: #0A1F3D; }

/* ── Area utama: SEMUA teks terang di atas bg gelap ── */
[data-testid="stMain"] * {
    color: #E8F1FF !important;
}
[data-testid="stMain"] [role="tab"][aria-selected="true"] {
    color: #4EA8DE !important;
    border-bottom: 2px solid #4EA8DE !important;
}
[data-testid="stMain"] [role="tab"] { color: #8AABDF !important; }
[data-testid="stMain"] [data-testid="stCaptionContainer"] * { color: #8AABDF !important; }
[data-testid="stMain"] a { color: #4EA8DE !important; }

/* Sidebar */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0A1F3D 0%, #0D2B55 100%);
}
[data-testid="stSidebar"] *                              { color: #C8DCFF !important; }
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3                            { color: #FFFFFF !important; }
[data-testid="stSidebar"] .stSelectbox > div > div,
[data-testid="stSidebar"] .stMultiSelect > div > div {
    background: #162F5A !important;
    border-color: #2A4A80 !important;
    color: #E8F1FF !important;
}
[data-testid="stSidebar"] .stSelectbox svg,
[data-testid="stSidebar"] .stMultiSelect svg             { fill: #8AABDF !important; }
[data-testid="stSidebar"] hr                             { border-color: #1E3A6E; }
[data-testid="stSidebar"] [data-baseweb="tag"]           { background: #1A4A8A !important; }
[data-testid="stSidebar"] [data-baseweb="tag"] span      { color: #FFFFFF !important; }
[data-testid="stSidebar"] [data-testid="stDateInput"] input {
    background: #162F5A !important;
    color: #E8F1FF !important;
    border-color: #2A4A80 !important;
}

/* KPI Cards */
.kpi-grid { display: grid; grid-template-columns: repeat(4,1fr); gap: 16px; margin-bottom: 24px; }
.kpi-card {
    background: #162F5A;
    border-radius: 16px;
    padding: 20px 24px;
    box-shadow: 0 2px 16px rgba(0,0,0,0.3);
    border-top: 4px solid #4EA8DE;
    position: relative; overflow: hidden;
}
.kpi-card::after {
    content: ''; position: absolute; right: -20px; top: -20px;
    width: 80px; height: 80px; border-radius: 50%;
    background: rgba(26,111,191,0.06);
}
.kpi-label { font-size: 0.7rem; font-weight: 600; color: #A8C8FF !important;
             letter-spacing: .07em; text-transform: uppercase; margin-bottom: 6px; }
.kpi-value { font-size: 2rem; font-weight: 700; color: #FFFFFF !important; line-height: 1.1; }
.kpi-sub   { font-size: 0.75rem; color: #A8C8FF !important; margin-top: 4px; }

/* Section titles */
.sec-title {
    font-size: 0.95rem; font-weight: 700; color: #E8F1FF !important;
    border-left: 4px solid #4EA8DE; padding-left: 10px;
    margin: 28px 0 14px 0;
}

/* Insight box */
.insight {
    background: #162F5A; border-radius: 12px; padding: 14px 18px;
    font-size: 0.82rem; color: #C8DCFF !important;
    box-shadow: 0 1px 8px rgba(0,0,0,0.2);
    border-left: 3px solid #4EA8DE;
    margin-bottom: 8px;
}

#MainMenu, footer, header { visibility: hidden; }
.block-container { padding-top: 1.4rem; padding-bottom: 1rem; background: transparent; }
[data-testid="stMain"] { background: #0A1F3D; }
.js-plotly-plot .plotly { border-radius: 12px; }
</style>
""", unsafe_allow_html=True)


# ── Load data (hardcoded, no upload needed) ───────────────────────────────────
DATA_FILE = os.path.join(os.path.dirname(__file__), "data_air.xlsx")

@st.cache_data
def load_data():
    df = pd.read_excel(DATA_FILE)
    df.columns = df.columns.str.strip()
    df["Date"] = pd.to_datetime(df["Date"])
    df["Jam"]       = df["Date"].dt.hour
    df["Hari"]      = df["Date"].dt.date
    df["Minggu"]    = df["Date"].dt.to_period("W").apply(lambda r: r.start_time.date())
    df["Bulan_Num"] = df["Date"].dt.month
    df["Bulan"]     = df["Date"].dt.strftime("%B %Y")
    df["DayOfWeek"] = df["Date"].dt.dayofweek
    return df.sort_values("Date").reset_index(drop=True)

df_raw = load_data()

BULAN_ID = {1:"Januari",2:"Februari",3:"Maret",4:"April",5:"Mei",6:"Juni",
            7:"Juli",8:"Agustus",9:"September",10:"Oktober",11:"November",12:"Desember"}
HARI_LABEL = {0:"Sen",1:"Sel",2:"Rab",3:"Kam",4:"Jum",5:"Sab",6:"Min"}
COLORS = ["#1A6FBF","#E05C2A","#2CA87F","#9B5DE5","#F4B942","#E84393","#00B4D8"]

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 💧 Konsumsi Air")
    st.caption(f"Data: **{df_raw['Date'].min().strftime('%d %b %Y')}** – **{df_raw['Date'].max().strftime('%d %b %Y')}**")
    st.divider()

    bulan_list = df_raw["Bulan"].unique().tolist()
    sel_bulan  = st.multiselect("Filter Bulan", bulan_list, default=bulan_list)

    min_d, max_d = df_raw["Date"].min().date(), df_raw["Date"].max().date()
    date_range   = st.date_input("Rentang Tanggal", value=(min_d, max_d),
                                  min_value=min_d, max_value=max_d)
    st.divider()

    st.markdown("### ⚙️ Pengaturan Grafik")
    agg_mode   = st.selectbox("Agregasi", ["Per Jam","Per Hari","Per Minggu","Per Bulan"])
    chart_type = st.selectbox("Tipe Grafik", ["Line","Bar","Area"])
    compare_on = st.toggle("Bandingkan antar Bulan", value=True)
    st.divider()

    show_heatmap   = st.toggle("Heatmap Jam × Hari", value=True)
    show_dist      = st.toggle("Distribusi & Statistik", value=True)
    show_jam_bulan = st.toggle("Perbandingan Per Jam Tiap Bulan", value=True)

# ── Filter ────────────────────────────────────────────────────────────────────
d_start, d_end = (date_range[0], date_range[1]) if len(date_range) == 2 else (min_d, max_d)
df = df_raw[
    df_raw["Bulan"].isin(sel_bulan) &
    (df_raw["Date"].dt.date >= d_start) &
    (df_raw["Date"].dt.date <= d_end)
].copy()

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("<h1 style='color:#0A1F3D !important;font-family:Inter,sans-serif;font-size:2rem;font-weight:700;margin-bottom:0.5rem'>💧 Dashboard Analisis Konsumsi Air</h1>", unsafe_allow_html=True)

if df.empty:
    st.warning("Tidak ada data untuk filter yang dipilih.")
    st.stop()

# ── KPI ───────────────────────────────────────────────────────────────────────
total      = df["Consume Per Hour"].sum()
rata_jam   = df["Consume Per Hour"].mean()
maks_val   = df["Consume Per Hour"].max()
maks_tgl   = df.loc[df["Consume Per Hour"].idxmax(), "Date"].strftime("%d %b %Y %H:%M")
hari_aktif = df["Hari"].nunique()

st.markdown(f"""
<div class="kpi-grid">
  <div class="kpi-card">
    <div class="kpi-label">Total Konsumsi</div>
    <div class="kpi-value">{total:,.0f} M³</div>
    <div class="kpi-sub">{hari_aktif} hari aktif</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-label">Rata-rata / Jam</div>
    <div class="kpi-value">{rata_jam:,.1f} M³</div>
    <div class="kpi-sub">per jam</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-label">Puncak Konsumsi</div>
    <div class="kpi-value">{maks_val:,.0f} M³</div>
    <div class="kpi-sub">{maks_tgl}</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-label">Rata-rata / Hari</div>
    <div class="kpi-value">{total/hari_aktif:,.0f} M³</div>
    <div class="kpi-sub">per hari</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ── Helper chart layout ───────────────────────────────────────────────────────
def base_layout(height=380, show_legend=True):
    return dict(
        plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF",
        font=dict(family="Inter", size=12, color="#0A1F3D"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
                    font=dict(color="#0A1F3D")),
        showlegend=show_legend,
        margin=dict(l=4, r=4, t=24, b=4),
        height=height,
        xaxis=dict(showgrid=False, linecolor="#D4E0EE", tickfont=dict(color="#3A5A8A")),
        yaxis=dict(gridcolor="#EAF1FA", linecolor="#D4E0EE", tickfont=dict(color="#3A5A8A")),
    )

def make_chart(df_plot, x, y, color=None, ctype="Bar", height=380):
    kw = dict(x=x, y=y, color=color, color_discrete_sequence=COLORS,
              labels={y: "Konsumsi (M³)", x: ""})
    if ctype == "Bar":
        fig = px.bar(df_plot, **kw, barmode="group")
        fig.update_traces(marker_line_width=0)
    elif ctype == "Line":
        fig = px.line(df_plot, **kw, markers=True)
    else:
        fig = px.area(df_plot, **kw)
    fig.update_layout(**base_layout(height=height, show_legend=color is not None))
    return fig

# ── Grafik Utama ──────────────────────────────────────────────────────────────
st.markdown('<div class="sec-title">📊 Grafik Konsumsi Utama</div>', unsafe_allow_html=True)

if agg_mode == "Per Jam":
    df_agg = df.groupby(["Date","Bulan"])["Consume Per Hour"].sum().reset_index()
    x_col = "Date"
elif agg_mode == "Per Hari":
    df_agg = df.groupby(["Hari","Bulan"])["Consume Per Hour"].sum().reset_index()
    df_agg["Hari"] = pd.to_datetime(df_agg["Hari"])
    x_col = "Hari"
elif agg_mode == "Per Minggu":
    df_agg = df.groupby(["Minggu","Bulan"])["Consume Per Hour"].sum().reset_index()
    df_agg["Minggu"] = pd.to_datetime(df_agg["Minggu"])
    x_col = "Minggu"
else:
    df_agg = df.groupby(["Bulan","Bulan_Num"])["Consume Per Hour"].sum().reset_index().sort_values("Bulan_Num")
    x_col = "Bulan"

color_col = "Bulan" if (compare_on and agg_mode != "Per Bulan") else None
st.plotly_chart(make_chart(df_agg, x_col, "Consume Per Hour", color_col, chart_type, 400),
                use_container_width=True)

# ── Pola per Jam ──────────────────────────────────────────────────────────────
st.markdown('<div class="sec-title">🕐 Pola Konsumsi per Jam dalam Sehari</div>', unsafe_allow_html=True)
col_l, col_r = st.columns([3, 2])

with col_l:
    if compare_on:
        df_jam = df.groupby(["Jam","Bulan"])["Consume Per Hour"].mean().reset_index()
        fig_jam = px.line(df_jam, x="Jam", y="Consume Per Hour", color="Bulan",
                          color_discrete_sequence=COLORS, markers=True,
                          labels={"Consume Per Hour":"Rata-rata (M³)","Jam":"Jam"})
    else:
        df_jam = df.groupby("Jam")["Consume Per Hour"].mean().reset_index()
        fig_jam = px.bar(df_jam, x="Jam", y="Consume Per Hour",
                         color="Consume Per Hour", color_continuous_scale="Blues",
                         labels={"Consume Per Hour":"Rata-rata (M³)","Jam":"Jam"})
        fig_jam.update_coloraxes(showscale=False)

    fig_jam.update_layout(**base_layout(height=300))
    fig_jam.update_xaxes(tickmode="linear", dtick=1)
    st.plotly_chart(fig_jam, use_container_width=True)

with col_r:
    st.markdown("**🔝 5 Jam Konsumsi Tertinggi**")
    df_top = (df.groupby("Jam")["Consume Per Hour"].mean()
               .sort_values(ascending=False).head(5).reset_index())
    df_top["Jam"] = df_top["Jam"].apply(lambda h: f"{h:02d}:00")
    df_top["Rata-rata (M³)"] = df_top["Consume Per Hour"].round(1)
    st.dataframe(df_top[["Jam","Rata-rata (M³)"]], use_container_width=True, hide_index=True)

    bulan_terbesar = (df.groupby("Bulan")["Consume Per Hour"].sum()
                       .sort_values(ascending=False).index[0])
    total_terbesar = df.groupby("Bulan")["Consume Per Hour"].sum().max()
    st.markdown(f"""
    <div class="insight">
    📌 Bulan dengan konsumsi tertinggi adalah <b>{bulan_terbesar}</b>
    dengan total <b>{total_terbesar:,.0f} M³</b>
    </div>
    """, unsafe_allow_html=True)

# ── Perbandingan Per Jam Tiap Bulan (BARU) ───────────────────────────────────
if show_jam_bulan:
    st.markdown('<div class="sec-title">📅 Perbandingan Konsumsi Per Jam Tiap Bulan</div>',
                unsafe_allow_html=True)

    st.caption("Rata-rata konsumsi tiap jam, dibandingkan antar bulan — cocok untuk melihat pergeseran pola penggunaan air.")

    df_jb = df.groupby(["Jam","Bulan","Bulan_Num"])["Consume Per Hour"].mean().reset_index()
    df_jb = df_jb.sort_values(["Bulan_Num","Jam"])

    # Line chart overlay semua bulan
    fig_jb = px.line(df_jb, x="Jam", y="Consume Per Hour", color="Bulan",
                     color_discrete_sequence=COLORS, markers=True,
                     labels={"Consume Per Hour":"Rata-rata (M³)","Jam":"Jam","Bulan":"Bulan"},
                     hover_data={"Jam": True, "Consume Per Hour":":.1f", "Bulan": True})
    fig_jb.update_layout(**base_layout(height=380))
    fig_jb.update_xaxes(tickmode="linear", dtick=1,
                         ticktext=[f"{h:02d}:00" for h in range(24)],
                         tickvals=list(range(24)))
    fig_jb.update_traces(line=dict(width=2.5), marker=dict(size=5))
    st.plotly_chart(fig_jb, use_container_width=True)

    # Bar chart grouped per jam, tiap bulan sebagai grup
    st.caption("Tampilan bar chart — lebih mudah membandingkan nilai tiap jam antar bulan.")
    tab_line, tab_bar, tab_area = st.tabs(["📈 Line","📊 Bar","🌊 Area"])

    with tab_line:
        st.plotly_chart(fig_jb, use_container_width=True, key="jb_line")

    with tab_bar:
        fig_jb_bar = px.bar(df_jb, x="Jam", y="Consume Per Hour", color="Bulan",
                             barmode="group", color_discrete_sequence=COLORS,
                             labels={"Consume Per Hour":"Rata-rata (M³)","Jam":"Jam"})
        fig_jb_bar.update_layout(**base_layout(height=380))
        fig_jb_bar.update_xaxes(tickmode="linear", dtick=1)
        fig_jb_bar.update_traces(marker_line_width=0)
        st.plotly_chart(fig_jb_bar, use_container_width=True)

    with tab_area:
        fig_jb_area = px.area(df_jb, x="Jam", y="Consume Per Hour", color="Bulan",
                               color_discrete_sequence=COLORS,
                               labels={"Consume Per Hour":"Rata-rata (M³)","Jam":"Jam"})
        fig_jb_area.update_layout(**base_layout(height=380))
        fig_jb_area.update_xaxes(tickmode="linear", dtick=1)
        st.plotly_chart(fig_jb_area, use_container_width=True)

    # Tabel perbandingan per jam tiap bulan
    with st.expander("📋 Lihat Tabel Perbandingan"):
        pivot_jb = df_jb.pivot(index="Jam", columns="Bulan", values="Consume Per Hour").round(1)
        pivot_jb.index = [f"{h:02d}:00" for h in pivot_jb.index]
        st.dataframe(pivot_jb, use_container_width=True)

# ── Heatmap ───────────────────────────────────────────────────────────────────
if show_heatmap:
    st.markdown('<div class="sec-title">🗓️ Heatmap Jam × Hari dalam Seminggu</div>',
                unsafe_allow_html=True)

    df_heat = df.groupby(["DayOfWeek","Jam"])["Consume Per Hour"].mean().reset_index()
    df_heat["DayLabel"] = df_heat["DayOfWeek"].map(HARI_LABEL)
    pivot = (df_heat.pivot(index="DayLabel", columns="Jam", values="Consume Per Hour")
                    .reindex([HARI_LABEL[i] for i in range(7)]))

    fig_heat = px.imshow(pivot, color_continuous_scale="Blues", aspect="auto",
                          text_auto=".0f",
                          labels=dict(x="Jam", y="Hari", color="Rata-rata (M³)"))
    fig_heat.update_layout(
        plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF",
        margin=dict(l=4, r=4, t=10, b=4), height=290,
        font=dict(family="Inter", size=11, color="#0A1F3D"),
        coloraxis_colorbar=dict(title="L/jam", tickfont=dict(color="#0A1F3D")),
        xaxis=dict(tickmode="linear", dtick=1,
                   ticktext=[f"{h:02d}" for h in range(24)],
                   tickvals=list(range(24)),
                   tickfont=dict(color="#0A1F3D")),
        yaxis=dict(tickfont=dict(color="#0A1F3D")),
    )
    st.plotly_chart(fig_heat, use_container_width=True)

# ── Distribusi & Statistik ────────────────────────────────────────────────────
if show_dist:
    st.markdown('<div class="sec-title">📦 Distribusi Konsumsi per Bulan</div>',
                unsafe_allow_html=True)
    col_d1, col_d2 = st.columns(2)

    with col_d1:
        fig_box = px.box(df, x="Bulan", y="Consume Per Hour", color="Bulan",
                          color_discrete_sequence=COLORS, points="outliers",
                          labels={"Consume Per Hour":"Konsumsi (M³)","Bulan":""})
        fig_box.update_layout(**base_layout(height=330, show_legend=False))
        fig_box.update_traces(marker=dict(color="#1A6FBF", size=3, opacity=0.5))
        st.plotly_chart(fig_box, use_container_width=True)

    with col_d2:
        st.markdown("**📋 Statistik per Bulan**")
        df_stat = (df.groupby("Bulan")["Consume Per Hour"]
                     .agg(Total="sum", Rata2="mean", Maks="max", Min="min", StdDev="std")
                     .reset_index())
        df_stat["Total"]  = df_stat["Total"].round(0).astype(int)
        df_stat["Rata2"]  = df_stat["Rata2"].round(1)
        df_stat["Maks"]   = df_stat["Maks"].round(0).astype(int)
        df_stat["Min"]    = df_stat["Min"].round(0).astype(int)
        df_stat["StdDev"] = df_stat["StdDev"].round(1)
        df_stat.columns   = ["Bulan","Total (M³)","Rata-rata (M³)","Maks (M³)","Min (M³)","Std Dev"]
        st.dataframe(df_stat, use_container_width=True, hide_index=True, height=300)

# ── Footer ────────────────────────────────────────────────────────────────────
st.divider()
st.markdown(
    "<div style='text-align:center;color:#7A9CC4;font-size:0.75rem;padding-bottom:12px'>"
    "💧 Dashboard Konsumsi Air · Data diperbarui otomatis dari <code>data_air.xlsx</code> · Satuan: M³"
    "</div>", unsafe_allow_html=True
)
