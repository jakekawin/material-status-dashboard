"""
Material Status Dashboard — Streamlit App
Project: Terra Mat | Author: auto-generated
"""

import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import time
import json

# ─── PAGE CONFIG ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Material Status Dashboard",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── STYLES ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1F4E78 0%, #2E75B6 100%);
        padding: 20px 28px; border-radius: 10px; margin-bottom: 20px;
        color: white;
    }
    .main-header h1 { margin: 0; font-size: 26px; font-weight: 800; }
    .main-header p  { margin: 4px 0 0; opacity: .8; font-size: 13px; }

    .kpi-card {
        background: white; border-radius: 10px; padding: 16px 20px;
        border-left: 5px solid #2E75B6; box-shadow: 0 2px 8px rgba(0,0,0,.08);
        text-align: center;
    }
    .kpi-label { font-size: 11px; color: #777; font-weight: 600; text-transform: uppercase; letter-spacing: .5px; }
    .kpi-value { font-size: 32px; font-weight: 800; margin: 4px 0; }
    .kpi-sub   { font-size: 11px; color: #999; }

    .section-header {
        background: #2E75B6; color: white; padding: 8px 16px;
        border-radius: 6px; font-weight: 700; font-size: 14px; margin: 16px 0 8px;
    }
    .stDataFrame { border-radius: 8px; overflow: hidden; }
    div[data-testid="stExpander"] { border: 1px solid #e0e0e0; border-radius: 8px; }

    .status-received { color: #375623; background: #E2EFDA; padding: 2px 10px; border-radius: 20px; font-size: 12px; font-weight: 600; }
    .status-shortage { color: #C00000; background: #FCE4D6; padding: 2px 10px; border-radius: 20px; font-size: 12px; font-weight: 600; }
    .status-done     { color: #375623; background: #E2EFDA; padding: 2px 10px; border-radius: 20px; font-size: 12px; font-weight: 600; }
    .status-pending  { color: #7F5A00; background: #FFF0D0; padding: 2px 10px; border-radius: 20px; font-size: 12px; font-weight: 600; }

    .update-success { background: #E2EFDA; border: 1px solid #70AD47; border-radius: 8px; padding: 12px 16px; color: #375623; }
    .footer { text-align: center; color: #aaa; font-size: 11px; margin-top: 30px; padding-top: 16px; border-top: 1px solid #f0f0f0; }
</style>
""", unsafe_allow_html=True)

# ─── CONSTANTS ───────────────────────────────────────────────────────────────
SHEET_NAME     = "Material_Status_Dashboard"   # ← ชื่อ Google Sheets ของคุณ
RISERS         = ["RS1","RS2","RS3","RS4","RS5","RS6","RS7","RS8"]
SYSTEMS        = ["CHWF","CHWR"]
RECV_OPTIONS   = ["Received","Shortage",""]
WORK_OPTIONS   = ["Done","Pending",""]
EDIT_PASSWORD  = st.secrets.get("EDIT_PASSWORD", "1234")   # ตั้งใน Streamlit secrets

# ─── GOOGLE SHEETS CONNECTION ────────────────────────────────────────────────
@st.cache_resource(ttl=300)  # cache 5 min
def get_gc():
    """Connect to Google Sheets using service account credentials."""
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]
    # credentials stored in .streamlit/secrets.toml  ← see SETUP GUIDE
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    return gspread.authorize(creds)


@st.cache_data(ttl=30)  # refresh every 30 seconds
def load_data():
    """Load all RS data from Google Sheets into a DataFrame."""
    gc = get_gc()
    sh = gc.open(SHEET_NAME)
    ws = sh.worksheet("Data")   # sheet ชื่อ "Data"
    records = ws.get_all_records()
    df = pd.DataFrame(records)
    # ensure correct dtypes — clean ALL key columns
    for col in ["Riser","System","Receiving Status","Work Status"]:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].fillna("").astype(str).str.strip().replace("nan","")
    return df


def save_row(row_index: int, recv_status: str, work_status: str):
    """Update one row in Google Sheets (1-indexed, +2 for header+1-based)."""
    gc = get_gc()
    sh = gc.open(SHEET_NAME)
    ws = sh.worksheet("Data")
    sheet_row = row_index + 2   # header row = 1, data starts row 2
    # columns: Receiving Status = col 6, Work Status = col 7
    ws.update_cell(sheet_row, 6, recv_status)
    ws.update_cell(sheet_row, 7, work_status)
    load_data.clear()           # bust cache → next load picks up new data


# ─── SESSION STATE ───────────────────────────────────────────────────────────
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "last_refresh" not in st.session_state:
    st.session_state.last_refresh = datetime.now()

# ─── HEADER ──────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
  <h1>📦 Material Status Dashboard</h1>
  <p>Project: Terra Mat  ·  Real-time material receiving & work progress tracking</p>
</div>
""", unsafe_allow_html=True)

# ─── SIDEBAR ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.image("https://img.icons8.com/color/96/combo-chart.png", width=60)
    st.markdown("## 🔧 Controls")

    # Auto-refresh toggle
    auto_refresh = st.toggle("Auto Refresh (30s)", value=True)
    if st.button("🔄 Refresh Now"):
        load_data.clear()
        st.rerun()

    st.markdown("---")

    # Filters — ดึง options จาก data จริง
    st.markdown("### 🔍 Filter")
    _df_opts = load_data()
    _risers  = sorted(_df_opts["Riser"].dropna().str.strip().unique().tolist())
    _systems = sorted(_df_opts["System"].dropna().str.strip().unique().tolist())
    sel_riser  = st.selectbox("Riser",  ["ALL"] + _risers)
    sel_system = st.selectbox("System", ["ALL"] + _systems)
    sel_recv   = st.selectbox("Receiving Status", ["ALL","Received","Shortage","Blank"])
    sel_work   = st.selectbox("Work Status",      ["ALL","Done","Pending","Blank"])

    st.markdown("---")

    # Edit mode authentication
    st.markdown("### ✏️ Edit Mode")
    if not st.session_state.authenticated:
        pwd = st.text_input("Password", type="password", placeholder="กรอก password...")
        if st.button("Unlock"):
            if pwd == EDIT_PASSWORD:
                st.session_state.authenticated = True
                st.success("✓ Unlocked!")
                st.rerun()
            else:
                st.error("Password ไม่ถูกต้อง")
    else:
        st.success("🔓 Edit mode active")
        if st.button("🔒 Lock"):
            st.session_state.authenticated = False
            st.rerun()

    st.markdown("---")
    st.markdown(f"<div style='color:#999;font-size:11px'>Last refresh: {st.session_state.last_refresh.strftime('%H:%M:%S')}</div>", unsafe_allow_html=True)


# ─── LOAD DATA ───────────────────────────────────────────────────────────────
try:
    df_all = load_data()
    st.session_state.last_refresh = datetime.now()
except Exception as e:
    st.error(f"❌ ไม่สามารถเชื่อม Google Sheets ได้: {e}")
    st.info("ตรวจสอบ credentials และชื่อ Sheet ใน SETUP_GUIDE.md")
    st.stop()

# ─── APPLY FILTERS ───────────────────────────────────────────────────────────
df = df_all.copy()
df["Riser"]  = df["Riser"].astype(str).str.strip()
df["System"] = df["System"].astype(str).str.strip()

if sel_riser != "ALL":
    df = df[df["Riser"] == sel_riser]
if sel_system != "ALL":
    df = df[df["System"] == sel_system]
if sel_recv != "ALL":
    if sel_recv == "Blank":
        df = df[df["Receiving Status"] == ""]
    else:
        df = df[df["Receiving Status"] == sel_recv]
if sel_work != "ALL":
    if sel_work == "Blank":
        df = df[df["Work Status"] == ""]
    else:
        df = df[df["Work Status"] == sel_work]

total     = len(df)
received  = len(df[df["Receiving Status"] == "Received"])
shortage  = len(df[df["Receiving Status"] == "Shortage"])
recv_blank= len(df[df["Receiving Status"] == ""])
done      = len(df[df["Work Status"] == "Done"])
pending   = len(df[df["Work Status"] == "Pending"])
pct_recv  = received / total * 100 if total > 0 else 0

# ─── KPI CARDS ───────────────────────────────────────────────────────────────
c1,c2,c3,c4,c5,c6 = st.columns(6)
kpi_data = [
    (c1, "Total Items",    total,    "#1F4E78", "#DEEAF1"),
    (c2, "Received",       received, "#375623", "#E2EFDA"),
    (c3, "Shortage",       shortage, "#C00000", "#FCE4D6"),
    (c4, "Recv Blank",     recv_blank,"#595959","#F2F2F2"),
    (c5, "Done",           done,     "#2E75B6", "#BDD7EE"),
    (c6, f"% Received",    f"{pct_recv:.1f}%", "#7F5A00", "#FFF0D0"),
]
for col, label, value, color, bg in kpi_data:
    with col:
        st.markdown(f"""
        <div class="kpi-card" style="border-left-color:{color}; background:{bg}">
          <div class="kpi-label">{label}</div>
          <div class="kpi-value" style="color:{color}">{value}</div>
        </div>
        """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ─── DEBUG (ซ่อนไว้ กด expand เพื่อดู) ─────────────────────────────────────
with st.expander("🔍 Debug — ตรวจสอบ Filter & Data"):
    c_a, c_b = st.columns(2)
    with c_a:
        st.markdown("**Filter ที่เลือก:**")
        st.write(f"- Riser: `{sel_riser}`")
        st.write(f"- System: `{sel_system}`")
        st.write(f"- Receiving Status: `{sel_recv}`")
        st.write(f"- Work Status: `{sel_work}`")
        st.write(f"- แถวใน Google Sheets ทั้งหมด: **{len(df_all)}**")
        st.write(f"- แถวหลัง Filter: **{total}**")
    with c_b:
        st.markdown("**ค่าจริงๆ ใน Google Sheets:**")
        st.write("Riser:", sorted(df_all["Riser"].unique().tolist()))
        st.write("System:", sorted(df_all["System"].unique().tolist()))
        st.write("Receiving Status:", sorted(df_all["Receiving Status"].unique().tolist()))
        st.write("Work Status:", sorted(df_all["Work Status"].unique().tolist()))
    st.markdown("**ตัวอย่าง 5 แถวแรก:**")
    st.dataframe(df_all.head(), use_container_width=True)

# ─── CHARTS ──────────────────────────────────────────────────────────────────
col_chart1, col_chart2 = st.columns(2)

with col_chart1:
    st.markdown('<div class="section-header">📊 Receiving Status by Riser</div>', unsafe_allow_html=True)
    chart_data = df.groupby(["Riser","Receiving Status"]).size().reset_index(name="Count")
    fig1 = px.bar(
        chart_data, x="Riser", y="Count", color="Receiving Status",
        color_discrete_map={"Received":"#70AD47","Shortage":"#C00000","":"#BFBFBF"},
        text="Count", barmode="stack",
    )
    fig1.update_layout(
        plot_bgcolor="white", paper_bgcolor="white",
        legend_title_text="", height=280, margin=dict(t=10,b=10,l=10,r=10),
        font=dict(family="Calibri", size=12),
    )
    fig1.update_traces(textposition="inside", textfont_size=10)
    st.plotly_chart(fig1, use_container_width=True)

with col_chart2:
    st.markdown('<div class="section-header">🥧 Overall Progress</div>', unsafe_allow_html=True)
    pie_labels = ["Received","Shortage","Not Updated"]
    pie_values = [received, shortage, recv_blank]
    pie_colors = ["#70AD47","#C00000","#BFBFBF"]
    fig2 = go.Figure(go.Pie(
        labels=pie_labels, values=pie_values,
        hole=.55, marker_colors=pie_colors,
        textinfo="label+percent", textfont_size=11,
    ))
    fig2.update_layout(
        showlegend=False, height=280,
        margin=dict(t=10,b=10,l=10,r=10),
        annotations=[dict(text=f"{pct_recv:.0f}%", x=0.5, y=0.5,
                          font_size=28, font_color="#1F4E78",
                          font_family="Calibri", showarrow=False)]
    )
    st.plotly_chart(fig2, use_container_width=True)

# ─── SUMMARY TABLE BY MATERIAL ───────────────────────────────────────────────
st.markdown('<div class="section-header">📋 Summary by Material & Size</div>', unsafe_allow_html=True)

summary = df.groupby(["Material","Size"]).agg(
    Total=("ITEM No.", "count"),
    Received=("Receiving Status", lambda x: (x=="Received").sum()),
    Shortage=("Receiving Status", lambda x: (x=="Shortage").sum()),
    Recv_Blank=("Receiving Status", lambda x: (x=="").sum()),
    Done=("Work Status", lambda x: (x=="Done").sum()),
    Pending=("Work Status", lambda x: (x=="Pending").sum()),
).reset_index()
summary = summary[summary["Total"] > 0]
summary.columns = ["Material","Size","Total","Received","Shortage","Recv Blank","Done","Pending"]

# Color-code the dataframe
def color_shortage(val):
    if isinstance(val, (int,float)) and val > 0:
        return "background-color:#FCE4D6; color:#C00000; font-weight:bold"
    return ""
def color_received(val):
    if isinstance(val, (int,float)) and val > 0:
        return "background-color:#E2EFDA; color:#375623"
    return ""
def color_done(val):
    if isinstance(val, (int,float)) and val > 0:
        return "background-color:#DEEAF1; color:#1F4E78"
    return ""

styled = summary.style \
    .map(color_shortage, subset=["Shortage"]) \
    .map(color_received, subset=["Received"]) \
    .map(color_done,     subset=["Done"]) \
    .set_properties(**{"text-align":"center"}) \
    .set_table_styles([{"selector":"th","props":[("background","#1F4E78"),("color","white"),("font-weight","bold"),("text-align","center")]}])

st.dataframe(styled, use_container_width=True, height=350)

# ─── DETAIL TABLE + EDIT ─────────────────────────────────────────────────────
st.markdown('<div class="section-header">📝 Detail — Item List</div>', unsafe_allow_html=True)

# Group by Riser for expandable sections
risers_in_view = df["Riser"].unique() if len(df) > 0 else []

for riser in sorted(risers_in_view):
    df_r = df[df["Riser"] == riser].copy()
    for system in sorted(df_r["System"].unique()):
        df_rs = df_r[df_r["System"] == system].copy()
        label = f"**{riser} — {system}**  ({len(df_rs)} items)"
        done_cnt = (df_rs["Work Status"] == "Done").sum()
        recv_cnt = (df_rs["Receiving Status"] == "Received").sum()

        with st.expander(f"{riser}-{system}  ·  {recv_cnt}/{len(df_rs)} received  ·  {done_cnt}/{len(df_rs)} done"):
            if st.session_state.authenticated:
                # EDIT MODE: show editable form
                st.info("✏️ Edit Mode — เลือก row แล้วแก้ไขด้านล่าง")

                # Display table
                display_cols = ["ITEM No.","Size","Material","Receiving Status","Work Status"]
                st.dataframe(df_rs[display_cols].reset_index(drop=True), use_container_width=True, height=200)

                # Edit form
                with st.form(f"edit_{riser}_{system}"):
                    st.markdown(f"**แก้ไข {riser}-{system}:**")
                    edit_col1, edit_col2, edit_col3 = st.columns(3)
                    with edit_col1:
                        item_options = df_rs["ITEM No."].tolist()
                        sel_item = st.selectbox("Item No.", item_options)
                    with edit_col2:
                        new_recv = st.selectbox("Receiving Status", ["","Received","Shortage"])
                    with edit_col3:
                        new_work = st.selectbox("Work Status", ["","Done","Pending"])

                    submitted = st.form_submit_button("💾 Save", type="primary", use_container_width=True)
                    if submitted:
                        # Find the row index in the full dataframe
                        mask = (df_all["Riser"] == riser) & (df_all["System"] == system) & (df_all["ITEM No."] == sel_item)
                        idx = df_all[mask].index
                        if len(idx) > 0:
                            save_row(int(idx[0]), new_recv, new_work)
                            st.success(f"✓ บันทึกแล้ว: {riser}-{system} Item {sel_item} → {new_recv or 'Blank'} / {new_work or 'Blank'}")
                            time.sleep(1)
                            st.rerun()
            else:
                # VIEW MODE: read-only table
                display_cols = ["ITEM No.","Size","Material","Receiving Status","Work Status"]

                def highlight_row(row):
                    styles = [""] * len(row)
                    recv = row.get("Receiving Status","")
                    work = row.get("Work Status","")
                    if recv == "Shortage":
                        styles = ["background-color:#FCE4D6"] * len(row)
                    elif recv == "Received" and work == "Done":
                        styles = ["background-color:#E2EFDA"] * len(row)
                    elif recv == "Received":
                        styles = ["background-color:#EBF3E8"] * len(row)
                    return styles

                disp = df_rs[display_cols].reset_index(drop=True)
                st.dataframe(
                    disp.style.apply(highlight_row, axis=1),
                    use_container_width=True, height=min(200, len(disp)*36+40)
                )

# ─── FOOTER ──────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="footer">
  Material Status Dashboard · Project Terra Mat ·
  Auto-refresh every 30s · Last updated: {datetime.now().strftime("%d %b %Y %H:%M")}
</div>
""", unsafe_allow_html=True)

# ─── AUTO REFRESH ────────────────────────────────────────────────────────────
# Use HTML meta-refresh (non-blocking) instead of time.sleep()
if auto_refresh:
    st.markdown('<meta http-equiv="refresh" content="30">', unsafe_allow_html=True)
