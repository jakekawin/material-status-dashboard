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
import io

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

    @media print {
        [data-testid="stSidebar"], [data-testid="stToolbar"],
        button, .stButton, .stDownloadButton { display: none !important; }
        .main-header { background: #1F4E78 !important; -webkit-print-color-adjust: exact; }
        .kpi-card { border: 1px solid #ccc !important; break-inside: avoid; }
        body { font-size: 11px; }
    }
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


@st.cache_data(ttl=300)  # refresh every 5 minutes
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


def parse_excel_for_import(excel_file) -> tuple:
    """Parse uploaded Excel, return (df, errors)."""
    errors = []
    try:
        xls = pd.ExcelFile(excel_file)
    except Exception as e:
        return None, [f"ไม่สามารถอ่านไฟล์ได้: {e}"]

    required_sheets = ["RS1-2","RS3-4","RS5-6","RS7-8"]
    missing = [s for s in required_sheets if s not in xls.sheet_names]
    if missing:
        return None, [f"ไม่พบ sheet: {', '.join(missing)}"]

    all_rows = []
    group_starts = [1, 6, 11, 16]
    for sheet in required_sheets:
        df_raw = pd.read_excel(excel_file, sheet_name=sheet, header=None)
        for gs in group_starts:
            group_header = str(df_raw.iloc[1, gs]).strip() if gs < df_raw.shape[1] else ""
            if not group_header or group_header == "nan":
                continue
            parts = group_header.split("-")
            riser  = parts[0] if len(parts) >= 1 else ""
            system = parts[1] if len(parts) >= 2 else ""
            for row_idx in range(3, len(df_raw)):
                row = df_raw.iloc[row_idx]
                if gs >= len(row): continue
                item_no = row.iloc[gs]
                if pd.isna(item_no) or str(item_no).strip() == "": continue
                recv = row.iloc[gs+3] if gs+3 < len(row) else ""
                work = row.iloc[gs+4] if gs+4 < len(row) else ""
                all_rows.append({
                    "Riser":            str(riser).strip(),
                    "System":           str(system).strip(),
                    "ITEM No.":         str(int(item_no)),
                    "Size":             str(row.iloc[gs+1]).strip() if not pd.isna(row.iloc[gs+1]) else "",
                    "Material":         str(row.iloc[gs+2]).strip() if not pd.isna(row.iloc[gs+2]) else "",
                    "Receiving Status": "" if pd.isna(recv) else str(recv).strip(),
                    "Work Status":      "" if pd.isna(work) else str(work).strip(),
                })
    return pd.DataFrame(all_rows), errors


def validate_and_diff(df_sheet: pd.DataFrame, df_excel: pd.DataFrame) -> tuple:
    """Return (changes_df, warnings). Changes = rows where Recv/Work status differs."""
    warnings = []
    key = ["Riser","System","ITEM No."]

    # Check structure match
    sheet_keys = set(zip(df_sheet["Riser"], df_sheet["System"], df_sheet["ITEM No."]))
    excel_keys = set(zip(df_excel["Riser"], df_excel["System"], df_excel["ITEM No."]))
    only_excel = excel_keys - sheet_keys
    only_sheet = sheet_keys - excel_keys
    if only_excel:
        warnings.append(f"⚠️ {len(only_excel)} items ใน Excel ที่ไม่มีใน Sheet (จะถูกข้าม)")
    if only_sheet:
        warnings.append(f"⚠️ {len(only_sheet)} items ใน Sheet ที่ไม่มีใน Excel (จะไม่ถูกแตะต้อง)")

    # Diff: merge on key
    merged = df_sheet[key + ["Receiving Status","Work Status"]].merge(
        df_excel[key + ["Receiving Status","Work Status"]],
        on=key, suffixes=("_old","_new"), how="inner"
    )
    changed = merged[
        (merged["Receiving Status_old"] != merged["Receiving Status_new"]) |
        (merged["Work Status_old"]      != merged["Work Status_new"])
    ].copy()
    changed = changed.rename(columns={
        "Receiving Status_old": "Receiving Status (เก่า)",
        "Work Status_old":      "Work Status (เก่า)",
        "Receiving Status_new": "Receiving Status (ใหม่)",
        "Work Status_new":      "Work Status (ใหม่)",
    })
    return changed.reset_index(drop=True), warnings


def do_import(df_sheet: pd.DataFrame, df_excel: pd.DataFrame) -> int:
    """Batch-update Receiving Status & Work Status in Google Sheets."""
    gc = get_gc()
    sh = gc.open(SHEET_NAME)
    ws = sh.worksheet("Data")

    key = ["Riser","System","ITEM No."]
    df_sheet_idx = df_sheet[key].copy()
    df_sheet_idx["_idx"] = df_sheet_idx.index

    merged = df_sheet_idx.merge(
        df_excel[key + ["Receiving Status","Work Status"]],
        on=key, how="inner"
    )

    updates = []
    for _, row in merged.iterrows():
        gsheet_row = int(row["_idx"]) + 2   # +1 header, +1 for 1-based
        updates.append({
            "range": f"F{gsheet_row}:G{gsheet_row}",
            "values": [[row["Receiving Status"], row["Work Status"]]]
        })

    for i in range(0, len(updates), 100):   # batch in chunks of 100
        ws.batch_update(updates[i:i+100])

    load_data.clear()
    return len(updates)


def generate_excel_report(df: pd.DataFrame, summary_df: pd.DataFrame) -> bytes:
    """Generate Excel report with two sheets: Summary and Detail."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, index=False, sheet_name="Summary")
        df[["Riser","System","ITEM No.","Size","Material","Receiving Status","Work Status"]].to_excel(
            writer, index=False, sheet_name="Detail"
        )
    return output.getvalue()


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
    auto_refresh = st.toggle("Auto Refresh (5 min)", value=True)
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

    # Report buttons
    st.markdown("### 📄 Report")

    # Excel download (populated after df is ready — placeholder here)
    report_placeholder = st.empty()

    # Print button
    if st.button("🖨️ Print Page", use_container_width=True):
        st.markdown("<script>window.print();</script>", unsafe_allow_html=True)

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

# ─── FILL EXCEL DOWNLOAD BUTTON ──────────────────────────────────────────────
excel_bytes = generate_excel_report(df, summary)
fname = f"Material_Status_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
report_placeholder.download_button(
    label="📥 Download Excel",
    data=excel_bytes,
    file_name=fname,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)

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

# ─── IMPORT FROM EXCEL ───────────────────────────────────────────────────────
if st.session_state.authenticated:
    st.markdown('<div class="section-header">📤 Import จาก Excel (อัพเดต Receiving/Work Status)</div>', unsafe_allow_html=True)

    uploaded_file = st.file_uploader(
        "อัพโหลดไฟล์ Excel (ต้องมี sheet: RS1-2, RS3-4, RS5-6, RS7-8)",
        type=["xlsx"], key="import_uploader"
    )

    if uploaded_file:
        with st.spinner("กำลังตรวจสอบไฟล์..."):
            df_excel, parse_errors = parse_excel_for_import(uploaded_file)

        if parse_errors:
            for e in parse_errors:
                st.error(e)
        else:
            changes_df, imp_warnings = validate_and_diff(df_all, df_excel)

            # Show warnings
            for w in imp_warnings:
                st.warning(w)

            # Structure check summary
            col_ok1, col_ok2 = st.columns(2)
            col_ok1.metric("Items ใน Excel", len(df_excel))
            col_ok2.metric("Items ใน Google Sheet", len(df_all))

            if len(changes_df) == 0:
                st.success("✅ ข้อมูลตรงกันทั้งหมด — ไม่มีการเปลี่ยนแปลง")
            else:
                st.info(f"พบ **{len(changes_df)} รายการ** ที่ Receiving Status หรือ Work Status เปลี่ยนแปลง:")
                st.dataframe(changes_df, use_container_width=True, height=min(350, len(changes_df)*36+50))

                st.warning("⚠️ กด **ยืนยัน Import** เพื่ออัพเดต Google Sheets — ไม่สามารถ undo ได้")
                if st.button("✅ ยืนยัน Import", type="primary", use_container_width=False):
                    with st.spinner("กำลัง import..."):
                        n_updated = do_import(df_all, df_excel)
                    st.success(f"✓ อัพเดต {n_updated} รายการสำเร็จ!")
                    time.sleep(1)
                    st.rerun()

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
    st.markdown('<meta http-equiv="refresh" content="300">', unsafe_allow_html=True)
