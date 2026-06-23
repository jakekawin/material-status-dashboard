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
@st.cache_data(ttl=86400)
def ensure_dwg_column():
    """Add 'Dwg Status' header to Google Sheets column H if missing."""
    gc = get_gc()
    sh = gc.open(SHEET_NAME)
    ws = sh.worksheet("Data")
    headers = ws.row_values(1)
    if len(headers) < 8 or headers[7] != "Dwg Status":
        ws.update_cell(1, 8, "Dwg Status")
    return True


@st.cache_data(ttl=300)  # refresh every 5 minutes
def load_data():
    """Load all RS data from Google Sheets into a DataFrame."""
    gc = get_gc()
    sh = gc.open(SHEET_NAME)
    ws = sh.worksheet("Data")
    records = ws.get_all_records()
    df = pd.DataFrame(records)
    for col in ["Riser","System","Receiving Status","Work Status","Dwg Status"]:
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

    _LV7_CODES = [
        "CDWF750-01", "CDWF500-01", "CDWF500-02", "CDWF500-03",
        "CDWR750-01", "CDWR500-01", "CDWR500-02", "CDWR500-03",
        "EQ600-01",
        "P1-FUF150-01", "P1-FUR150-01", "P1-FUR150-02",
        "P2-FUF150-01", "P2-FUR150-01", "P2-FUR150-02",
        "P3-FUF150-01", "P3-FUR150-01", "P3-FUR150-02",
        "FUW100-01", "FUW100-02", "FUW100-03",
        "CDP100-01",
    ]
    required_sheets  = ["RS1-2","RS3-4","RS5-6","RS7-8"]
    optional_sheets  = _LV7_CODES
    missing = [s for s in required_sheets if s not in xls.sheet_names]
    if missing:
        return None, [f"ไม่พบ sheet: {', '.join(missing)}"]

    all_rows = []
    sheet_group_starts = {
        "RS1-2": [1, 7, 13, 19],
        "RS3-4": [1, 7, 13, 19],
        "RS5-6": [1, 7, 13, 19],
        "RS7-8": [1, 7, 13, 19],
    }
    for _code in _LV7_CODES:
        sheet_group_starts[_code] = [1]
    all_sheets = required_sheets + [s for s in optional_sheets if s in xls.sheet_names]
    for sheet in all_sheets:
        df_raw = pd.read_excel(excel_file, sheet_name=sheet, header=None)
        group_starts = sheet_group_starts.get(sheet, [1, 7, 13, 19])
        for gs in group_starts:
            group_header = str(df_raw.iloc[1, gs]).strip() if gs < df_raw.shape[1] else ""
            if not group_header or group_header == "nan":
                continue
            if sheet in _LV7_CODES:
                riser  = f"Lvl7-{sheet}"
                system = ""
            else:
                parts  = group_header.split("-")
                riser  = parts[0] if len(parts) >= 1 else ""
                system = "-".join(parts[1:]) if len(parts) >= 2 else ""
            for row_idx in range(3, len(df_raw)):
                row = df_raw.iloc[row_idx]
                if gs >= len(row): continue
                item_no = row.iloc[gs]
                if pd.isna(item_no) or str(item_no).strip() == "": continue
                try:
                    item_no_str = str(int(float(str(item_no))))
                except (ValueError, TypeError):
                    continue  # skip header rows / non-numeric cells
                recv = row.iloc[gs+3] if gs+3 < len(row) else ""
                work = row.iloc[gs+4] if gs+4 < len(row) else ""
                dwg  = row.iloc[gs+5] if gs+5 < len(row) else ""
                all_rows.append({
                    "Riser":            str(riser).strip(),
                    "System":           str(system).strip(),
                    "ITEM No.":         item_no_str,
                    "Size":             str(row.iloc[gs+1]).strip() if not pd.isna(row.iloc[gs+1]) else "",
                    "Material":         str(row.iloc[gs+2]).strip() if not pd.isna(row.iloc[gs+2]) else "",
                    "Receiving Status": "" if pd.isna(recv) else str(recv).strip(),
                    "Work Status":      "" if pd.isna(work) else str(work).strip(),
                    "Dwg Status":       "" if pd.isna(dwg)  else str(dwg).strip(),
                })
    return pd.DataFrame(all_rows), errors


def validate_and_diff(df_sheet: pd.DataFrame, df_excel: pd.DataFrame) -> tuple:
    """Return (changes_df, warnings). Changes = rows where Recv/Work status differs."""
    warnings = []
    key = ["Riser","System","ITEM No."]

    # Normalize types before merge — convert ITEM No. to string in both
    df_sheet = df_sheet.copy()
    df_excel = df_excel.copy()
    df_sheet["ITEM No."] = df_sheet["ITEM No."].astype(str).str.strip()
    df_excel["ITEM No."] = df_excel["ITEM No."].astype(str).str.strip()

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
    status_cols = ["Receiving Status","Work Status","Dwg Status"]
    # only use cols that exist in both
    status_cols = [c for c in status_cols if c in df_sheet.columns and c in df_excel.columns]

    merged = df_sheet[key + status_cols].merge(
        df_excel[key + status_cols],
        on=key, suffixes=("_old","_new"), how="inner"
    )
    mask = False
    for c in status_cols:
        mask = mask | (merged[f"{c}_old"] != merged[f"{c}_new"])
    changed = merged[mask].copy()
    rename_map = {}
    for c in status_cols:
        rename_map[f"{c}_old"] = f"{c} (เก่า)"
        rename_map[f"{c}_new"] = f"{c} (ใหม่)"
    changed = changed.rename(columns=rename_map)
    return changed.reset_index(drop=True), warnings


def do_import(df_sheet: pd.DataFrame, df_excel: pd.DataFrame) -> int:
    """Batch-update Receiving Status & Work Status in Google Sheets."""
    gc = get_gc()
    sh = gc.open(SHEET_NAME)
    ws = sh.worksheet("Data")

    key = ["Riser","System","ITEM No."]
    df_sheet = df_sheet.copy()
    df_excel = df_excel.copy()
    df_sheet["ITEM No."] = df_sheet["ITEM No."].astype(str).str.strip()
    df_excel["ITEM No."] = df_excel["ITEM No."].astype(str).str.strip()

    df_sheet_idx = df_sheet[key].copy()
    df_sheet_idx["_idx"] = df_sheet_idx.index

    import_cols = ["Receiving Status","Work Status"]
    if "Dwg Status" in df_excel.columns:
        import_cols.append("Dwg Status")

    merged = df_sheet_idx.merge(
        df_excel[key + import_cols],
        on=key, how="inner"
    )

    updates = []
    for _, row in merged.iterrows():
        gsheet_row = int(row["_idx"]) + 2
        values = [row["Receiving Status"], row["Work Status"]]
        if "Dwg Status" in import_cols:
            values.append(row["Dwg Status"])
        col_end = "H" if "Dwg Status" in import_cols else "G"
        updates.append({
            "range": f"F{gsheet_row}:{col_end}{gsheet_row}",
            "values": [values]
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
        df[["Riser","System","ITEM No.","Size","Material","Receiving Status","Work Status","Dwg Status"]].to_excel(
            writer, index=False, sheet_name="Detail"
        )
    return output.getvalue()


def save_row(row_index: int, recv_status: str, work_status: str, dwg_status: str = ""):
    """Update one row in Google Sheets (1-indexed, +2 for header+1-based)."""
    gc = get_gc()
    sh = gc.open(SHEET_NAME)
    ws = sh.worksheet("Data")
    sheet_row = row_index + 2
    ws.update_cell(sheet_row, 6, recv_status)
    ws.update_cell(sheet_row, 7, work_status)
    ws.update_cell(sheet_row, 8, dwg_status)
    load_data.clear()


# ─── SESSION STATE ───────────────────────────────────────────────────────────
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "last_refresh" not in st.session_state:
    st.session_state.last_refresh = datetime.now()

# Ensure Dwg Status column exists in Google Sheets
ensure_dwg_column()

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
    auto_refresh = st.toggle("Auto Refresh (15 min)", value=True)
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
    sel_dwg    = st.selectbox("Dwg Status",       ["ALL","Approved","Wait for Approved","Blank"])

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
if sel_dwg != "ALL":
    if sel_dwg == "Blank":
        df = df[df["Dwg Status"] == ""]
    else:
        df = df[df["Dwg Status"] == sel_dwg]

# ─── SPLIT RS vs LV7 ────────────────────────────────────────────────────────
df_rs  = df[df["Riser"].str.startswith("RS",    na=False)]
df_lv7 = df[df["Riser"].str.startswith("Lvl7-", na=False)]

# legacy vars for excel report
total        = len(df)
received     = (df["Receiving Status"] == "Received").sum()
shortage     = (df["Receiving Status"] == "Shortage").sum()
done         = (df["Work Status"] == "Done").sum()
dwg_approved = (df["Dwg Status"] == "Approved").sum()
dwg_wait     = (df["Dwg Status"] == "Wait for Approved").sum()

# ─── DEBUG ───────────────────────────────────────────────────────────────────
with st.expander("🔍 Debug — ตรวจสอบ Filter & Data"):
    c_a, c_b = st.columns(2)
    with c_a:
        st.write(f"- Riser: `{sel_riser}`  System: `{sel_system}`")
        st.write(f"- Total rows: **{len(df_all)}** → after filter: **{total}**")
        st.write(f"- RS rows: **{len(df_rs)}**  LV7 rows: **{len(df_lv7)}**")
    with c_b:
        st.write("Riser values:", sorted(df_all["Riser"].unique().tolist()))
    st.dataframe(df_all.head(), use_container_width=True)

# ─── COLOR HELPERS ────────────────────────────────────────────────────────────
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

# ─── EXCEL DOWNLOAD (all data) ───────────────────────────────────────────────
summary_all = df.groupby(["Material","Size"]).agg(
    Total=("ITEM No.","count"),
    Received=("Receiving Status", lambda x: (x=="Received").sum()),
    Shortage=("Receiving Status", lambda x: (x=="Shortage").sum()),
    Done=("Work Status", lambda x: (x=="Done").sum()),
    Pending=("Work Status", lambda x: (x=="Pending").sum()),
    Dwg_Approved=("Dwg Status", lambda x: (x=="Approved").sum()),
    Dwg_Wait=("Dwg Status", lambda x: (x=="Wait for Approved").sum()),
).reset_index()
summary_all.columns = ["Material","Size","Total","Received","Shortage","Done","Pending","Dwg Approved","Dwg Wait"]
excel_bytes = generate_excel_report(df, summary_all)
fname = f"Material_Status_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
report_placeholder.download_button(
    label="📥 Download Excel",
    data=excel_bytes,
    file_name=fname,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)

# ─── EXPANDER HELPER (defined before tabs) ───────────────────────────────────
def _render_group_expander(riser, system, df_grp):
    """Render a single group expander (view + edit modes)."""
    done_cnt = (df_grp["Work Status"] == "Done").sum()
    recv_cnt = (df_grp["Receiving Status"] == "Received").sum()
    dwg_cnt  = (df_grp["Dwg Status"] == "Approved").sum()
    title    = f"{riser}-{system}" if system else riser
    lbl = f"{title}  ·  {recv_cnt}/{len(df_grp)} received  ·  {done_cnt}/{len(df_grp)} done  ·  {dwg_cnt}/{len(df_grp)} dwg approved"
    with st.expander(lbl):
        if st.session_state.authenticated:
            st.info("✏️ Edit Mode — เลือก row แล้วแก้ไขด้านล่าง")
            display_cols = ["ITEM No.","Size","Material","Receiving Status","Work Status","Dwg Status"]
            st.dataframe(df_grp[display_cols].reset_index(drop=True), use_container_width=True, height=200)
            safe_riser = riser.replace("-","_")
            safe_key   = system.replace("-","_").replace(".","_") if system else "grp"
            with st.form(f"edit_{safe_riser}_{safe_key}"):
                st.markdown(f"**แก้ไข {title}:**")
                ec1, ec2, ec3, ec4 = st.columns(4)
                with ec1: sel_item = st.selectbox("Item No.", df_grp["ITEM No."].tolist())
                with ec2: new_recv = st.selectbox("Receiving Status", ["","Received","Shortage"])
                with ec3: new_work = st.selectbox("Work Status", ["","Done","Pending"])
                with ec4: new_dwg  = st.selectbox("Dwg Status", ["","Approved","Wait for Approved"])
                if st.form_submit_button("💾 Save", type="primary", use_container_width=True):
                    mask = (df_all["Riser"]==riser) & (df_all["System"]==system) & (df_all["ITEM No."].astype(str)==str(sel_item))
                    idx = df_all[mask].index
                    if len(idx) > 0:
                        save_row(int(idx[0]), new_recv, new_work, new_dwg)
                        st.success(f"✓ บันทึกแล้ว: {title} Item {sel_item}")
                        time.sleep(1)
                        st.rerun()
        else:
            display_cols = ["ITEM No.","Size","Material","Receiving Status","Work Status","Dwg Status"]
            def highlight_row(row):
                recv = row.get("Receiving Status","")
                work = row.get("Work Status","")
                if recv == "Shortage":   return ["background-color:#FCE4D6"] * len(row)
                if recv == "Received" and work == "Done": return ["background-color:#E2EFDA"] * len(row)
                if recv == "Received":   return ["background-color:#EBF3E8"] * len(row)
                return [""] * len(row)
            disp = df_grp[display_cols].reset_index(drop=True)
            st.dataframe(disp.style.apply(highlight_row, axis=1),
                use_container_width=True, height=min(200, len(disp)*36+40))

# ─── TABS: RS vs LV7 ─────────────────────────────────────────────────────────
tab_rs, tab_lv7 = st.tabs(["📦 RS Risers (RS1–RS8)", "🏗️ Level 7 Items"])


# ══════════════════════════════ RS TAB ══════════════════════════════════════
with tab_rs:
    _t = len(df_rs)
    _recv = (df_rs["Receiving Status"]=="Received").sum()
    _shrt = (df_rs["Receiving Status"]=="Shortage").sum()
    _done = (df_rs["Work Status"]=="Done").sum()
    _pend = (df_rs["Work Status"]=="Pending").sum()
    _dapp = (df_rs["Dwg Status"]=="Approved").sum()
    _dwat = (df_rs["Dwg Status"]=="Wait for Approved").sum()
    _pr = _recv/_t*100 if _t else 0
    _pd = _done/_t*100 if _t else 0
    _pdwg = _dapp/_t*100 if _t else 0

    c1,c2,c3,c4,c5,c6,c7,c8 = st.columns(8)
    for col, label, value, color, bg in [
        (c1,"Total Items",_t,         "#1F4E78","#DEEAF1"),
        (c2,"Received",   _recv,      "#375623","#E2EFDA"),
        (c3,"Shortage",   _shrt,      "#C00000","#FCE4D6"),
        (c4,"Done",       _done,      "#2E75B6","#BDD7EE"),
        (c5,"% Received", f"{_pr:.1f}%","#7F5A00","#FFF0D0"),
        (c6,"% Done",     f"{_pd:.1f}%","#2E75B6","#DEEAF1"),
        (c7,"Dwg Approved",_dapp,     "#375623","#E2EFDA"),
        (c8,"Dwg Wait",   _dwat,      "#C00000","#FCE4D6"),
    ]:
        with col:
            st.markdown(f"""<div class="kpi-card" style="border-left-color:{color};background:{bg}">
              <div class="kpi-label">{label}</div>
              <div class="kpi-value" style="color:{color}">{value}</div>
            </div>""", unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    # Charts row 1: Receiving
    col1, col2 = st.columns(2)
    with col1:
        st.markdown('<div class="section-header">📊 Receiving Status by Riser</div>', unsafe_allow_html=True)
        d = df_rs.groupby(["Riser","Receiving Status"]).size().reset_index(name="Count")
        fig = px.bar(d, x="Riser", y="Count", color="Receiving Status",
            color_discrete_map={"Received":"#70AD47","Shortage":"#C00000","":"#BFBFBF"},
            text="Count", barmode="stack")
        fig.update_layout(plot_bgcolor="white",paper_bgcolor="white",legend_title_text="",
            height=280,margin=dict(t=10,b=10,l=10,r=10),font=dict(family="Calibri",size=12))
        fig.update_traces(textposition="inside",textfont_size=10)
        st.plotly_chart(fig, use_container_width=True)
    with col2:
        st.markdown('<div class="section-header">🥧 Overall — Receiving</div>', unsafe_allow_html=True)
        fig = go.Figure(go.Pie(
            labels=["Received","Shortage","Not Updated"],
            values=[_recv,_shrt,(_t-_recv-_shrt)],
            hole=.55, marker_colors=["#70AD47","#C00000","#BFBFBF"],
            textinfo="label+percent",textfont_size=11))
        fig.update_layout(showlegend=False,height=280,margin=dict(t=10,b=10,l=10,r=10),
            annotations=[dict(text=f"{_pr:.0f}%",x=0.5,y=0.5,font_size=28,font_color="#1F4E78",font_family="Calibri",showarrow=False)])
        st.plotly_chart(fig, use_container_width=True)

    # Charts row 2: Work
    col3, col4 = st.columns(2)
    with col3:
        st.markdown('<div class="section-header">🔧 Work Status by Riser</div>', unsafe_allow_html=True)
        d = df_rs.groupby(["Riser","Work Status"]).size().reset_index(name="Count")
        fig = px.bar(d, x="Riser", y="Count", color="Work Status",
            color_discrete_map={"Done":"#2E75B6","Pending":"#FFC000","":"#BFBFBF"},
            text="Count", barmode="stack")
        fig.update_layout(plot_bgcolor="white",paper_bgcolor="white",legend_title_text="",
            height=280,margin=dict(t=10,b=10,l=10,r=10),font=dict(family="Calibri",size=12))
        fig.update_traces(textposition="inside",textfont_size=10)
        st.plotly_chart(fig, use_container_width=True)
    with col4:
        st.markdown('<div class="section-header">🥧 Overall — Work</div>', unsafe_allow_html=True)
        fig = go.Figure(go.Pie(
            labels=["Done","Pending","Not Updated"],
            values=[_done,_pend,(_t-_done-_pend)],
            hole=.55, marker_colors=["#2E75B6","#FFC000","#BFBFBF"],
            textinfo="label+percent",textfont_size=11))
        fig.update_layout(showlegend=False,height=280,margin=dict(t=10,b=10,l=10,r=10),
            annotations=[dict(text=f"{_pd:.0f}%",x=0.5,y=0.5,font_size=28,font_color="#1F4E78",font_family="Calibri",showarrow=False)])
        st.plotly_chart(fig, use_container_width=True)

    # Charts row 3: Dwg
    col5, col6 = st.columns(2)
    with col5:
        st.markdown('<div class="section-header">📐 Dwg Status by Riser</div>', unsafe_allow_html=True)
        d = df_rs.groupby(["Riser","Dwg Status"]).size().reset_index(name="Count")
        fig = px.bar(d, x="Riser", y="Count", color="Dwg Status",
            color_discrete_map={"Approved":"#375623","Wait for Approved":"#FFC000","":"#BFBFBF"},
            text="Count", barmode="stack")
        fig.update_layout(plot_bgcolor="white",paper_bgcolor="white",legend_title_text="",
            height=280,margin=dict(t=10,b=10,l=10,r=10),font=dict(family="Calibri",size=12))
        fig.update_traces(textposition="inside",textfont_size=10)
        st.plotly_chart(fig, use_container_width=True)
    with col6:
        st.markdown('<div class="section-header">🥧 Overall — Dwg</div>', unsafe_allow_html=True)
        fig = go.Figure(go.Pie(
            labels=["Approved","Wait for Approved","Not Updated"],
            values=[_dapp,_dwat,(_t-_dapp-_dwat)],
            hole=.55, marker_colors=["#375623","#FFC000","#BFBFBF"],
            textinfo="label+percent",textfont_size=11))
        fig.update_layout(showlegend=False,height=280,margin=dict(t=10,b=10,l=10,r=10),
            annotations=[dict(text=f"{_pdwg:.0f}%",x=0.5,y=0.5,font_size=28,font_color="#1F4E78",font_family="Calibri",showarrow=False)])
        st.plotly_chart(fig, use_container_width=True)

    # % Done bar
    st.markdown('<div class="section-header">📊 % Done by Riser</div>', unsafe_allow_html=True)
    dbr = df_rs.groupby("Riser").apply(
        lambda x: round((x["Work Status"]=="Done").sum()/len(x)*100,1)
    ).reset_index(name="% Done").sort_values("Riser")
    fig = px.bar(dbr, x="Riser", y="% Done",
        text=dbr["% Done"].apply(lambda v: f"{v:.1f}%"),
        color="% Done", color_continuous_scale=[[0,"#BFBFBF"],[0.5,"#2E75B6"],[1,"#1F4E78"]], range_color=[0,100])
    fig.update_layout(plot_bgcolor="white",paper_bgcolor="white",height=260,
        margin=dict(t=10,b=10,l=10,r=10),font=dict(family="Calibri",size=12),
        coloraxis_showscale=False,yaxis=dict(range=[0,110],ticksuffix="%"))
    fig.update_traces(textposition="outside",textfont_size=11)
    st.plotly_chart(fig, use_container_width=True)

    # Summary table
    st.markdown('<div class="section-header">📋 Summary by Material & Size</div>', unsafe_allow_html=True)
    smry = df_rs.groupby(["Material","Size"]).agg(
        Total=("ITEM No.","count"),
        Received=("Receiving Status",lambda x:(x=="Received").sum()),
        Shortage=("Receiving Status",lambda x:(x=="Shortage").sum()),
        Done=("Work Status",lambda x:(x=="Done").sum()),
        Pending=("Work Status",lambda x:(x=="Pending").sum()),
        Dwg_Approved=("Dwg Status",lambda x:(x=="Approved").sum()),
        Dwg_Wait=("Dwg Status",lambda x:(x=="Wait for Approved").sum()),
    ).reset_index()
    smry = smry[smry["Total"]>0]
    smry.columns = ["Material","Size","Total","Received","Shortage","Done","Pending","Dwg Approved","Dwg Wait"]
    st.dataframe(
        smry.style.map(color_shortage,subset=["Shortage"])
                  .map(color_received,subset=["Received","Dwg Approved"])
                  .map(color_done,subset=["Done"])
                  .set_properties(**{"text-align":"center"})
                  .set_table_styles([{"selector":"th","props":[("background","#1F4E78"),("color","white"),("font-weight","bold"),("text-align","center")]}]),
        use_container_width=True, height=350)

    # Detail
    st.markdown('<div class="section-header">📝 Detail — Riser Item List (RS1–RS8)</div>', unsafe_allow_html=True)
    for riser in sorted(df_rs["Riser"].unique()):
        df_r = df_rs[df_rs["Riser"]==riser]
        for system in sorted(df_r["System"].unique()):
            _render_group_expander(riser, system, df_r[df_r["System"]==system].copy())


# ══════════════════════════════ LV7 TAB ═════════════════════════════════════
with tab_lv7:
    if len(df_lv7) == 0:
        st.info("ยังไม่มีข้อมูล Level 7 — กรุณากรอกข้อมูลใน Excel sheets (CDWF750-01 ฯลฯ) แล้วรัน prepare_gsheet.py อีกครั้งครับ")
    else:
        _t7 = len(df_lv7)
        _rv7 = (df_lv7["Receiving Status"]=="Received").sum()
        _sh7 = (df_lv7["Receiving Status"]=="Shortage").sum()
        _dn7 = (df_lv7["Work Status"]=="Done").sum()
        _pn7 = (df_lv7["Work Status"]=="Pending").sum()
        _da7 = (df_lv7["Dwg Status"]=="Approved").sum()
        _dw7 = (df_lv7["Dwg Status"]=="Wait for Approved").sum()
        _pr7 = _rv7/_t7*100 if _t7 else 0
        _pd7 = _dn7/_t7*100 if _t7 else 0
        _pg7 = _da7/_t7*100 if _t7 else 0

        c1,c2,c3,c4,c5,c6,c7,c8 = st.columns(8)
        for col, label, value, color, bg in [
            (c1,"Total Items",_t7,          "#1F4E78","#DEEAF1"),
            (c2,"Received",   _rv7,         "#375623","#E2EFDA"),
            (c3,"Shortage",   _sh7,         "#C00000","#FCE4D6"),
            (c4,"Done",       _dn7,         "#2E75B6","#BDD7EE"),
            (c5,"% Received", f"{_pr7:.1f}%","#7F5A00","#FFF0D0"),
            (c6,"% Done",     f"{_pd7:.1f}%","#2E75B6","#DEEAF1"),
            (c7,"Dwg Approved",_da7,        "#375623","#E2EFDA"),
            (c8,"Dwg Wait",   _dw7,         "#C00000","#FCE4D6"),
        ]:
            with col:
                st.markdown(f"""<div class="kpi-card" style="border-left-color:{color};background:{bg}">
                  <div class="kpi-label">{label}</div>
                  <div class="kpi-value" style="color:{color}">{value}</div>
                </div>""", unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

        # Charts
        col1, col2 = st.columns(2)
        with col1:
            st.markdown('<div class="section-header">📊 Receiving Status by Item</div>', unsafe_allow_html=True)
            d = df_lv7.groupby(["Riser","Receiving Status"]).size().reset_index(name="Count")
            fig = px.bar(d, x="Riser", y="Count", color="Receiving Status",
                color_discrete_map={"Received":"#70AD47","Shortage":"#C00000","":"#BFBFBF"},
                text="Count", barmode="stack")
            fig.update_layout(plot_bgcolor="white",paper_bgcolor="white",legend_title_text="",
                height=300,margin=dict(t=10,b=10,l=10,r=10),font=dict(family="Calibri",size=11),xaxis_tickangle=-45)
            fig.update_traces(textposition="inside",textfont_size=9)
            st.plotly_chart(fig, use_container_width=True)
        with col2:
            st.markdown('<div class="section-header">📐 Dwg Status by Item</div>', unsafe_allow_html=True)
            d = df_lv7.groupby(["Riser","Dwg Status"]).size().reset_index(name="Count")
            fig = px.bar(d, x="Riser", y="Count", color="Dwg Status",
                color_discrete_map={"Approved":"#375623","Wait for Approved":"#FFC000","":"#BFBFBF"},
                text="Count", barmode="stack")
            fig.update_layout(plot_bgcolor="white",paper_bgcolor="white",legend_title_text="",
                height=300,margin=dict(t=10,b=10,l=10,r=10),font=dict(family="Calibri",size=11),xaxis_tickangle=-45)
            fig.update_traces(textposition="inside",textfont_size=9)
            st.plotly_chart(fig, use_container_width=True)

        col3, col4 = st.columns(2)
        with col3:
            st.markdown('<div class="section-header">🥧 Overall — Receiving</div>', unsafe_allow_html=True)
            fig = go.Figure(go.Pie(
                labels=["Received","Shortage","Not Updated"],
                values=[_rv7,_sh7,(_t7-_rv7-_sh7)],
                hole=.55, marker_colors=["#70AD47","#C00000","#BFBFBF"],
                textinfo="label+percent",textfont_size=11))
            fig.update_layout(showlegend=False,height=260,margin=dict(t=10,b=10,l=10,r=10),
                annotations=[dict(text=f"{_pr7:.0f}%",x=0.5,y=0.5,font_size=28,font_color="#1F4E78",font_family="Calibri",showarrow=False)])
            st.plotly_chart(fig, use_container_width=True)
        with col4:
            st.markdown('<div class="section-header">🥧 Overall — Dwg</div>', unsafe_allow_html=True)
            fig = go.Figure(go.Pie(
                labels=["Approved","Wait for Approved","Not Updated"],
                values=[_da7,_dw7,(_t7-_da7-_dw7)],
                hole=.55, marker_colors=["#375623","#FFC000","#BFBFBF"],
                textinfo="label+percent",textfont_size=11))
            fig.update_layout(showlegend=False,height=260,margin=dict(t=10,b=10,l=10,r=10),
                annotations=[dict(text=f"{_pg7:.0f}%",x=0.5,y=0.5,font_size=28,font_color="#1F4E78",font_family="Calibri",showarrow=False)])
            st.plotly_chart(fig, use_container_width=True)

        # Summary table
        st.markdown('<div class="section-header">📋 Summary by Material & Size</div>', unsafe_allow_html=True)
        smry7 = df_lv7.groupby(["Material","Size"]).agg(
            Total=("ITEM No.","count"),
            Received=("Receiving Status",lambda x:(x=="Received").sum()),
            Shortage=("Receiving Status",lambda x:(x=="Shortage").sum()),
            Done=("Work Status",lambda x:(x=="Done").sum()),
            Pending=("Work Status",lambda x:(x=="Pending").sum()),
            Dwg_Approved=("Dwg Status",lambda x:(x=="Approved").sum()),
            Dwg_Wait=("Dwg Status",lambda x:(x=="Wait for Approved").sum()),
        ).reset_index()
        smry7 = smry7[smry7["Total"]>0]
        smry7.columns = ["Material","Size","Total","Received","Shortage","Done","Pending","Dwg Approved","Dwg Wait"]
        st.dataframe(
            smry7.style.map(color_shortage,subset=["Shortage"])
                       .map(color_received,subset=["Received","Dwg Approved"])
                       .map(color_done,subset=["Done"])
                       .set_properties(**{"text-align":"center"})
                       .set_table_styles([{"selector":"th","props":[("background","#1F4E78"),("color","white"),("font-weight","bold"),("text-align","center")]}]),
            use_container_width=True, height=350)

        # Detail
        st.markdown('<div class="section-header">🏗️ Detail — Level 7 Item List</div>', unsafe_allow_html=True)
        for riser in sorted(df_lv7["Riser"].unique()):
            _render_group_expander(riser, "", df_lv7[df_lv7["Riser"]==riser].copy())


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
  Auto-refresh every 15 min · Last updated: {datetime.now().strftime("%d %b %Y %H:%M")}
</div>
""", unsafe_allow_html=True)

# ─── AUTO REFRESH ────────────────────────────────────────────────────────────
# Use HTML meta-refresh (non-blocking) instead of time.sleep()
if auto_refresh:
    st.markdown('<meta http-equiv="refresh" content="900">', unsafe_allow_html=True)
