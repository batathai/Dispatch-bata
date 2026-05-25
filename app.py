"""
Dispatch + SIR Merger — Streamlit Web App
Deploy ฟรีบน https://streamlit.io/cloud
"""

import io
import zipfile
from datetime import datetime

import pandas as pd
import pyxlsb
import streamlit as st

# ── Config ──────────────────────────────────────────────────────
SIR_COLUMNS = [
    "Article", "Category", "SubCategory", "Cost", "Price",
    "globalDepartmentName", "globalCategoryName", "globalBrandName",
]

st.set_page_config(
    page_title="Dispatch + SIR Merger",
    page_icon="⚡",
    layout="centered",
)

# ── Custom CSS ───────────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600;700&display=swap');

    html, body, [class*="css"] { font-family: 'IBM Plex Mono', monospace; }
    .stApp { background: #0d0f1a; }

    h1 { color: #e94560 !important; letter-spacing: -1px; }
    h2, h3 { color: #8be9fd !important; }
    p, label, .stMarkdown { color: #c8d0e0 !important; }

    .stFileUploader {
        background: #16213e !important;
        border: 1px dashed #e94560 !important;
        border-radius: 8px !important;
        padding: 12px !important;
    }
    .stButton > button {
        background: #e94560 !important;
        color: white !important;
        font-family: 'IBM Plex Mono', monospace !important;
        font-weight: 700 !important;
        border: none !important;
        border-radius: 6px !important;
        padding: 10px 32px !important;
        width: 100% !important;
        font-size: 16px !important;
    }
    .stButton > button:hover { background: #c73652 !important; }

    .stAlert { border-radius: 6px !important; }
    .metric-box {
        background: #16213e;
        border: 1px solid #334;
        border-radius: 8px;
        padding: 16px;
        text-align: center;
    }
    .metric-val { font-size: 28px; font-weight: 700; color: #50fa7b; }
    .metric-lbl { font-size: 11px; color: #a8b2d8; margin-top: 4px; }

    div[data-testid="stExpander"] {
        background: #16213e !important;
        border: 1px solid #334 !important;
        border-radius: 8px !important;
    }
</style>
""", unsafe_allow_html=True)


# ── Header ───────────────────────────────────────────────────────
st.title("⚡ Dispatch + SIR Merger")
st.markdown("อัปเดตข้อมูล Dispatch ด้วย SIR โดยอัตโนมัติ · Match ด้วย **Article**")
st.divider()


# ── Helpers ─────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_dispatch(file_bytes: bytes) -> pd.DataFrame:
    rows, headers = [], None
    with pyxlsb.open_workbook(io.BytesIO(file_bytes)) as wb:
        with wb.get_sheet(1) as sheet:
            for i, row in enumerate(sheet.rows()):
                vals = [c.v for c in row]
                if i == 0:
                    headers = vals
                else:
                    rows.append(vals)
    return pd.DataFrame(rows, columns=headers)


@st.cache_data(show_spinner=False)
def load_sir(file_bytes: bytes, filename: str) -> pd.DataFrame:
    if filename.endswith(".zip"):
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
            xlsx_names = [n for n in zf.namelist() if n.endswith(".xlsx")]
            if not xlsx_names:
                raise ValueError("ไม่พบ .xlsx ใน zip")
            target = next((n for n in xlsx_names if "All" in n), xlsx_names[0])
            with zf.open(target) as f:
                return pd.read_excel(io.BytesIO(f.read()))
    else:
        return pd.read_excel(io.BytesIO(file_bytes))


def merge_data(dispatch_df: pd.DataFrame, sir_df: pd.DataFrame):
    # Normalize Article key
    if "Article" not in sir_df.columns:
        cands = [c for c in sir_df.columns if "article" in c.lower()]
        if not cands:
            raise KeyError(f"ไม่พบคอลัมน์ 'Article' ใน SIR\nคอลัมน์: {list(sir_df.columns[:15])}")
        sir_df = sir_df.rename(columns={cands[0]: "Article"})

    dispatch_df["Article"] = dispatch_df["Article"].astype(str).str.strip()
    sir_df["Article"] = sir_df["Article"].astype(str).str.strip()

    available = [c for c in SIR_COLUMNS if c in sir_df.columns]
    missing = [c for c in SIR_COLUMNS if c not in sir_df.columns]

    # ดึงเฉพาะคอลัมน์ที่ต้องการ (ไม่รวม Article ที่ซ้ำกับ Dispatch)
    sir_cols_no_article = [c for c in available if c != "Article"]
    sir_lookup = sir_df.drop_duplicates(subset="Article")[["Article"] + sir_cols_no_article]
    merged = dispatch_df.merge(sir_lookup, on="Article", how="left", suffixes=("", "_SIR"))

    check_col = sir_cols_no_article[0] if sir_cols_no_article else None
    matched = int(merged[check_col].notna().sum()) if check_col else 0
    return merged, matched, sir_cols_no_article, missing


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Dispatch_Merged")
    return buf.getvalue()


# ── Upload section ───────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    st.markdown("### 📊 Dispatch File")
    st.caption("รองรับ .xlsb")
    dispatch_file = st.file_uploader(
        "dispatch", type=["xlsb"],
        label_visibility="collapsed", key="dispatch"
    )

with col2:
    st.markdown("### 📦 SIR File")
    st.caption("รองรับ .zip (มี .xlsx ข้างใน) หรือ .xlsx โดยตรง")
    sir_file = st.file_uploader(
        "sir", type=["zip", "xlsx"],
        label_visibility="collapsed", key="sir"
    )

st.divider()

# ── Preview section ──────────────────────────────────────────────
if dispatch_file or sir_file:
    st.markdown("### 👀 Preview")
    prev1, prev2 = st.columns(2)

    if dispatch_file:
        with prev1:
            with st.spinner("กำลังอ่าน Dispatch..."):
                try:
                    dispatch_df = load_dispatch(dispatch_file.read())
                    dispatch_file.seek(0)
                    st.success(f"✅ {len(dispatch_df):,} rows · {len(dispatch_df.columns)} cols")
                    with st.expander("ดูตัวอย่าง 5 แถวแรก"):
                        st.dataframe(dispatch_df.head(), use_container_width=True)
                except Exception as e:
                    st.error(f"❌ {e}")
                    dispatch_df = None

    if sir_file:
        with prev2:
            with st.spinner("กำลังอ่าน SIR..."):
                try:
                    sir_df = load_sir(sir_file.read(), sir_file.name)
                    sir_file.seek(0)
                    st.success(f"✅ {len(sir_df):,} rows · {len(sir_df.columns)} cols")
                    with st.expander("ดูตัวอย่าง 5 แถวแรก"):
                        st.dataframe(sir_df.head(), use_container_width=True)
                except Exception as e:
                    st.error(f"❌ {e}")
                    sir_df = None

st.divider()

# ── Run button ───────────────────────────────────────────────────
ready = dispatch_file is not None and sir_file is not None

if not ready:
    st.info("📂 กรุณา upload ทั้ง Dispatch และ SIR ก่อนรัน")

run_clicked = st.button("▶  รัน Merge", disabled=not ready)

if run_clicked and ready:
    try:
        dispatch_file.seek(0)
        sir_file.seek(0)

        with st.spinner("กำลังโหลดและ merge ข้อมูล..."):
            d_df = load_dispatch(dispatch_file.read())
            s_df = load_sir(sir_file.read(), sir_file.name)
            merged_df, matched, available, missing = merge_data(d_df, s_df)

        # Results metrics
        st.markdown("### ✅ ผลการ Merge")
        m1, m2, m3, m4 = st.columns(4)
        total = len(merged_df)
        pct = matched / total * 100 if total else 0

        with m1:
            st.markdown(f"""<div class="metric-box">
                <div class="metric-val">{total:,}</div>
                <div class="metric-lbl">Total Rows</div></div>""",
                unsafe_allow_html=True)
        with m2:
            st.markdown(f"""<div class="metric-box">
                <div class="metric-val">{matched:,}</div>
                <div class="metric-lbl">Matched</div></div>""",
                unsafe_allow_html=True)
        with m3:
            color = "#50fa7b" if pct >= 80 else "#ffb86c" if pct >= 50 else "#ff5555"
            st.markdown(f"""<div class="metric-box">
                <div class="metric-val" style="color:{color}">{pct:.1f}%</div>
                <div class="metric-lbl">Match Rate</div></div>""",
                unsafe_allow_html=True)
        with m4:
            st.markdown(f"""<div class="metric-box">
                <div class="metric-val">{len(merged_df.columns)}</div>
                <div class="metric-lbl">Total Columns</div></div>""",
                unsafe_allow_html=True)

        if missing:
            st.warning(f"⚠️ คอลัมน์ SIR ที่ไม่พบ: `{', '.join(missing)}`")

        st.caption(f"คอลัมน์ที่ merge แล้ว: {', '.join(available)}")

        # Preview merged
        with st.expander("ดูตัวอย่างผลลัพธ์ 10 แถวแรก"):
            st.dataframe(merged_df.head(10), use_container_width=True)

        # Download
        st.divider()
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        fname = f"Dispatch_merged_{ts}.xlsx"

        with st.spinner("กำลังสร้างไฟล์ Excel..."):
            xlsx_bytes = to_excel_bytes(merged_df)

        st.download_button(
            label=f"⬇️  Download {fname}",
            data=xlsx_bytes,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"❌ Error: {e}")
        st.exception(e)

# ── Footer ───────────────────────────────────────────────────────
st.divider()
st.caption("Dispatch + SIR Merger · Match key: Article · Output: A–BE + SIR columns")
