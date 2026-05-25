"""
Dispatch Report + SIR Merger
============================
รัน script นี้บน PC ออฟฟิต เพื่อ merge ข้อมูลจาก:
  - Dispatch report (.xlsb)  →  คอลัมน์ A ถึง BE
  - Official_SIR (.zip/.xlsx) →  Article, Category, SubCategory, Cost, Price,
                                  globalDepartmentName, globalCategoryName, globalBrandName

ผลลัพธ์จะเป็นไฟล์ Excel ใหม่ที่มีคอลัมน์ครบ A-BP

วิธีใช้:
  python dispatch_sir_merge.py

ติดตั้ง dependencies ครั้งแรก:
  pip install pyxlsb pandas openpyxl
"""

import os
import re
import zipfile
import glob
import io
from datetime import datetime
from pathlib import Path

import pandas as pd
import pyxlsb


# ============================================================
# ⚙️  CONFIG — แก้ path ตามจริงก่อนรัน
# ============================================================

# Path ไฟล์ Dispatch report (.xlsb)
DISPATCH_FILE = r"C:\Users\YourName\Downloads\Dispatch_report_May_2026.xlsb"

# Root folder ของ SIR บน Network Drive
SIR_ROOT = r"O:\2026 Official SIR"

# ชื่อ sheet ใน SIR xlsx ที่มีข้อมูล (ปกติ sheet แรก หรือชื่อ "sir")
SIR_SHEET = 0  # ใส่ 0 = sheet แรก หรือใส่ชื่อ เช่น "sir"

# คอลัมน์ที่ดึงจาก SIR มาต่อท้าย (BF-BM)
SIR_COLUMNS = [
    "Article",
    "Category",
    "SubCategory",
    "Cost",
    "Price",
    "globalDepartmentName",
    "globalCategoryName",
    "globalBrandName",
]

# Output file
OUTPUT_DIR = r"C:\Users\YourName\Downloads"


# ============================================================
# 📦  FUNCTIONS
# ============================================================

def find_latest_sir_week() -> int:
    """หา folder week ล่าสุดใน SIR_ROOT (เลขสูงสุด)"""
    folders = [
        int(p.name)
        for p in Path(SIR_ROOT).iterdir()
        if p.is_dir() and p.name.isdigit()
    ]
    if not folders:
        raise FileNotFoundError(f"ไม่พบ folder ตัวเลขใน {SIR_ROOT}")
    latest = max(folders)
    print(f"📁 พบ folder SIR week ล่าสุด: {latest}")
    return latest


def find_sir_file(week: int) -> Path:
    """หาไฟล์ SIR (.zip หรือ .xlsx) ใน folder ของ week นั้น"""
    folder = Path(SIR_ROOT) / str(week)
    
    # ค้นหาไฟล์ที่ชื่อขึ้นต้นด้วย Official_SIR
    patterns = [
        folder / f"Official_SIR 26_{week} All.zip",
        folder / f"Official_SIR 26_{week} All.xlsx",
        folder / f"Official_SIR*{week}*.zip",
        folder / f"Official_SIR*{week}*.xlsx",
    ]
    
    for pat in patterns[:2]:
        if pat.exists():
            print(f"📄 พบไฟล์ SIR: {pat}")
            return pat
    
    # glob fallback
    for pat in patterns[2:]:
        matches = list(folder.glob(pat.name))
        if matches:
            print(f"📄 พบไฟล์ SIR: {matches[0]}")
            return matches[0]
    
    raise FileNotFoundError(f"ไม่พบไฟล์ Official_SIR ใน {folder}")


def load_sir_dataframe(sir_path: Path) -> pd.DataFrame:
    """โหลด SIR data จาก .zip หรือ .xlsx"""
    
    if sir_path.suffix.lower() == ".zip":
        print("📦 กำลัง unzip ไฟล์ SIR...")
        with zipfile.ZipFile(sir_path, "r") as zf:
            # หาไฟล์ xlsx ข้างใน
            xlsx_names = [n for n in zf.namelist() if n.endswith(".xlsx")]
            if not xlsx_names:
                raise FileNotFoundError("ไม่พบ .xlsx ใน zip")
            
            # เอาไฟล์แรก (หรือที่ชื่อ All)
            target = next(
                (n for n in xlsx_names if "All" in n),
                xlsx_names[0]
            )
            print(f"  → อ่าน: {target}")
            
            with zf.open(target) as f:
                data = io.BytesIO(f.read())
                df = pd.read_excel(data, sheet_name=SIR_SHEET)
    else:
        print("📊 กำลังอ่านไฟล์ SIR xlsx...")
        df = pd.read_excel(sir_path, sheet_name=SIR_SHEET)
    
    print(f"  SIR rows: {len(df):,} | columns: {list(df.columns[:10])}...")
    return df


def load_dispatch_dataframe(dispatch_path: str) -> pd.DataFrame:
    """โหลด Dispatch report จาก .xlsb"""
    print(f"📊 กำลังอ่าน Dispatch file: {dispatch_path}")
    
    rows = []
    headers = None
    
    with pyxlsb.open_workbook(dispatch_path) as wb:
        with wb.get_sheet(1) as sheet:
            for i, row in enumerate(sheet.rows()):
                vals = [c.v for c in row]
                if i == 0:
                    headers = vals
                else:
                    rows.append(vals)
    
    df = pd.DataFrame(rows, columns=headers)
    print(f"  Dispatch rows: {len(df):,} | columns: {len(df.columns)}")
    return df


def merge_data(dispatch_df: pd.DataFrame, sir_df: pd.DataFrame) -> pd.DataFrame:
    """
    Merge: ใช้ Article (คอลัมน์ Y ใน Dispatch) เป็น key
    match กับ Article ใน SIR แล้วดึง SIR_COLUMNS มาต่อท้าย
    """
    print("\n🔗 กำลัง Merge ข้อมูล...")
    
    # ตรวจสอบ column Article ใน SIR
    if "Article" not in sir_df.columns:
        # หา column ที่ใกล้เคียง
        candidates = [c for c in sir_df.columns if "article" in c.lower()]
        if candidates:
            sir_df = sir_df.rename(columns={candidates[0]: "Article"})
            print(f"  ⚠️  ใช้คอลัมน์ '{candidates[0]}' เป็น Article key")
        else:
            raise KeyError(f"ไม่พบคอลัมน์ 'Article' ใน SIR\nคอลัมน์ที่มี: {list(sir_df.columns)}")
    
    # ทำ Article ให้เป็น string เพื่อ match
    dispatch_df["Article"] = dispatch_df["Article"].astype(str).str.strip()
    sir_df["Article"] = sir_df["Article"].astype(str).str.strip()
    
    # เลือกเฉพาะคอลัมน์ที่ต้องการจาก SIR
    available_sir_cols = [c for c in SIR_COLUMNS if c in sir_df.columns]
    missing_sir_cols = [c for c in SIR_COLUMNS if c not in sir_df.columns]
    
    if missing_sir_cols:
        print(f"  ⚠️  คอลัมน์ใน SIR ที่ไม่พบ: {missing_sir_cols}")
    print(f"  ✅ คอลัมน์ที่จะ merge: {available_sir_cols}")
    
    # Deduplicate SIR on Article (เอาแถวแรก)
    sir_lookup = sir_df.drop_duplicates(subset="Article")[["Article"] + available_sir_cols]
    
    # Left merge
    merged = dispatch_df.merge(
        sir_lookup,
        on="Article",
        how="left",
        suffixes=("", "_SIR")
    )
    
    matched = merged[available_sir_cols[0]].notna().sum() if available_sir_cols else 0
    print(f"  Match: {matched:,} / {len(merged):,} rows ({matched/len(merged)*100:.1f}%)")
    
    return merged


def save_output(df: pd.DataFrame, week: int) -> str:
    """บันทึกเป็น .xlsx"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"Dispatch_merged_week{week:02d}_{timestamp}.xlsx"
    output_path = os.path.join(OUTPUT_DIR, filename)
    
    print(f"\n💾 กำลังบันทึก: {output_path}")
    print(f"   จำนวน rows: {len(df):,} | columns: {len(df.columns)}")
    
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Dispatch_Merged")
    
    print(f"✅ บันทึกเสร็จ: {output_path}")
    return output_path


# ============================================================
# 🚀  MAIN
# ============================================================

def main():
    print("=" * 60)
    print("  Dispatch + SIR Merger")
    print(f"  รันเมื่อ: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    print("=" * 60)
    
    # 1. โหลด Dispatch
    if not os.path.exists(DISPATCH_FILE):
        print(f"\n❌ ไม่พบไฟล์ Dispatch: {DISPATCH_FILE}")
        print("   กรุณาแก้ DISPATCH_FILE ในส่วน CONFIG")
        input("\nกด Enter เพื่อออก...")
        return
    
    dispatch_df = load_dispatch_dataframe(DISPATCH_FILE)
    
    # 2. หา SIR week ล่าสุด
    try:
        week = find_latest_sir_week()
        sir_path = find_sir_file(week)
    except FileNotFoundError as e:
        print(f"\n❌ {e}")
        print("   กรุณาแก้ SIR_ROOT ในส่วน CONFIG")
        input("\nกด Enter เพื่อออก...")
        return
    
    # 3. โหลด SIR
    sir_df = load_sir_dataframe(sir_path)
    
    # 4. Merge
    merged_df = merge_data(dispatch_df, sir_df)
    
    # 5. บันทึก
    output_path = save_output(merged_df, week)
    
    print("\n" + "=" * 60)
    print("  ✅ เสร็จสิ้น!")
    print(f"  ไฟล์ผลลัพธ์: {output_path}")
    print("=" * 60)
    input("\nกด Enter เพื่อออก...")


if __name__ == "__main__":
    main()
