"""
Dispatch + SIR Merger — GUI Version
=====================================
ติดตั้ง dependencies ก่อนรันครั้งแรก:
  pip install pyxlsb pandas openpyxl

รันด้วย:
  python dispatch_sir_merger_gui.py
"""

import os
import re
import zipfile
import io
import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

import pandas as pd
import pyxlsb


# SIR columns ที่ต้องดึงมาใส่ใน Dispatch (BF-BM)
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

SIR_SHEET = 0  # sheet แรกใน SIR xlsx


class MergerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Dispatch + SIR Merger")
        self.geometry("720x620")
        self.resizable(True, True)
        self.configure(bg="#1a1a2e")

        # State
        self.dispatch_path = tk.StringVar()
        self.sir_root = tk.StringVar(value=r"O:\2026 Official SIR")
        self.output_dir = tk.StringVar(value=str(Path.home() / "Downloads"))
        self.week_var = tk.StringVar(value="auto")
        self.status_var = tk.StringVar(value="พร้อมใช้งาน")
        self.running = False

        self._build_ui()

    # ───────────────────────────── UI ─────────────────────────────

    def _build_ui(self):
        # Title bar
        title_frame = tk.Frame(self, bg="#16213e", pady=14)
        title_frame.pack(fill="x")
        tk.Label(
            title_frame, text="⚡ Dispatch + SIR Merger",
            font=("Consolas", 18, "bold"),
            fg="#e94560", bg="#16213e"
        ).pack()
        tk.Label(
            title_frame, text="อัปเดตข้อมูล Dispatch ด้วย SIR โดยอัตโนมัติ",
            font=("Consolas", 10), fg="#a8b2d8", bg="#16213e"
        ).pack()

        # Main content
        content = tk.Frame(self, bg="#1a1a2e", padx=24, pady=16)
        content.pack(fill="both", expand=True)

        # Section: Files
        self._section(content, "📁  ไฟล์และโฟลเดอร์")

        row1 = tk.Frame(content, bg="#1a1a2e")
        row1.pack(fill="x", pady=4)
        self._label(row1, "Dispatch File (.xlsb)")
        self._entry_browse(row1, self.dispatch_path, self._browse_dispatch)

        row2 = tk.Frame(content, bg="#1a1a2e")
        row2.pack(fill="x", pady=4)
        self._label(row2, "SIR Root Folder")
        self._entry_browse(row2, self.sir_root, self._browse_sir_root)

        row3 = tk.Frame(content, bg="#1a1a2e")
        row3.pack(fill="x", pady=4)
        self._label(row3, "Output Folder")
        self._entry_browse(row3, self.output_dir, self._browse_output)

        # Section: Week
        self._section(content, "📅  Week SIR")
        week_frame = tk.Frame(content, bg="#1a1a2e")
        week_frame.pack(fill="x", pady=4)
        self._label(week_frame, "Week No.")
        self.week_entry = tk.Entry(
            week_frame, textvariable=self.week_var, width=10,
            bg="#0f3460", fg="#e0e0e0", insertbackground="white",
            font=("Consolas", 11), relief="flat", bd=0,
            highlightthickness=1, highlightcolor="#e94560",
            highlightbackground="#334"
        )
        self.week_entry.pack(side="left", ipady=4, padx=(0, 8))
        tk.Label(
            week_frame, text="(ใส่ตัวเลข เช่น 21, หรือ 'auto' เพื่อหา week ล่าสุด)",
            fg="#a8b2d8", bg="#1a1a2e", font=("Consolas", 9)
        ).pack(side="left")

        # Section: Log
        self._section(content, "📋  Log")
        self.log_box = scrolledtext.ScrolledText(
            content, height=10, bg="#0d0d1a", fg="#a8b2d8",
            font=("Consolas", 9), relief="flat",
            insertbackground="white", wrap="word"
        )
        self.log_box.pack(fill="both", expand=True, pady=(4, 0))
        self.log_box.tag_config("ok", foreground="#50fa7b")
        self.log_box.tag_config("err", foreground="#ff5555")
        self.log_box.tag_config("info", foreground="#8be9fd")
        self.log_box.tag_config("warn", foreground="#ffb86c")

        # Bottom: Run button + status
        bottom = tk.Frame(self, bg="#16213e", padx=24, pady=12)
        bottom.pack(fill="x", side="bottom")

        self.run_btn = tk.Button(
            bottom, text="▶  รัน Merge",
            font=("Consolas", 12, "bold"),
            bg="#e94560", fg="white",
            activebackground="#c73652", activeforeground="white",
            relief="flat", bd=0, padx=24, pady=8,
            cursor="hand2",
            command=self._run_threaded
        )
        self.run_btn.pack(side="left")

        self.progress = ttk.Progressbar(
            bottom, mode="indeterminate", length=200
        )
        self.progress.pack(side="left", padx=16)

        tk.Label(
            bottom, textvariable=self.status_var,
            fg="#a8b2d8", bg="#16213e", font=("Consolas", 9)
        ).pack(side="left")

    def _section(self, parent, text):
        f = tk.Frame(parent, bg="#1a1a2e")
        f.pack(fill="x", pady=(14, 2))
        tk.Label(
            f, text=text, font=("Consolas", 10, "bold"),
            fg="#e94560", bg="#1a1a2e"
        ).pack(side="left")
        tk.Frame(f, bg="#334", height=1).pack(side="left", fill="x", expand=True, padx=(8, 0))

    def _label(self, parent, text):
        tk.Label(
            parent, text=text, width=18, anchor="w",
            fg="#c8d0e0", bg="#1a1a2e", font=("Consolas", 10)
        ).pack(side="left")

    def _entry_browse(self, parent, var, cmd):
        tk.Entry(
            parent, textvariable=var, width=42,
            bg="#0f3460", fg="#e0e0e0", insertbackground="white",
            font=("Consolas", 10), relief="flat", bd=0,
            highlightthickness=1, highlightcolor="#e94560",
            highlightbackground="#334"
        ).pack(side="left", ipady=4, padx=(0, 6))
        tk.Button(
            parent, text="เลือก", command=cmd,
            bg="#0f3460", fg="#e94560", font=("Consolas", 9),
            relief="flat", bd=0, padx=8, pady=2, cursor="hand2",
            activebackground="#16213e", activeforeground="#e94560"
        ).pack(side="left")

    # ───────────────────── Browse handlers ───────────────────────

    def _browse_dispatch(self):
        p = filedialog.askopenfilename(
            title="เลือกไฟล์ Dispatch",
            filetypes=[("Excel Binary", "*.xlsb"), ("All", "*.*")]
        )
        if p:
            self.dispatch_path.set(p)
            # auto set output dir
            self.output_dir.set(str(Path(p).parent))

    def _browse_sir_root(self):
        p = filedialog.askdirectory(title="เลือก SIR Root Folder")
        if p:
            self.sir_root.set(p)

    def _browse_output(self):
        p = filedialog.askdirectory(title="เลือก Output Folder")
        if p:
            self.output_dir.set(p)

    # ─────────────────────── Logging ─────────────────────────────

    def log(self, msg, tag=""):
        ts = datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] {msg}\n"
        self.log_box.insert("end", line, tag)
        self.log_box.see("end")
        self.update_idletasks()

    # ──────────────────────── Run logic ──────────────────────────

    def _run_threaded(self):
        if self.running:
            return
        self.running = True
        self.run_btn.config(state="disabled")
        self.progress.start(12)
        self.status_var.set("กำลังทำงาน...")
        self.log_box.delete("1.0", "end")
        threading.Thread(target=self._run, daemon=True).start()

    def _run(self):
        try:
            self._do_merge()
        except Exception as e:
            self.log(f"❌ Error: {e}", "err")
            messagebox.showerror("Error", str(e))
        finally:
            self.running = False
            self.after(0, lambda: self.run_btn.config(state="normal"))
            self.after(0, self.progress.stop)
            self.status_var.set("เสร็จสิ้น")

    def _do_merge(self):
        dispatch_file = self.dispatch_path.get().strip()
        sir_root = self.sir_root.get().strip()
        output_dir = self.output_dir.get().strip()
        week_str = self.week_var.get().strip()

        # Validate
        if not dispatch_file or not os.path.exists(dispatch_file):
            raise FileNotFoundError(f"ไม่พบไฟล์ Dispatch:\n{dispatch_file}")
        if not os.path.isdir(sir_root):
            raise FileNotFoundError(f"ไม่พบ SIR folder:\n{sir_root}")

        # Find week
        if week_str.lower() == "auto":
            week = self._find_latest_week(sir_root)
        else:
            week = int(week_str)

        self.log(f"📁 SIR Week: {week}", "info")

        # Find SIR file
        sir_file = self._find_sir_file(sir_root, week)
        self.log(f"📄 SIR file: {sir_file}", "info")

        # Load Dispatch
        self.log("📊 กำลังโหลด Dispatch...", "info")
        dispatch_df = self._load_dispatch(dispatch_file)
        self.log(f"   ✅ Dispatch: {len(dispatch_df):,} rows", "ok")

        # Load SIR
        self.log("📊 กำลังโหลด SIR...", "info")
        sir_df = self._load_sir(sir_file)
        self.log(f"   ✅ SIR: {len(sir_df):,} rows", "ok")

        # Merge
        self.log("🔗 กำลัง merge...", "info")
        merged_df = self._merge(dispatch_df, sir_df)

        # Save
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_name = f"Dispatch_merged_week{week:02d}_{timestamp}.xlsx"
        out_path = os.path.join(output_dir, out_name)
        self.log(f"💾 กำลังบันทึก: {out_name}", "info")
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            merged_df.to_excel(writer, index=False, sheet_name="Dispatch_Merged")

        self.log(f"✅ เสร็จสิ้น! บันทึกที่: {out_path}", "ok")
        self.after(0, lambda: messagebox.showinfo(
            "สำเร็จ", f"บันทึกไฟล์เรียบร้อย:\n{out_path}"
        ))

    # ──────────────────────── Helpers ────────────────────────────

    def _find_latest_week(self, sir_root):
        folders = [
            int(p.name)
            for p in Path(sir_root).iterdir()
            if p.is_dir() and p.name.isdigit()
        ]
        if not folders:
            raise FileNotFoundError(f"ไม่พบ folder ตัวเลขใน {sir_root}")
        return max(folders)

    def _find_sir_file(self, sir_root, week):
        folder = Path(sir_root) / str(week)
        if not folder.exists():
            raise FileNotFoundError(f"ไม่พบ folder: {folder}")

        for ext in (".zip", ".xlsx"):
            for f in folder.iterdir():
                if f.suffix.lower() == ext and "Official_SIR" in f.name:
                    return f
        
        raise FileNotFoundError(
            f"ไม่พบไฟล์ Official_SIR ใน {folder}\n"
            f"ไฟล์ที่มี: {[f.name for f in folder.iterdir()]}"
        )

    def _load_dispatch(self, path):
        rows, headers = [], None
        with pyxlsb.open_workbook(path) as wb:
            with wb.get_sheet(1) as sheet:
                for i, row in enumerate(sheet.rows()):
                    vals = [c.v for c in row]
                    if i == 0:
                        headers = vals
                    else:
                        rows.append(vals)
        return pd.DataFrame(rows, columns=headers)

    def _load_sir(self, path):
        if path.suffix.lower() == ".zip":
            with zipfile.ZipFile(path, "r") as zf:
                xlsx_names = [n for n in zf.namelist() if n.endswith(".xlsx")]
                if not xlsx_names:
                    raise FileNotFoundError("ไม่พบ .xlsx ใน zip")
                target = next((n for n in xlsx_names if "All" in n), xlsx_names[0])
                self.log(f"   → unzip: {target}", "")
                with zf.open(target) as f:
                    data = io.BytesIO(f.read())
                    return pd.read_excel(data, sheet_name=SIR_SHEET)
        else:
            return pd.read_excel(path, sheet_name=SIR_SHEET)

    def _merge(self, dispatch_df, sir_df):
        # Normalize Article column
        if "Article" not in sir_df.columns:
            cands = [c for c in sir_df.columns if "article" in c.lower()]
            if cands:
                sir_df = sir_df.rename(columns={cands[0]: "Article"})
                self.log(f"   ⚠️  ใช้คอลัมน์ '{cands[0]}' เป็น Article key", "warn")
            else:
                raise KeyError(
                    f"ไม่พบคอลัมน์ 'Article' ใน SIR\n"
                    f"คอลัมน์ที่มี: {list(sir_df.columns[:15])}"
                )

        dispatch_df["Article"] = dispatch_df["Article"].astype(str).str.strip()
        sir_df["Article"] = sir_df["Article"].astype(str).str.strip()

        available = [c for c in SIR_COLUMNS if c in sir_df.columns]
        missing = [c for c in SIR_COLUMNS if c not in sir_df.columns]

        if missing:
            self.log(f"   ⚠️  คอลัมน์ SIR ที่ไม่พบ: {missing}", "warn")
        self.log(f"   ✅ merge columns: {available}", "ok")

        sir_lookup = sir_df.drop_duplicates(subset="Article")[["Article"] + available]
        merged = dispatch_df.merge(sir_lookup, on="Article", how="left", suffixes=("", "_SIR"))

        matched = merged[available[0]].notna().sum() if available else 0
        pct = matched / len(merged) * 100 if len(merged) else 0
        self.log(f"   Match: {matched:,} / {len(merged):,} rows ({pct:.1f}%)", "ok")

        return merged


# ──────────────────────── Entry point ────────────────────────────

if __name__ == "__main__":
    try:
        import pyxlsb
        import pandas
        import openpyxl
    except ImportError as e:
        import subprocess, sys
        print(f"กำลังติดตั้ง dependencies: {e}")
        subprocess.check_call([sys.executable, "-m", "pip", "install",
                               "pyxlsb", "pandas", "openpyxl"])

    app = MergerApp()
    app.mainloop()
