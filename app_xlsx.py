import io
import os
import re
import json
import time
import glob
import shutil
from datetime import datetime
from typing import Optional, List, Tuple, Dict
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# -------------------------------------------------
# الإعدادات الأساسية (مسار ثابت كما طلبت)
# -------------------------------------------------
EXCEL_PATH = r"D:\Project ILM\Tool store\المخزون.xlsx"
DATA_DIR = os.path.dirname(EXCEL_PATH)
os.makedirs(DATA_DIR, exist_ok=True)
CONFIG_PATH = os.path.join(DATA_DIR, "config.json")

# حقوق المطور
DEV_NAME = "حسان الحربي"

# إعدادات افتراضية
DEFAULT_CONFIG = {
    "global_min_level": 2,
    "enable_backups": False,
    "backup_keep": 0,
    "code_case": "upper",
    "auto_suffix_mode": "by_checkbox",
    "enforce_suffix": False,
    "suffix_text": "-S",
    "suffix_apply_on": ["scan", "bulk", "merge", "ops", "editor", "import"],
    "suffix_apply_on_contexts": ["scan", "bulk", "ops", "stocktake", "add"],
}

SCAN_HISTORY_MAX = 500

# -------------------------------------------------
# واجهة وتنسيق
# -------------------------------------------------
st.set_page_config(page_title="نظام مخزون قطع السيارات (Excel)", layout="wide")
st.markdown(
    """
   <style>
/* =========================================================
🌌 SEVENS NEXT Dashboard — تصميم احترافي فاخر
إصدار 2025 — أسلوب أزرق سماوي أنيق بخط Tajawal
========================================================= */

body {
  direction: rtl;
  text-align: right;
  font-family: 'Tajawal', sans-serif !important;
  background: linear-gradient(135deg, #f7faff 0%, #eef5fb 100%);
  color: #1f2d3d;
  margin: 0;
  padding: 0;
}

/* 🎯 الشريط الجانبي */
[data-testid="stSidebar"] {
  background: linear-gradient(180deg, #0052cc 0%, #00bcd4 100%) !important;
  color: white !important;
  box-shadow: 3px 0 20px rgba(0, 0, 0, 0.15);
}
[data-testid="stSidebar"] * {
  color: #fff !important;
  direction: rtl;
  text-align: right;
  font-size: 15px;
}
[data-testid="stSidebar"] .sidebar-content {
  padding-top: 20px !important;
}

/* 🧭 عناوين القائمة الجانبية */
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 {
  color: #fff !important;
  font-weight: 700;
  text-shadow: 0 2px 5px rgba(0, 0, 0, 0.25);
}

/* 📦 بطاقات المؤشرات */
.metric-box {
  background: linear-gradient(145deg, #ffffff, #f2f6fc);
  border-radius: 18px;
  padding: 25px;
  text-align: center;
  box-shadow: 0 6px 16px rgba(0, 0, 0, 0.06);
  transition: all 0.3s ease;
  border: 1px solid #e3ebf5;
}
.metric-box:hover {
  transform: translateY(-5px);
  box-shadow: 0 10px 25px rgba(0, 0, 0, 0.08);
}
.metric-box h3 {
  color: #007bff;
  margin: 8px 0;
  font-size: 30px;
  font-weight: 800;
}
.metric-box p {
  color: #666;
  font-size: 15px;
  margin: 0;
}

/* 🧩 البطاقات العامة */
.card {
  background: white;
  border-radius: 20px;
  padding: 28px;
  box-shadow: 0 5px 20px rgba(0, 0, 0, 0.05);
  margin-bottom: 25px;
  border: 1px solid #e9eef5;
  transition: all 0.3s ease;
}
.card:hover {
  box-shadow: 0 10px 25px rgba(0, 0, 0, 0.08);
  transform: translateY(-3px);
}

/* 💎 أزرار SEVENS */
.stButton>button, .btn-main {
  background: linear-gradient(90deg, #007bff 0%, #00bcd4 100%);
  color: white !important;
  border: none;
  padding: 10px 28px;
  border-radius: 12px;
  font-weight: 700;
  font-size: 15px;
  letter-spacing: 0.3px;
  transition: all 0.25s;
  box-shadow: 0 4px 12px rgba(0, 123, 255, 0.25);
}
.stButton>button:hover, .btn-main:hover {
  background: linear-gradient(90deg, #00bcd4 0%, #007bff 100%);
  transform: translateY(-2px);
  box-shadow: 0 6px 18px rgba(0, 123, 255, 0.35);
}

/* 🧾 الجداول */
table {
  border-collapse: collapse !important;
  border-radius: 12px;
  overflow: hidden;
}
th {
  background: linear-gradient(90deg, #007bff 0%, #00bcd4 100%) !important;
  color: white !important;
  font-weight: 600;
  font-size: 14px;
  text-align: center !important;
  border: none !important;
}
td {
  text-align: center !important;
  padding: 8px 10px !important;
  border: none !important;
}
tbody tr:nth-child(even) {
  background-color: #f9fbff !important;
}
tbody tr:hover {
  background: #eaf4ff !important;
  transition: 0.25s;
}

/* ✅ صناديق الحالة */
.success-box {
  background: #ecfff6;
  border: 1px solid #a8f5d0;
  color: #0f5132;
  padding: 12px 18px;
  border-radius: 14px;
}
.warn-box {
  background: #fff9e6;
  border: 1px solid #ffe680;
  color: #946200;
  padding: 12px 18px;
  border-radius: 14px;
}
.error-box {
  background: #fff2f2;
  border: 1px solid #ffb3b3;
  color: #991b1b;
  padding: 12px 18px;
  border-radius: 14px;
}

/* 💠 العناوين */
h1, h2, h3, h4 {
  color: #1b2734;
  font-weight: 800;
  letter-spacing: -0.2px;
}
h1 {
  font-size: 28px;
}
h2 {
  font-size: 22px;
}

/* ⚙️ الإدخالات */
input, select, textarea {
  border: 1px solid #cfd8e3 !important;
  border-radius: 12px !important;
  padding: 8px 12px !important;
  font-family: 'Tajawal', sans-serif !important;
  background-color: #fff;
}
input:focus, select:focus, textarea:focus {
  border-color: #00aaff !important;
  box-shadow: 0 0 6px rgba(0, 123, 255, 0.3);
  outline: none !important;
}

/* 🌟 شعار SEVENS */
.logo-box {
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 14px;
  margin-bottom: 25px;
}
.logo-box img {
  height: 48px;
  filter: drop-shadow(0 3px 4px rgba(0,0,0,0.2));
}
.logo-box h1 {
  font-size: 23px;
  font-weight: 800;
  color: #ffffff;
  text-shadow: 0 2px 6px rgba(0,0,0,0.2);
}

/* 🧠 حقوق المطور */
.dev-credit {
  position: fixed;
  bottom: 12px;
  right: 20px;
  background: rgba(0, 123, 255, 0.08);
  padding: 8px 14px;
  border-radius: 10px;
  font-size: 12px;
  color: #007bff;
  z-index: 999;
  backdrop-filter: blur(8px);
}

/* 📱 تجاوب كامل للجوال */
@media (max-width: 768px) {
  .metric-box h3 { font-size: 22px; }
  .metric-box p { font-size: 13px; }
  .card { padding: 18px; }
  .stButton>button { width: 100%; }
  h1 { font-size: 22px; }
}
</style>


    """,
    unsafe_allow_html=True,
)


# -------------------------------------------------
# Helpers عامة
# -------------------------------------------------
def now_iso() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _ts() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def _file_mtime(path: str) -> float:
    try:
        return os.path.getmtime(path)
    except Exception:
        return 0.0


def _safe_int(x, default=0):
    try:
        return int(float(x))
    except Exception:
        return default


def _unique_order(seq: List[str]) -> List[str]:
    return list(dict.fromkeys(seq))


# -------------------------------------------------
# إعداد/قراءة الإعدادات
# -------------------------------------------------
def load_config() -> dict:
    try:
        if os.path.exists(CONFIG_PATH):
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            for k, v in DEFAULT_CONFIG.items():
                cfg.setdefault(k, v)
            cfg["enable_backups"] = False
            cfg["backup_keep"] = 0
            return cfg
    except Exception:
        pass
    return DEFAULT_CONFIG.copy()


def save_config(cfg: dict):
    try:
        cfg["enable_backups"] = False
        cfg["backup_keep"] = 0
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# -------------------------------------------------
# قفل كتابة بسيط + كتابة ذرّية (بدون نسخ احتياطي)
# -------------------------------------------------
class SimpleFileLock:
    def __init__(self, target: str, timeout: float = 5.0, interval: float = 0.1):
        self.lock_path = target + ".lock"
        self.timeout = timeout
        self.interval = interval

    def __enter__(self):
        start = time.time()
        while True:
            try:
                fd = os.open(self.lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                os.close(fd)
                break
            except FileExistsError:
                if time.time() - start > self.timeout:
                    break
                time.sleep(self.interval)

    def __exit__(self, exc_type, exc, tb):
        try:
            if os.path.exists(self.lock_path):
                os.remove(self.lock_path)
        except Exception:
            pass


def _atomic_write_excel(writer_fn, dst_path: str):
    tmp_path = dst_path + ".__tmp__.xlsx"
    writer_fn(tmp_path)
    os.replace(tmp_path, dst_path)


def _backup_if_needed():
    return


def write_all_with_retry(stock: pd.DataFrame, minlvl_unused: pd.DataFrame, tx: pd.DataFrame,
                         retries: int = 3, sleep_s: float = 0.6):
    last_err = None
    for attempt in range(1, retries + 1):
        try:
            write_all(stock, minlvl_unused, tx)
            return
        except Exception as e:
            last_err = e
            time.sleep(sleep_s)
    raise last_err


# -------------------------------------------------
# تهيئة ملف الإكسل
# -------------------------------------------------
def ensure_excel_file():
    if os.path.exists(EXCEL_PATH):
        return
    stock = pd.DataFrame(columns=["الكود", "الوصف", "الموقع", "المخزون"])
    tx = pd.DataFrame(
        columns=["التاريخ", "النوع", "الكود", "الوصف", "من_موقع", "الى_موقع", "الكمية", "المستخدم", "ملاحظة"])

    def _write(p):
        with pd.ExcelWriter(p, engine="openpyxl", mode="w") as w:
            stock.to_excel(w, index=False, sheet_name="Stock")
            tx.to_excel(w, index=False, sheet_name="Transactions")

    with SimpleFileLock(EXCEL_PATH):
        _atomic_write_excel(_write, EXCEL_PATH)


def _drop_sheet_if_exists(path: str, sheet_name: str):
    try:
        if not os.path.exists(path):
            return
        wb = load_workbook(path)
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            wb.remove(ws)
            wb.save(path)
    except Exception:
        pass


# -------------------------------------------------
# منطق لاحقة الأصلي + مُحسّنات المسح (تم تعديله!)
# -------------------------------------------------
CODE_IN_BRACKETS = re.compile(r"\[([^\[\]]+)\]")
CODE_TOKEN = re.compile(r"[0-9A-Za-z\u0600-\u06FF\-_.\/]+")
_AR_NUM_MAP = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")


def _to_ascii_digits(s: str) -> str:
    return (s or "").translate(_AR_NUM_MAP)


def _sanitize_code_input(text: str) -> str:
    s = "" if text is None else str(text)
    s = _to_ascii_digits(s)
    m = CODE_IN_BRACKETS.search(s)
    if m:
        s = m.group(1)
    s = re.sub(r"[^0-9A-Za-z\u0600-\u06FF\-_.\/]", "", s)
    return s.strip()


def _suffix_to_use(cfg: dict) -> str:
    s = str(cfg.get("suffix_text", "-S"))
    cc = cfg.get("code_case", "upper")
    if cc == "upper":
        return s.upper()
    if cc == "lower":
        return s.lower()
    return s


def _normalize_code_text(text: str, cfg: dict, context: str = "") -> str:
    s = ("" if text is None else str(text)).strip()
    s = _to_ascii_digits(s)
    cc = cfg.get("code_case", "upper")
    if cc == "upper":
        s = s.upper()
    elif cc == "lower":
        s = s.lower()
    return s


def _extract_code_from_text(text: str) -> Optional[str]:
    s = _sanitize_code_input(text)
    if not s:
        return None
    m = CODE_TOKEN.search(s)
    return m.group(0).strip() if m else None


# ✅ تم تعديل هذا المنطق بالكامل
def is_original_code(code: str, cfg: dict) -> bool:
    """الكود الأصلي هو الذي لا يحتوي على -S في النهاية."""
    suf = _suffix_to_use(cfg)
    return not str(code or "").strip().endswith(suf)


def ensure_original_flag(code: str, cfg: dict, want_original: bool) -> str:
    """إذا أردنا كودًا أصليًا، نزيل -S. إذا أردنا تقليدًا، نضيف -S."""
    c = (code or "").strip()
    suf = _suffix_to_use(cfg)
    if want_original:
        # نزيل اللاحقة إن وُجدت
        return c[:-len(suf)] if c.endswith(suf) else c
    else:
        # نضمن وجود اللاحقة
        return c if c.endswith(suf) else (c + suf)


def apply_suffix_policy(raw_code: str, cfg: dict, context: str, checkbox_value: Optional[bool]) -> str:
    base = _normalize_code_text(_extract_code_from_text(raw_code) or raw_code, cfg, context=context)
    allowed_ctx = cfg.get("suffix_apply_on_contexts", ["scan", "bulk", "ops", "stocktake", "add"])
    if context not in allowed_ctx:
        return base
    mode = cfg.get("auto_suffix_mode", "by_checkbox")
    suf = _suffix_to_use(cfg)
    if mode == "off":
        return base
    if mode == "always":
        # في وضع "always"، نعتبر أن القيمة الافتراضية هي "أصلي"
        return ensure_original_flag(base, cfg, True)
    # في وضع "by_checkbox"، نستخدم قيمة الزر
    if checkbox_value is None:
        return base
    return ensure_original_flag(base, cfg, bool(checkbox_value))


# -------------------------------------------------
# تحميل أولي للورقة (بدون رؤوس) + اكتشاف الشبكة
# -------------------------------------------------
@st.cache_data(show_spinner=False)
def _load_raw_excel(path: str, _mtime: float) -> dict:
    xls = pd.ExcelFile(path, engine="openpyxl")
    sheets = {}
    for s in xls.sheet_names:
        sheets[s] = pd.read_excel(xls, sheet_name=s, header=None)
    return sheets


def _drop_all_nan(df: pd.DataFrame) -> pd.DataFrame:
    df = df.dropna(axis=1, how="all")
    df = df.dropna(axis=0, how="all")
    return df


def _detect_grid(df_raw: pd.DataFrame) -> pd.DataFrame:
    df = _drop_all_nan(df_raw)
    keep = []
    for c in df.columns:
        name = str(c).strip().lower()
        if name in ["", "nan", "none", "unnamed: 0"]:
            continue
        keep.append(c)
    if keep:
        df = df[keep]
    return df.reset_index(drop=True)


# -------------------------------------------------
# تطبيع Stock / Transactions
# -------------------------------------------------
def _first_row_looks_like_header(df: pd.DataFrame) -> bool:
    try:
        s = df.iloc[0].astype(str).str.strip()
        keywords = ["كود", "وصف", "موقع", "مخزون"]
        hits = sum(any(k in cell for k in keywords) for cell in s)
        return hits >= 2
    except Exception:
        return False


def _heuristic_rebuild_stock(df: pd.DataFrame) -> pd.DataFrame:
    """إعادة بناء الجدول من نص غير منظم."""
    df = df.copy()
    rows = []
    for _, r in df.iterrows():
        cells = [str(r[c]).strip() for c in df.columns]
        # افتراض: العمود 0 = الكود، العمود 1 = الوصف، العمود 2 = الموقع، العمود 3 = المخزون
        code = cells[0] if len(cells) > 0 else ""
        desc = cells[1] if len(cells) > 1 else ""
        loc = cells[2] if len(cells) > 2 else ""
        qty_str = cells[3] if len(cells) > 3 else ""
        try:
            qty = int(float(qty_str)) if qty_str else 0
        except ValueError:
            qty = 0
        # تنظيف الكود من الأقواس إن وُجدت
        if "[" in code and "]" in code:
            code_clean = code.split("[")[1].split("]")[0].strip()
            if code_clean:
                code = code_clean
        # تنظيف الوصف
        if "[" in desc and "]" in desc:
            desc_clean = desc.split("]", 1)[1].strip()
            if desc_clean:
                desc = desc_clean
        rows.append({"الكود": code, "الوصف": desc, "الموقع": loc, "المخزون": qty})
    out = pd.DataFrame(rows).dropna(how="all")
    out["الكود"] = out["الكود"].fillna("").astype(str).str.strip()
    out["الوصف"] = out["الوصف"].fillna("").astype(str).str.strip()
    out["الموقع"] = out["الموقع"].fillna("").astype(str).str.strip()
    out["المخزون"] = pd.to_numeric(out["المخزون"], errors="coerce").fillna(0).astype(int)
    mask_empty = (out[["الكود", "الوصف", "الموقع"]].astype(str).apply(lambda s: s.str.len()) == 0).all(axis=1)
    return out[~mask_empty].reset_index(drop=True)


def _normalize_stock_cols(df: pd.DataFrame) -> pd.DataFrame:
    df0 = df.copy()
    if not df0.empty and _first_row_looks_like_header(df0):
        df0.columns = df0.iloc[0].tolist()
        df0 = df0.iloc[1:].reset_index(drop=True)
    mapping = {}
    for col in df0.columns:
        t = str(col).strip()
        if "كود" in t:
            mapping[col] = "الكود"
        elif "وصف" in t:
            mapping[col] = "الوصف"
        elif "موقع" in t:
            mapping[col] = "الموقع"
        elif "مخزون" in t:
            mapping[col] = "المخزون"
    if mapping:
        df0 = df0.rename(columns=mapping)
        required = ["الكود", "الوصف", "الموقع", "المخزون"]
        if any(c not in df0.columns for c in required):
            df0 = _heuristic_rebuild_stock(df0)
        else:
            df0 = df0.dropna(subset=["الموقع"])
            df0["الكود"] = df0["الكود"].fillna("").astype(str).str.strip()
            df0["الوصف"] = df0["الوصف"].fillna("").astype(str).str.strip()
            df0["الموقع"] = df0["الموقع"].fillna("").astype(str).str.strip()
            df0["المخزون"] = pd.to_numeric(df0["المخزون"], errors="coerce").fillna(0).astype(int)
    else:
        df0 = _heuristic_rebuild_stock(df0)
    return df0[["الكود", "الوصف", "الموقع", "المخزون"]].reset_index(drop=True)


def _normalize_tx_cols(df: pd.DataFrame) -> pd.DataFrame:
    cols = ["التاريخ", "النوع", "الكود", "الوصف", "من_موقع", "الى_موقع", "الكمية", "المستخدم", "ملاحظة"]
    if df.empty:
        return pd.DataFrame(columns=cols)
    if df.iloc[0].astype(str).str.contains("التاريخ|النوع|الكود").any():
        df.columns = df.iloc[0].tolist()
        df = df.iloc[1:]
    mapping = {}
    for col in df.columns:
        t = str(col).strip()
        if "تاريخ" in t:
            mapping[col] = "التاريخ"
        elif "نوع" in t:
            mapping[col] = "النوع"
        elif "كود" in t:
            mapping[col] = "الكود"
        elif "وصف" in t:
            mapping[col] = "الوصف"
        elif "من" in t and "موقع" in t:
            mapping[col] = "من_موقع"
        elif "الى" in t and "موقع" in t:
            mapping[col] = "الى_موقع"
        elif "كمية" in t:
            mapping[col] = "الكمية"
        elif "مستخدم" in t:
            mapping[col] = "المستخدم"
        elif "ملاحظ" in t:
            mapping[col] = "ملاحظة"
    df = df.rename(columns=mapping)
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    df["الكمية"] = pd.to_numeric(df["الكمية"], errors="coerce").fillna(0).astype(int)
    return df[cols].reset_index(drop=True)


# -------------------------------------------------
# قراءة/كتابة موحّدة + تلوين ملف الإكسل بعد كل حفظ
# -------------------------------------------------
def _compact_stock(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    df["الكود"] = df["الكود"].fillna("").astype(str).str.strip()
    df["الموقع"] = df["الموقع"].fillna("").astype(str).str.strip()
    df["الوصف"] = df["الوصف"].fillna("").astype(str).str.strip()
    df["المخزون"] = pd.to_numeric(df["المخزون"], errors="coerce").fillna(0).astype(int)
    out = (df.groupby(["الكود", "الموقع"], as_index=False)
           .agg(المخزون=("المخزون", "sum"), الوصف=("الوصف", "first")))
    return out[["الكود", "الوصف", "الموقع", "المخزون"]].sort_values(["الكود", "الموقع"]).reset_index(drop=True)


def _apply_global_code_normalization(df: pd.DataFrame, context: str):
    cfg = load_config()
    if df.empty:
        return df
    df = df.copy()
    df["الكود"] = df["الكود"].apply(lambda s: _normalize_code_text(s, cfg, context=context))
    return df


def _header_col_index(ws, header_text: str) -> Optional[int]:
    for cell in ws[1]:
        if str(cell.value).strip() == header_text:
            return cell.column
    return None


def _apply_excel_coloring(path: str):
    try:
        cfg = load_config()
        min_level = int(cfg.get("global_min_level", 2))
        suf = _suffix_to_use(cfg)
        wb = load_workbook(path)
        if "Stock" not in wb.sheetnames:
            wb.save(path);
            return
        ws = wb["Stock"]
        c_code = _header_col_index(ws, "الكود")
        c_qty = _header_col_index(ws, "المخزون")
        if not c_code or not c_qty:
            wb.save(path);
            return
        fill_green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        fill_orange = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        fill_red = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        fill_yellow = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
        fill_clear = PatternFill()
        max_row = ws.max_row
        for r in range(2, max_row + 1):
            cell_code = ws.cell(row=r, column=c_code)
            cell_qty = ws.cell(row=r, column=c_qty)
            cell_code.fill = fill_clear
            cell_qty.fill = fill_clear
            code_val = str(cell_code.value or "").strip()
            if code_val:
                # ✅ الآن: الأصلي (بدون -S) = أخضر، التقليد (مع -S) = برتقالي
                if not code_val.endswith(suf):
                    cell_code.fill = fill_green
                else:
                    cell_code.fill = fill_orange
            try:
                q = int(float(cell_qty.value or 0))
            except Exception:
                q = 0
            if q <= 0:
                cell_qty.fill = fill_red
            elif q <= min_level:
                cell_qty.fill = fill_yellow
        wb.save(path)
    except Exception:
        pass


def read_all(preferred_sheet: Optional[str] = None) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, List[str]]:
    ensure_excel_file()
    _drop_sheet_if_exists(EXCEL_PATH, "MinLevels")
    mtime = _file_mtime(EXCEL_PATH)
    sheets_raw = _load_raw_excel(EXCEL_PATH, mtime)
    names = list(sheets_raw.keys())
    candidate = "Stock" if "Stock" in sheets_raw else (preferred_sheet or names[0])
    stock_raw = _detect_grid(sheets_raw[candidate])
    stock = _normalize_stock_cols(stock_raw)
    if "Transactions" in sheets_raw:
        tx = _normalize_tx_cols(_detect_grid(sheets_raw["Transactions"]))
    else:
        tx = pd.DataFrame(
            columns=["التاريخ", "النوع", "الكود", "الوصف", "من_موقع", "الى_موقع", "الكمية", "المستخدم", "ملاحظة"])
    minlvl = pd.DataFrame(columns=["الكود", "حد_إعادة_الطلب"])
    stock = _apply_global_code_normalization(stock, context="import")
    stock = _compact_stock(stock)
    return stock, minlvl, tx, names


def write_all(stock: pd.DataFrame, _minlvl_unused: pd.DataFrame, tx: pd.DataFrame):
    stock = _compact_stock(stock)
    _backup_if_needed()

    def _write(path):
        with pd.ExcelWriter(path, engine="openpyxl", mode="w") as w:
            stock.to_excel(w, index=False, sheet_name="Stock")
            tx.to_excel(w, index=False, sheet_name="Transactions")

    with SimpleFileLock(EXCEL_PATH):
        _atomic_write_excel(_write, EXCEL_PATH)
        _apply_excel_coloring(EXCEL_PATH)


# -------------------------------------------------
# دوال مجال العمل
# -------------------------------------------------
def get_unique_locations(stock: pd.DataFrame) -> List[str]:
    return sorted(stock["الموقع"].astype(str).unique().tolist())


def get_unique_codes(stock: pd.DataFrame) -> List[str]:
    return sorted(stock["الكود"].astype(str).unique().tolist())


def get_part_desc(stock: pd.DataFrame, code: str) -> str:
    m = stock[stock["الكود"] == code]
    return "" if m.empty else str(m["الوصف"].iloc[0])


def get_qty(stock: pd.DataFrame, code: str, location: str) -> int:
    m = stock[(stock["الكود"] == code) & (stock["الموقع"] == location)]
    return 0 if m.empty else int(m["المخزون"].iloc[0])


def get_locations_for_code(stock: pd.DataFrame, code: str) -> List[str]:
    return sorted(stock[stock["الكود"] == code]["الموقع"].unique().tolist())


def set_qty(stock: pd.DataFrame, code: str, location: str, qty: int) -> pd.DataFrame:
    cfg = load_config()
    code = _normalize_code_text(code, cfg, context="ops")
    location = ("" if location is None else str(location)).strip()
    mask = (stock["الكود"] == code) & (stock["الموقع"] == location)
    if mask.any():
        stock.loc[mask, "المخزون"] = int(qty)
    else:
        desc = get_part_desc(stock, code)
        new_row = {"الكود": code, "الوصف": desc, "الموقع": location, "المخزون": int(qty)}
        stock = pd.concat([stock, pd.DataFrame([new_row])], ignore_index=True)
    return stock


def add_qty(stock: pd.DataFrame, code: str, location: str, delta: int) -> Tuple[pd.DataFrame, int]:
    current = get_qty(stock, code, location)
    new_qty = current + delta
    if new_qty < 0:
        raise ValueError("لا يمكن أن تصبح الكمية سالبة.")
    stock = set_qty(stock, code, location, new_qty)
    return stock, new_qty


def append_txn(tx: pd.DataFrame, t_type: str, code: str, desc: str, qty: int,
               from_loc: Optional[str], to_loc: Optional[str],
               user: str = "", note: str = "") -> pd.DataFrame:
    new_row = {
        "التاريخ": now_iso(),
        "النوع": t_type,
        "الكود": code,
        "الوصف": desc,
        "من_موقع": from_loc,
        "الى_موقع": to_loc,
        "الكمية": int(qty),
        "المستخدم": user,
        "ملاحظة": note,
    }
    return pd.concat([tx, pd.DataFrame([new_row])], ignore_index=True)


# -------------------------------------------------
# تنبيهات
# -------------------------------------------------
def compute_low_and_oos(stock: pd.DataFrame, min_level: int) -> Tuple[pd.DataFrame, pd.DataFrame]:
    if stock.empty:
        return (pd.DataFrame(columns=["الكود", "الوصف", "الإجمالي"]),
                pd.DataFrame(columns=["الكود", "الوصف", "الإجمالي"]))
    agg = stock.groupby("الكود", as_index=False).agg(الإجمالي=("المخزون", "sum"),
                                                     الوصف=("الوصف", "first"))
    oos_df = agg[agg["الإجمالي"] <= 0].sort_values("الكود")
    low_df = agg[(agg["الإجمالي"] > 0) & (agg["الإجمالي"] <= int(min_level))].sort_values("الإجمالي")
    return low_df[["الكود", "الوصف", "الإجمالي"]], oos_df[["الكود", "الوصف", "الإجمالي"]]


# -------------------------------------------------
# بحث بسيط (للاستخدام في صفحة البحث) — ✅ مع زر "أصلي؟"
# -------------------------------------------------
def _parse_locations_text(loc_text: str) -> List[str]:
    tokens = [t.strip() for t in re.split(r"[,\n]+", (loc_text or "")) if t.strip()]
    return _unique_order(tokens)


def _apply_search(stock: pd.DataFrame, query_code: str, selected_locs: List[str], cfg: dict,
                  exact_code: bool = True, is_orig: bool = True) -> pd.DataFrame:
    df = stock.copy()
    if selected_locs:
        df = df[df["الموقع"].isin(selected_locs)]
    q = (_to_ascii_digits(query_code or "")).strip()
    if q:
        # ✅ تطبيق سياسة اللاحقة على الكود المدخل في البحث
        norm_q = apply_suffix_policy(q, cfg, context="scan", checkbox_value=is_orig)
        if exact_code:
            df = df[df["الكود"].astype(str).str.strip() == norm_q.strip()]
        else:
            df = df[
                df["الكود"].astype(str).str.contains(norm_q, case=False, na=False) |
                df["الوصف"].astype(str).str.contains(norm_q, case=False, na=False)
                ]
    elif selected_locs:
        # عرض كل القطع في الموقع المحدد (بدون كود)
        pass
    else:
        df = pd.DataFrame(columns=["الكود", "الوصف", "الموقع", "المخزون"])
    return df.reset_index(drop=True)


def _summary_by_code(df: pd.DataFrame, min_level: int) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["الكود", "الوصف", "الإجمالي", "المواقع", "الحالة"])
    out = df.groupby("الكود", as_index=False).agg(
        الإجمالي=("المخزون", "sum"),
        الوصف=("الوصف", "first"),
        المواقع=("الموقع", lambda x: ", ".join(sorted(x.astype(str).unique())))
    )

    def status(q):
        if q <= 0: return "غير متوفر"
        if q <= min_level: return "منخفض"
        return "جيد"

    out["الحالة"] = out["الإجمالي"].apply(status)
    return out.sort_values(["الحالة", "الإجمالي", "الكود"]).reset_index(drop=True)


def _lookup_code(stock: pd.DataFrame, code: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
    df = stock[stock["الكود"] == code].copy()
    if df.empty:
        return df, pd.DataFrame(columns=["الكود", "الوصف", "الإجمالي"])
    s = df.groupby("الكود", as_index=False).agg(الإجمالي=("المخزون", "sum"), الوصف=("الوصف", "first"))
    return df.sort_values("الموقع"), s


# -------------------------------------------------
# تنقّل بين الصفحات
# -------------------------------------------------
PAGES = [
    "لوحة التحكم", "بحث/مسح", "العمليات", "الجرد",
    "إضافة قطعة جديدة",
    "دمج ملف جديد", "تحرير البيانات", "استيراد/تصدير", "إعدادات"
]


def nav_to(page_name: str):
    st.session_state.menu = page_name
    st.rerun()


def top_nav():
    st.markdown("#### التنقل السريع")
    cols = st.columns(len(PAGES))
    for i, p in enumerate(PAGES):
        if cols[i].button(p, key=f"topnav_{p}"):
            nav_to(p)


# -------------------------------------------------
# أدوات حالة الملف
# -------------------------------------------------
def file_status_badge():
    try:
        ok = os.path.exists(EXCEL_PATH)
        size = os.path.getsize(EXCEL_PATH) if ok else 0
        mtime = datetime.fromtimestamp(os.path.getmtime(EXCEL_PATH)).strftime("%Y-%m-%d %H:%M:%S") if ok else "-"
        st.caption(f"📄 ملف البيانات: {'موجود' if ok else 'غير موجود'} | الحجم: {size:,} بايت | آخر تعديل: {mtime}")
    except Exception as e:
        st.caption(f"⚠️ تعذّر قراءة حالة الملف: {e}")


# -------------------------------------------------
# صفحة: بحث/مسح — ✅ مع زر "أصلي؟"
# -------------------------------------------------
def page_find_and_scan():
    st.subheader("بحث داخل المخزن (بدون استلام)")
    top_nav()
    file_status_badge()
    stock, minlvl, tx, _ = read_all()
    cfg = load_config()
    min_level = int(cfg.get("global_min_level", 2))

    loc_text = st.text_input("فلترة بالموقع (يمكن إدخال عدة مواقع بفواصل أو أسطر)", value="", key="simple_loc_text",
                             placeholder="مثال: رف-أ1, صندوق-2")
    selected_locs = _parse_locations_text(loc_text)

    col_a, col_b, col_c = st.columns([3, 3, 1])
    with col_a:
        manual_code = st.text_input("الكود (كتابي)", key="manual_code_input",
                                    placeholder="اكتب الكود أو اتركه فارغًا لعرض كل القطع في الموقع")
    with col_b:
        st.caption("ضع المؤشر هنا ثم امسح الباركود أو أدخل الكود واضغط Enter.")

        def _on_scan():
            raw = st.session_state.get("scanner_code_input", "")
            st.session_state.scanner_code_input = ""
            st.session_state.last_search_code = raw

        st.text_input("الكود (ماسح ضوئي)", key="scanner_code_input", on_change=_on_scan)
    with col_c:
        st.markdown('<div class="orig-checkbox">', unsafe_allow_html=True)
        is_orig = st.checkbox("أصلي؟", value=True, key="search_orig")
        st.markdown('</div>', unsafe_allow_html=True)

    search_code = st.session_state.get("last_search_code", "").strip() or manual_code.strip()
    filtered = _apply_search(stock, search_code, selected_locs, cfg=cfg, exact_code=bool(search_code), is_orig=is_orig)

    if selected_locs and not search_code:
        filtered = stock[stock["الموقع"].isin(selected_locs)].copy().reset_index(drop=True)

    st.markdown("**الملخص حسب الكود:**")
    st.dataframe(_summary_by_code(filtered, min_level), use_container_width=True, height=180)
    st.markdown("**تفاصيل حسب المواقع:**")
    st.dataframe(filtered.sort_values(["الكود", "الموقع"]), use_container_width=True, height=320)
    if not filtered.empty and not search_code and selected_locs:
        st.info(f"عرض جميع القطع في المواقع: {', '.join(selected_locs)}")
    elif search_code and filtered.empty:
        st.info("لا توجد نتائج لهذا الكود ضمن نطاق المواقع المحدد.")


# -------------------------------------------------
# صفحة: الجرد (مُحسّنة جدًا مع التحقق الصحيح من الموقع)
# -------------------------------------------------
def _init_stocktake_state():
    if "stocktake" not in st.session_state:
        st.session_state.stocktake = {
            "scope": "all",
            "loc": "",
            "is_orig": True,
            "items": {},
            "manual_rev": 0,
            "scan_rev": 0,
            "last_code": "",
        }


def _scan_callback(scan_key: str):
    raw = st.session_state.get(scan_key, "")
    st.session_state.stocktake["last_code"] = raw
    st.session_state.stocktake["scan_rev"] += 1
    st.rerun()


def _clear_inputs_and_rerun():
    st.session_state.stocktake["last_code"] = ""
    st.session_state.stocktake["manual_rev"] += 1
    st.session_state.stocktake["scan_rev"] += 1
    st.rerun()


def page_stocktake():
    st.subheader("الجرد المبسّط")
    top_nav()
    file_status_badge()
    _init_stocktake_state()
    cfg = load_config()
    stock, minlvl, tx, _ = read_all()
    min_level = int(cfg.get("global_min_level", 2))

    c1, c2 = st.columns([2, 2])
    with c1:
        scope = st.radio("نطاق الجرد", ["المخزن كامل", "حسب موقع محدد"], horizontal=True,
                         index=0 if st.session_state.stocktake["scope"] == "all" else 1)
        st.session_state.stocktake["scope"] = "all" if scope == "المخزن كامل" else "loc"
    with c2:
        if st.session_state.stocktake["scope"] == "loc":
            loc_input = st.text_input("الموقع (كتابي)", value=st.session_state.stocktake.get("loc", ""),
                                      placeholder="مثال: رف-أ1")
            st.session_state.stocktake["loc"] = loc_input
        else:
            st.text_input("الموقع (معطّل في وضع المخزن كامل)", value="", disabled=True)

    c3, c4, c5 = st.columns([1, 3, 3])
    with c3:
        st.markdown('<div class="orig-checkbox">', unsafe_allow_html=True)
        is_orig = st.checkbox("أصلي؟", value=st.session_state.stocktake.get("is_orig", True), key="stocktake_orig")
        st.markdown('</div>', unsafe_allow_html=True)
        st.session_state.stocktake["is_orig"] = is_orig
    with c4:
        manual_key = f"stk_manual_code_{st.session_state.stocktake['manual_rev']}"
        manual_code = st.text_input("الكود (كتابي)", key=manual_key, placeholder="اكتب الكود أو امسح الباركود")
    with c5:
        st.caption("امسح الباركود هنا ثم اضغط Enter")
        scan_key = f"stk_scanner_code_{st.session_state.stocktake['scan_rev']}"
        st.text_input("الكود (ماسح ضوئي)", key=scan_key, on_change=_scan_callback, args=(scan_key,))

    qty = st.number_input("العدد الفعلي", min_value=0, value=0, step=1, key="stk_count_simple")

    if st.button("إضافة إلى سلة الجرد"):
        raw = (st.session_state.stocktake.get("last_code", "") or manual_code or "").strip()
        if not raw:
            st.warning("أدخل الكود أولًا.")
        else:
            code_with_suffix = apply_suffix_policy(raw, cfg, context="stocktake",
                                                   checkbox_value=st.session_state.stocktake["is_orig"])
            code_normalized = _normalize_code_text(code_with_suffix, cfg, context="stocktake")

            # --- 🔑 التحقق المرن من الموقع (الكود مع وبدون -S) ---
            all_codes_in_system = set(stock["الكود"].astype(str))
            suf = _suffix_to_use(cfg)

            # إنشاء قائمة بالمرشحات: الكود كما هو، بدون -S، مع -S
            candidates = {code_normalized}
            if code_normalized.endswith(suf):
                candidates.add(code_normalized[:-len(suf)])
            else:
                candidates.add(code_normalized + suf)

            # البحث عن أي تطابق
            matched_rows = stock[stock["الكود"].isin(candidates)]
            sys_locs = sorted(matched_rows["الموقع"].unique().tolist())
            # --- نهاية التحقق المرن ---

            loc_entered = st.session_state.stocktake["loc"].strip() if st.session_state.stocktake[
                                                                           "scope"] == "loc" else None

            if st.session_state.stocktake["scope"] == "loc":
                if not loc_entered:
                    st.error("أدخل الموقع أولًا أو بدّل إلى 'المخزن كامل'.")
                else:
                    # تنبيه إذا كان الموقع غير موجود في النظام لهذا الكود
                    if sys_locs and loc_entered not in sys_locs:
                        st.warning(
                            f"⚠️ الموقع '{loc_entered}' غير مسجل لهذا الكود في النظام. المواقع المسجلة: {', '.join(sys_locs)}")
                    key = (code_normalized, loc_entered)
                    sys_qty = matched_rows[matched_rows["الموقع"] == loc_entered][
                        "المخزون"].sum() if not matched_rows.empty else 0
                    row = st.session_state.stocktake["items"].get(key, {"count": 0, "sys_qty": sys_qty})
                    row["count"] = int(qty)
                    row["sys_qty"] = sys_qty
                    st.session_state.stocktake["items"][key] = row
                    st.success(f"أُضيف: {code_normalized} @ {loc_entered} | فعلي: {qty} | نظام: {sys_qty}")
                    _clear_inputs_and_rerun()
            else:
                # نطاق المخزن كامل
                sys_total = matched_rows["المخزون"].sum() if not matched_rows.empty else 0
                key = (code_normalized, None)
                row = st.session_state.stocktake["items"].get(key, {"count": 0, "sys_qty": sys_total})
                row["count"] = int(qty)
                row["sys_qty"] = sys_total
                st.session_state.stocktake["items"][key] = row
                st.success(f"أُضيف: {code_normalized} (المخزن كامل) | فعلي: {qty} | نظام: {sys_total}")
                _clear_inputs_and_rerun()

    st.markdown("### سلة الجرد")
    basket_rows = []
    for (code, loc), data in st.session_state.stocktake["items"].items():
        basket_rows.append({
            "الكود": code,
            "النوع": "أصلي" if is_original_code(code, cfg) else "تجاري",
            "الموقع": ("المخزن كامل" if loc is None else loc),
            "العدد الفعلي": int(data.get("count", 0)),
            "عدد النظام": int(data.get("sys_qty", 0)),
        })
    basket_df = pd.DataFrame(basket_rows) if basket_rows else pd.DataFrame(
        columns=["الكود", "النوع", "الموقع", "العدد الفعلي", "عدد النظام"])
    st.dataframe(basket_df.sort_values(["الكود", "الموقع"]), use_container_width=True, height=280)

    col_clear, col_apply = st.columns(2)
    with col_clear:
        if st.button("تفريغ السلة"):
            st.session_state.stocktake["items"] = {}
            st.success("تم تفريغ سلة الجرد.")
    with col_apply:
        if st.button("تطبيق التسوية"):
            if not st.session_state.stocktake["items"]:
                st.warning("سلة الجرد فارغة.")
            else:
                try:
                    stock_cur, minlvl_cur, tx_cur, _ = read_all()
                    DEFAULT_LOC_FOR_NEW = "MAIN"
                    adjustments = 0
                    for (code, loc), data in st.session_state.stocktake["items"].items():
                        actual = int(data.get("count", 0))
                        sys_qty = int(data.get("sys_qty", 0))
                        delta = actual - sys_qty
                        if delta == 0:
                            continue
                        if loc is not None:
                            stock_cur, new_qty = add_qty(stock_cur, code, loc, delta)
                            tx_cur = append_txn(
                                tx_cur, "ADJUST", code, get_part_desc(stock_cur, code),
                                abs(delta),
                                loc if delta < 0 else None,
                                loc if delta > 0 else None,
                                "STOCKTAKE", "تسوية جرد (حسب موقع)"
                            )
                            adjustments += 1
                        else:
                            existing_rows = stock_cur[stock_cur["الكود"] == code].sort_values("المخزون",
                                                                                              ascending=False)
                            target_loc = str(
                                existing_rows["الموقع"].iloc[0]) if not existing_rows.empty else DEFAULT_LOC_FOR_NEW
                            stock_cur, new_qty = add_qty(stock_cur, code, target_loc, delta)
                            tx_cur = append_txn(
                                tx_cur, "ADJUST", code, get_part_desc(stock_cur, code),
                                abs(delta),
                                target_loc if delta < 0 else None,
                                target_loc if delta > 0 else None,
                                "STOCKTAKE", "تسوية جرد (المخزن كامل)"
                            )
                            adjustments += 1
                    write_all_with_retry(stock_cur, minlvl_cur, tx_cur)
                    st.cache_data.clear()
                    st.success(f"تم تطبيق التسوية. عدد القطع المسوّاة: {adjustments}.")
                except Exception as e:
                    st.error(f"فشل تطبيق التسوية: {e}")


# -------------------------------------------------
# باقي الصفحات (بدون تغيير جوهري لأنها تعمل جيدًا)
# -------------------------------------------------
def _exists_pair(stock: pd.DataFrame, code: str, loc: str) -> bool:
    return ((stock["الكود"] == code) & (stock["الموقع"] == loc)).any()


def page_add_new_item():
    st.subheader("إضافة قطعة جديدة")
    top_nav()
    file_status_badge()
    cfg = load_config()
    stock, minlvl, tx, _ = read_all()
    with st.form("add_item_form_simple_ordered", clear_on_submit=False):
        col_qty, col_loc = st.columns(2)
        with col_qty:
            qty = st.number_input("الكمية", min_value=0, value=0, step=1, key="add_qty")
        with col_loc:
            loc = st.text_input("الموقع", placeholder="مثال: رف-أ1", key="add_loc")
        desc = st.text_input("الوصف", placeholder="وصف القطعة", key="add_desc")
        col_code, col_orig = st.columns([3, 1])
        with col_code:
            raw_code = st.text_input("الكود", placeholder="مثال: ABC-123 أو ABC-123-S", key="add_code")
        with col_orig:
            is_orig = st.checkbox("أصلي؟", value=True, key="add_isorig")
        submitted = st.form_submit_button("إضافة / زيادة")
    if submitted:
        try:
            if not raw_code.strip():
                st.error("الرجاء إدخال الكود.")
                return
            if not loc.strip():
                st.error("الرجاء إدخال الموقع.")
                return
            norm_code = apply_suffix_policy(raw_code, cfg, context="add", checkbox_value=is_orig)
            norm_code = _normalize_code_text(norm_code, cfg, context="add")
            loc = loc.strip()
            qty = int(qty)
            stock_cur, minlvl_cur, tx_cur, _ = read_all()
            if _exists_pair(stock_cur, norm_code, loc):
                current = get_qty(stock_cur, norm_code, loc)
                stock_cur, new_qty = add_qty(stock_cur, norm_code, loc, qty)
                if str(desc).strip():
                    mask = (stock_cur["الكود"] == norm_code) & (stock_cur["الموقع"] == loc)
                    cur_desc = str(stock_cur.loc[mask, "الوصف"].iloc[0]) if mask.any() else ""
                    if (cur_desc is None) or (str(cur_desc).strip() == ""):
                        stock_cur.loc[mask, "الوصف"] = desc.strip()
                if qty > 0:
                    tx_cur = append_txn(tx_cur, "RECEIVE", norm_code,
                                        get_part_desc(stock_cur, norm_code) or desc.strip() or norm_code, qty, None,
                                        loc, user="ADD", note="Add-page increment")
                write_all_with_retry(stock_cur, minlvl_cur, tx_cur)
                st.cache_data.clear()
                st.success(f"تمت الزيادة: {norm_code} @ {loc} | {current} → {new_qty}")
            else:
                new_row = {"الكود": norm_code, "الوصف": desc.strip(), "الموقع": loc, "المخزون": int(qty)}
                stock_cur = pd.concat([stock_cur, pd.DataFrame([new_row])], ignore_index=True)
                if qty > 0:
                    tx_cur = append_txn(tx_cur, "RECEIVE", norm_code, desc.strip() or norm_code, qty, None, loc,
                                        user="ADD", note="Add-page create")
                write_all_with_retry(stock_cur, minlvl_cur, tx_cur)
                st.cache_data.clear()
                st.success(f"تمت الإضافة: {norm_code} @ {loc} بكمية {qty}")
            details, summary = _lookup_code(stock_cur, norm_code)
            if not details.empty:
                st.markdown("**الوضع الحالي للقطعة:**")
                st.dataframe(details.sort_values("الموقع"), use_container_width=True, height=180)
        except Exception as e:
            st.error(f"فشل الإضافة/الزيادة: {e}")


def _uploaded_sheets(file) -> List[str]:
    file.seek(0)
    xls = pd.ExcelFile(file, engine="openpyxl")
    return xls.sheet_names


def _read_uploaded_stock(file, sheet_name: str) -> pd.DataFrame:
    file.seek(0)
    xls = pd.ExcelFile(file, engine="openpyxl")
    raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    df = _normalize_stock_cols(_detect_grid(raw))
    return _apply_global_code_normalization(df, context="merge")


def _make_diff(base: pd.DataFrame, incoming: pd.DataFrame, mode: str, only_new: bool) -> pd.DataFrame:
    base_key = base.assign(_key=base["الكود"].astype(str) + "||" + base["الموقع"].astype(str))
    inc_key = incoming.assign(_key=incoming["الكود"].astype(str) + "||" + incoming["الموقع"].astype(str))
    m = base_key.merge(inc_key, on="_key", how="outer", suffixes=("_current", "_new"))
    if only_new:
        m = m[m["الكود_current"].isna()]

    def _final_qty(row):
        cur = _safe_int(row.get("المخزون_current"), 0)
        new = _safe_int(row.get("المخزون_new"), 0)
        if pd.isna(row.get("الكود_new")) and pd.isna(row.get("الموقع_new")):
            return cur
        return cur + new if mode == "add" else new

    out = pd.DataFrame({
        "الكود": m["الكود_new"].fillna(m["الكود_current"]).astype(str),
        "الموقع": m["الموقع_new"].fillna(m["الموقع_current"]).astype(str),
        "الوصف_حالي": m["الوصف_current"],
        "الوصف_جديد": m["الوصف_new"],
        "كمية_حالية": m["المخزون_current"].fillna(0).astype(int),
        "كمية_قادمة": m["المخزون_new"].fillna(0).astype(int),
    })
    out["الكمية_بعد_الدمج"] = m.apply(_final_qty, axis=1).astype(int)

    def _action(row):
        if row["كمية_حالية"] == row["الكمية_بعد_الدمج"]:
            return "بدون تغيير"
        if row["كمية_حالية"] == 0 and row["كمية_قادمة"] > 0 and (
                pd.isna(row["الوصف_حالي"]) or str(row["الوصف_حالي"]).strip() == ""):
            return "إضافة صف جديد"
        return "تحديث كمية"

    out["الإجراء"] = out.apply(_action, axis=1)
    return out.sort_values(["الإجراء", "الكود", "الموقع"]).reset_index(drop=True)


def _apply_merge(base: pd.DataFrame, incoming: pd.DataFrame, mode: str,
                 desc_policy: str, only_new: bool) -> Tuple[pd.DataFrame, int, int]:
    updated, added = 0, 0
    result = base.copy()
    if only_new:
        mask = ~incoming.set_index(["الكود", "الموقع"]).index.isin(
            base.set_index(["الكود", "الموقع"]).index
        )
        incoming = incoming[mask].copy()
    for _, r in incoming.iterrows():
        code = str(r["الكود"]).strip()
        loc = str(r["الموقع"]).strip()
        qty_new = int(r["المخزون"])
        desc_new = str(r.get("الوصف", "")).strip()
        if mode == "add":
            cur = get_qty(result, code, loc)
            result, new_qty = add_qty(result, code, loc, qty_new)
            if (code, loc) in set(base.set_index(["الكود", "الموقع"]).index):
                if new_qty != cur:
                    updated += 1
            else:
                added += 1
        else:
            existed = ((result["الكود"] == code) & (result["الموقع"] == loc)).any()
            result = set_qty(result, code, loc, qty_new)
            if existed:
                updated += 1
            else:
                added += 1
        mask = (result["الكود"] == code) & (result["الموقع"] == loc)
        cur_desc = str(result.loc[mask, "الوصف"].iloc[0]) if mask.any() else ""
        if desc_policy == "replace":
            result.loc[mask, "الوصف"] = desc_new
        elif desc_policy == "fill_blank":
            if (cur_desc is None) or (str(cur_desc).strip() == ""):
                result.loc[mask, "الوصف"] = desc_new
    return result, updated, added


def page_merge():
    st.subheader("دمج ملف جديد مع الملف الحالي")
    top_nav()
    file_status_badge()
    base_stock, minlvl, tx, _ = read_all()
    st.caption(f"الأكواد الحالية: {base_stock['الكود'].nunique():,} | الصفوف: {len(base_stock):,}")
    up = st.file_uploader("اختر ملف Excel اليومي للقطع الجديدة", type=["xlsx", "xls"])
    if not up:
        st.info("ارفع ملفًا للبدء.")
        return
    try:
        sheets = _uploaded_sheets(up)
    except Exception as e:
        st.error(f"تعذّر قراءة الملف: {e}")
        return
    sheet = st.selectbox("اختر الورقة داخل الملف:", options=sheets)
    try:
        incoming = _read_uploaded_stock(up, sheet)
    except Exception as e:
        st.error(f"فشل التعرف على الأعمدة داخل الورقة المحددة: {e}")
        return
    cfg = load_config()
    if not incoming.empty:
        incoming["الكود"] = incoming["الكود"].apply(lambda s: _normalize_code_text(s, cfg, context="merge"))
    st.success(f"تم التعرف على {len(incoming)} صفًا من الملف الجديد. معاينة:")
    st.dataframe(incoming.head(30), use_container_width=True, height=240)
    st.markdown("### إعدادات الدمج")
    c1, c2, c3 = st.columns(3)
    with c1:
        mode = st.radio("إستراتيجية الكمية", ["استبدال الكمية (Set)", "إضافة على الكمية (Add)"], horizontal=False)
        mode_key = "set" if mode.startswith("استبدال") else "add"
    with c2:
        desc_policy = st.selectbox("سياسة الوصف", ["لا تغيّر الوصف الحالي", "حدّث الوصف إذا كان الحالي فارغًا",
                                                   "استبدل الوصف دائمًا بالقادم"])
        desc_key = {"لا تغيّر الوصف الحالي": "keep", "حدّث الوصف إذا كان الحالي فارغًا": "fill_blank",
                    "استبدل الوصف دائمًا بالقادم": "replace"}[desc_policy]
    with c3:
        only_new = st.checkbox("استيراد الأكواد/المواقع الجديدة فقط", value=False)
    st.markdown("### المعاينة قبل الحفظ (Diff)")
    diff_df = _make_diff(base_stock, incoming, mode_key, only_new)
    st.dataframe(diff_df, use_container_width=True, height=320)
    add_count = (diff_df["الإجراء"] == "إضافة صف جديد").sum()
    upd_count = (diff_df["الإجراء"] == "تحديث كمية").sum()
    nochg_count = (diff_df["الإجراء"] == "بدون تغيير").sum()
    st.caption(f"إحصائيات: جديد: {add_count} | تحديث: {upd_count} | بدون تغيير: {nochg_count}")
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        diff_df.to_excel(w, index=False, sheet_name="Diff")
        incoming.to_excel(w, index=False, sheet_name="Incoming")
        base_stock.to_excel(w, index=False, sheet_name="Current")
    st.download_button("تنزيل تقرير المقارنة (Excel)", data=out.getvalue(),
                       file_name="تقرير_دمج_المخزون.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if st.button("تنفيذ الدمج والحفظ داخل الملف الحالي"):
        try:
            merged, updated, added = _apply_merge(base_stock, incoming, mode_key, desc_key, only_new)
            tx = append_txn(tx, "ADJUST", "BULK_MERGE", "دمج ملف يومي", int(len(incoming)), None, None, user="MERGE",
                            note=f"mode={mode_key}, desc={desc_key}, only_new={only_new}")
            write_all_with_retry(merged, minlvl, tx)
            st.cache_data.clear()
            st.success(f"تم الدمج بنجاح. تمت إضافة {added} وتحديث {updated} صفًا.")
            if st.button("الانتقال إلى لوحة التحكم"):
                nav_to("لوحة التحكم")
        except Exception as e:
            st.error(f"فشل الدمج: {e}")


def page_data_editor():
    st.subheader("تحرير البيانات مباشرة (Stock)")
    top_nav()
    file_status_badge()
    stock, minlvl, tx, _ = read_all()
    edited_stock = st.data_editor(
        stock,
        key="stock_editor",
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "الكود": st.column_config.TextColumn(required=True),
            "الوصف": st.column_config.TextColumn(required=False),
            "الموقع": st.column_config.TextColumn(required=True),
            "المخزون": st.column_config.NumberColumn(min_value=0, step=1),
        },
    )
    if st.button("حفظ التغييرات"):
        try:
            cfg = load_config()
            edited_stock["الكود"] = edited_stock["الكود"].apply(
                lambda s: _normalize_code_text(s, cfg, context="editor"))
            edited_stock["الموقع"] = edited_stock["الموقع"].fillna("").astype(str).str.strip()
            edited_stock["الوصف"] = edited_stock["الوصف"].fillna("").astype(str).str.strip()
            edited_stock["المخزون"] = pd.to_numeric(edited_stock["المخزون"], errors="coerce").fillna(0).astype(int)
            write_all_with_retry(edited_stock, minlvl, tx)
            st.cache_data.clear()
            st.success("تم الحفظ داخل نفس الملف.")
        except Exception as e:
            st.error(f"خطأ أثناء الحفظ: {e}")


def page_operations():
    st.subheader("العمليات (صرف / تحويل)")
    top_nav()
    file_status_badge()
    stock, minlvl, tx, _ = read_all()
    cfg = load_config()
    min_level = int(cfg.get("global_min_level", 2))
    codes_list = get_unique_codes(stock)
    locs_list = get_unique_locations(stock)
    op = st.selectbox("اختر العملية", ["صرف (ISSUE)", "تحويل (TRANSFER)"])
    mode_code = st.radio("طريقة إدخال الكود", ["كتابي", "قائمة"], horizontal=True, index=0)
    mode_loc = st.radio("طريقة إدخال الموقع", ["كتابي", "قائمة"], horizontal=True, index=0)

    def input_code(label_key: str):
        if mode_code == "قائمة" and codes_list:
            return st.selectbox(label_key, options=codes_list, key=label_key + "_select"), None
        cols = st.columns([3, 1])
        with cols[0]:
            raw = st.text_input(label_key, key=label_key + "_text", placeholder="امسح الباركود أو اكتب الكود")
        with cols[1]:
            orig = st.checkbox("أصلي؟", value=True, key=label_key + "_isorig")
        return raw, orig

    def input_loc(label_key: str):
        if mode_loc == "قائمة" and locs_list:
            return st.selectbox(label_key, options=locs_list, key=label_key + "_select")
        return st.text_input(label_key, key=label_key + "_text")

    def preview_qty(code: str, loc: Optional[str] = None):
        if not code:
            return
        details, summary = _lookup_code(stock, code)
        if details.empty:
            st.info("هذا الكود غير موجود في المخزون.")
            return
        with st.expander("عرض سريع للمخزون الحالي لهذا الكود", expanded=False):
            st.dataframe(details.sort_values("الموقع"), use_container_width=True, height=160)
            if loc:
                st.caption(f"الكمية في [{loc}]: {get_qty(stock, code, loc)}")

    with st.form("ops_form"):
        col1, col2 = st.columns(2)
        with col1:
            code_raw, isorig = input_code("الكود")
        with col2:
            if mode_code == "قائمة":
                norm_code = _normalize_code_text(code_raw or "", cfg, context="ops")
            else:
                norm_code = apply_suffix_policy(code_raw or "", cfg, context="ops", checkbox_value=isorig)
            desc_default = get_part_desc(stock, norm_code) if norm_code else ""
            description = st.text_input("الوصف (اختياري)", value=desc_default)
        qty = st.number_input("الكمية", min_value=1, value=1, step=1)
        note = st.text_input("ملاحظة")
        user = st.text_input("المستخدم (اختياري)")
        if op == "صرف (ISSUE)":
            from_loc = input_loc("من موقع")
            if norm_code and from_loc:
                preview_qty(norm_code, from_loc)
            submitted = st.form_submit_button("تنفيذ الصرف")
            if submitted:
                if not norm_code or not from_loc:
                    st.error("أدخل الكود وموقع الصرف.")
                else:
                    current = get_qty(stock, norm_code, from_loc)
                    if current <= 0:
                        st.markdown(
                            f"<div class='error-box'>❌ الكود <b>{norm_code}</b> غير موجود/صفر في {from_loc}.</div>",
                            unsafe_allow_html=True)
                    elif int(qty) > current:
                        st.markdown(
                            f"<div class='error-box'>❌ الكمية المطلوبة ({int(qty)}) أكبر من المتاح ({current}) في {from_loc}.</div>",
                            unsafe_allow_html=True)
                    else:
                        try:
                            stock, new_qty = add_qty(stock, norm_code, from_loc, -int(qty))
                        except ValueError as e:
                            st.error(str(e))
                        else:
                            tx = append_txn(tx, "ISSUE", norm_code, get_part_desc(stock, norm_code), int(qty), from_loc,
                                            None, user, note)
                            write_all_with_retry(stock, minlvl, tx)
                            st.success(f"تم الصرف. الكمية الحالية في {from_loc}: {new_qty}")
                            if new_qty == 0:
                                st.error("⚠️ نفدت الكمية لهذا الكود في هذا الموقع.")
                            elif new_qty <= min_level:
                                st.warning(f"⚠️ الكمية منخفضة (≤ {min_level}).")
        elif op == "تحويل (TRANSFER)":
            c1, c2 = st.columns(2)
            with c1:
                from_loc = input_loc("من موقع")
            with c2:
                to_loc = input_loc("إلى موقع")
            if norm_code and from_loc:
                preview_qty(norm_code, from_loc)
            submitted = st.form_submit_button("تنفيذ التحويل")
            if submitted:
                if not norm_code or not from_loc or not to_loc:
                    st.error("أدخل الكود والموقعين.")
                elif from_loc == to_loc:
                    st.error("اختر موقعين مختلفين.")
                else:
                    current = get_qty(stock, norm_code, from_loc)
                    if current <= 0:
                        st.error("لا توجد كمية لهذا الكود في موقع التحويل (من).")
                    elif int(qty) > current:
                        st.error(f"الكمية المطلوبة ({int(qty)}) أكبر من المتاح ({current}) في {from_loc}.")
                    else:
                        try:
                            stock, new_from = add_qty(stock, norm_code, from_loc, -int(qty))
                        except ValueError as e:
                            st.error(str(e))
                        else:
                            stock, new_to = add_qty(stock, norm_code, to_loc, int(qty))
                            tx = append_txn(tx, "TRANSFER", norm_code, get_part_desc(stock, norm_code), int(qty),
                                            from_loc, to_loc, user, note)
                            write_all_with_retry(stock, minlvl, tx)
                            st.success(f"تم التحويل. المتبقي في {from_loc}: {new_from} — الحالي في {to_loc}: {new_to}")
                            if new_from == 0:
                                st.error(f"⚠️ نفدت الكمية في {from_loc}.")
                            elif new_from <= min_level:
                                st.warning(f"⚠️ الكمية منخفضة في {from_loc} (≤ {min_level}).")


def page_import_export():
    st.subheader("استيراد / تصدير")
    top_nav()
    file_status_badge()
    stock, minlvl, tx, _ = read_all()
    st.markdown("### تنزيل نسخة عمل (Stock + Transactions)")
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        stock.to_excel(w, index=False, sheet_name="Stock")
        tx.to_excel(w, index=False, sheet_name="Transactions")
    st.download_button(
        "تحميل المخزون_الحالي.xlsx",
        data=out.getvalue(),
        file_name="المخزون_الحالي.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.markdown("---")
    st.caption("لعملية دمج متقدمة استخدم صفحة 'دمج ملف جديد' من القائمة.")


def page_settings():
    st.subheader("إعدادات")
    top_nav()
    st.caption(f"المسار الحالي لملف البيانات: {EXCEL_PATH}")
    file_status_badge()
    cfg = load_config()
    colA, colB = st.columns([2, 2])
    with colA:
        min_level = st.number_input("الحد الأدنى الافتراضي للتنبيه (إعادة الطلب)", min_value=0,
                                    value=int(cfg.get("global_min_level", 2)), step=1)
        code_case = st.selectbox("تطبيع حروف الكود", ["upper", "lower", "none"],
                                 index=["upper", "lower", "none"].index(cfg.get("code_case", "upper")))
    with colB:
        auto_suffix_mode = st.selectbox("منطق اللاحقة -S (تمييز الأصلي)", ["by_checkbox", "always", "off"],
                                        index=["by_checkbox", "always", "off"].index(
                                            cfg.get("auto_suffix_mode", "by_checkbox")))
        suffix_text = st.text_input("نص اللاحقة للأصلي", value=str(cfg.get("suffix_text", "-S")))
        apply_on = st.multiselect("تطبيق التطبيع الحرفي عند",
                                  options=["scan", "bulk", "merge", "ops", "editor", "import", "add"],
                                  default=_unique_order(cfg.get("suffix_apply_on",
                                                                ["scan", "bulk", "merge", "ops", "editor", "import",
                                                                 "add"])))
    contexts_all = ["scan", "bulk", "ops", "stocktake", "merge", "editor", "import", "add"]
    suffix_ctx = st.multiselect(
        "تطبيق منطق اللاحقة -S في هذه السياقات",
        options=contexts_all,
        default=_unique_order(cfg.get("suffix_apply_on_contexts", ["scan", "bulk", "ops", "stocktake", "add"]))
    )
    if st.button("حفظ الإعدادات"):
        cfg["global_min_level"] = int(min_level)
        cfg["code_case"] = code_case
        cfg["auto_suffix_mode"] = auto_suffix_mode
        cfg["suffix_text"] = suffix_text
        cfg["suffix_apply_on"] = apply_on
        cfg["suffix_apply_on_contexts"] = suffix_ctx
        cfg["enable_backups"] = False
        cfg["backup_keep"] = 0
        save_config(cfg)
        st.success("تم حفظ الإعدادات.")
        st.cache_data.clear()
    try:
        _, _, _, names = read_all()
        st.write("الأوراق الموجودة الآن:", names)
    except Exception:
        pass
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        if st.button("تحديث الكاش / إعادة التحميل"):
            st.cache_data.clear()
            st.success("تم مسح الكاش. أعد تحميل الصفحة (Ctrl+F5).")
    with c2:
        if st.button("حذف ورقة MinLevels (إن وُجدت)"):
            _drop_sheet_if_exists(EXCEL_PATH, "MinLevels")
            st.cache_data.clear()
            st.success("تم حذف ورقة MinLevels.")
    with c3:
        if st.button("تجميع المكررات وحفظ"):
            stock, minlvl, tx, _ = read_all()
            stock2 = _compact_stock(stock)
            write_all_with_retry(stock2, minlvl, tx)
            st.cache_data.clear()
            st.success("تم التجميع والحفظ.")
    with c4:
        if st.button("إنشاء/تجديد الهيكل القياسي"):
            ensure_excel_file()
            st.success("تم التأكد من وجود الملف والأوراق القياسية (بدون MinLevels).")


def render_credits():
    year = datetime.now().year
    with st.sidebar:
        st.markdown(f"<div class='sidebar-credit'>© {year} — <b>{DEV_NAME}</b></div>", unsafe_allow_html=True)
    st.markdown(f"<div class='dev-credit'>© {year} — <b>{DEV_NAME}</b></div>", unsafe_allow_html=True)


def page_dashboard():
    st.subheader("لوحة التحكم")
    top_nav()
    file_status_badge()
    stock, minlvl, tx, _ = read_all()
    cfg = load_config()
    min_level = int(cfg.get("global_min_level", 2))
    total_items = stock["الكود"].nunique()
    total_qty = int(stock["المخزون"].sum()) if not stock.empty else 0
    loc_count = len(get_unique_locations(stock))
    suf = _suffix_to_use(cfg)
    # ✅ الآن: الأصلي = لا يحتوي على -S
    orig_count = (~stock["الكود"].astype(str).str.endswith(suf)).sum()
    comm_count = total_items - orig_count
    low_df, oos_df = compute_low_and_oos(stock, min_level)
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("عدد الأكواد", total_items)
    c2.metric("أكواد أصلية", int(orig_count))
    c3.metric("أكواد تجارية", int(comm_count))
    c4.metric("غير متوفر (0)", int(len(oos_df)))
    c5.metric(f"قريب من النفاد (≤ {min_level})", int(len(low_df)))
    if len(oos_df) > 0:
        st.error("قطع غير متوفرة حاليًا (المخزون = 0):")
        st.dataframe(oos_df, use_container_width=True, height=200)
    if len(low_df) > 0:
        st.warning(f"قطع اقتربت من النفاد (≤ {min_level}):")
        st.dataframe(low_df, use_container_width=True, height=200)
    st.markdown("### المخزون الحالي")
    st.dataframe(stock.sort_values(["الكود", "الموقع"]), use_container_width=True, height=420)


# -------------------------------------------------
# Main
# -------------------------------------------------
def main():
    st.title("نظام إدارة مخزون قطع السيارات (يعتمد ملف Excel واحد)")
    st.caption("قراءة وكتابة مباشرة داخل: " + EXCEL_PATH)
    if "menu" not in st.session_state:
        st.session_state.menu = "لوحة التحكم"
    default_index = PAGES.index(st.session_state.menu) if st.session_state.menu in PAGES else 0
    menu = st.sidebar.radio("القائمة", PAGES, index=default_index, key="sidebar_menu")
    if menu != st.session_state.menu:
        st.session_state.menu = menu
    if menu == "لوحة التحكم":
        page_dashboard()
    elif menu == "بحث/مسح":
        page_find_and_scan()
    elif menu == "العمليات":
        page_operations()
    elif menu == "الجرد":
        page_stocktake()
    elif menu == "إضافة قطعة جديدة":
        page_add_new_item()
    elif menu == "دمج ملف جديد":
        page_merge()
    elif menu == "تحرير البيانات":
        page_data_editor()
    elif menu == "استيراد/تصدير":
        page_import_export()
    elif menu == "إعدادات":
        page_settings()
    render_credits()


if __name__ == "__main__":
    main()