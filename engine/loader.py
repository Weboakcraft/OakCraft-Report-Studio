"""
Loader: read uploaded exports, detect which report each file is, and
standardise everything into two tidy tables.

    leads  – one row per enquiry from Meta / IndiaMART / JustDial
    pipe   – one row per qualified lead from the CRM "qualified" export
"""
from __future__ import annotations

import io
import os
import re

import pandas as pd

LEAD_COLS = ["Source", "Sub-Source", "Lead ID", "Lead Date", "Customer", "Phone",
             "Salesperson", "Original Status", "Remark"]
PIPE_COLS = ["Source", "Record Id", "Lead Date", "Qualified On", "Customer", "Contact",
             "Customer Type", "Purchase Qty", "Value", "Salesperson", "Original Status"]

KIND_LABELS = {"meta": "Meta Leads", "indiamart": "IndiaMART", "justdial": "JustDial",
               "pipeline": "Qualified Pipeline"}


# ----------------------------------------------------------------- helpers
def norm(s) -> str:
    return re.sub(r"\s+", " ", str(s)).strip().lower()


def col(df: pd.DataFrame, names, letter=None) -> pd.Series:
    """Column by header name (case-insensitive); falls back to an Excel column letter."""
    lut = {norm(c): c for c in df.columns}
    for n in names:
        if norm(n) in lut:
            return df[lut[norm(n)]]
    if letter is not None:
        idx = ord(letter.upper()) - 65
        if idx < df.shape[1]:
            return df.iloc[:, idx]
    return pd.Series([None] * len(df), index=df.index, dtype=object)


def has(df: pd.DataFrame, *names) -> bool:
    cs = {norm(c) for c in df.columns}
    return any(norm(n) in cs for n in names)


def person(x) -> str:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return "Unassigned"
    s = str(x).strip()
    if not s or s.lower() in ("nan", "none", "-", "na", "n/a"):
        return "Unassigned"
    return " ".join(w.capitalize() for w in s.split())


def to_dt(s: pd.Series) -> pd.Series:
    return pd.to_datetime(s, dayfirst=True, errors="coerce", format="mixed")


def phone(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    digits = re.sub(r"\D", "", str(v).split(".")[0])
    return digits[-10:] if len(digits) >= 10 else digits


def lead_bucket(s) -> str:
    """Map any lead-file status onto four standard buckets."""
    k = norm(s)
    if k in ("", "nan", "none"):
        return "Pending (Not Worked)"
    if "not" in k and "qual" in k:
        return "Not Qualified"
    if any(t in k for t in ("unqual", "junk", "fake", "spam", "not interest", "lost")):
        return "Not Qualified"
    if any(t in k for t in ("qualif", "quot", "won", "negoti", "follow", "hot")):
        return "Qualified"
    if any(t in k for t in ("pick", "call", "ring", "busy", "switch", "unreach", "dnp")):
        return "Call Not Pick"
    if k in ("new", "created", "open", "pending", "fresh", "assigned") or "new" in k:
        return "Pending (Not Worked)"
    return "Other"


def pipe_stage(s) -> str:
    """Pipeline status: Qualified = no work after qualification."""
    k = norm(s)
    if k in ("", "nan", "none", "qualified"):
        return "No Action"
    if "won" in k or "converted" in k:
        return "Won"
    if "lost" in k or "dead" in k:
        return "Lost"
    return " ".join(w.capitalize() for w in str(s).split())


def src_name(s) -> str:
    k = norm(s)
    if k in ("meta", "facebook", "instagram", "fb", "ig") or "meta" in k:
        return "Meta"
    if "india" in k or k == "im":
        return "IndiaMART"
    if "just" in k or k in ("enquiry", "enquiries", "jd"):
        return "JustDial"
    return str(s).strip() or "Unknown"


# ----------------------------------------------------------------- detection
def detect(df: pd.DataFrame, fname: str):
    f = fname.lower()
    if has(df, "qualified on", "record id") or "qualified" in f:
        return "pipeline"
    if has(df, "lead_status", "assigned_to", "platform") or "meta" in f:
        return "meta"
    if has(df, "enquiry no", "query time", "assigned to") or "indiamart" in f:
        return "indiamart"
    if has(df, "qualification status", "salesperson") or "enquir" in f or "justdial" in f:
        return "justdial"
    return None


def read_table(name: str, data: bytes) -> pd.DataFrame:
    buf = io.BytesIO(data)
    if name.lower().endswith((".xlsx", ".xlsm", ".xls")):
        df = pd.read_excel(buf)
    else:
        df = pd.read_csv(buf)
    return df.dropna(how="all")


# ----------------------------------------------------------------- standardise
def _meta(df):
    sub = col(df, ["platform"]).fillna("").astype(str).str.upper().replace({"FB": "Facebook", "IG": "Instagram"})
    return pd.DataFrame({
        "Source": "Meta", "Sub-Source": sub.replace("", "Meta"),
        "Lead ID": col(df, ["enquiry_id", "lead_id", "id"], "A").astype(str),
        "Lead Date": to_dt(col(df, ["created_time_ist", "created_time", "date"], "B")),
        "Customer": col(df, ["full_name", "name"], "D"),
        "Phone": col(df, ["phone_number", "phone"], "E").map(phone),
        "Salesperson": col(df, ["assigned_to"], "H").map(person),
        "Original Status": col(df, ["lead_status", "status"], "G").fillna("").astype(str),
        "Remark": ""})


def _indiamart(df):
    return pd.DataFrame({
        "Source": "IndiaMART", "Sub-Source": "IndiaMART",
        "Lead ID": col(df, ["enquiry no", "enquiry id"], "G").astype(str),
        "Lead Date": to_dt(col(df, ["query time", "date"], "A")),
        "Customer": col(df, ["buyer", "name"], "B"),
        "Phone": col(df, ["mobile", "phone"], "D").map(phone),
        "Salesperson": col(df, ["assigned to"], "L").map(person),
        "Original Status": col(df, ["status"], "K").fillna("").astype(str),
        "Remark": ""})


def _justdial(df):
    return pd.DataFrame({
        "Source": "JustDial", "Sub-Source": "JustDial",
        "Lead ID": col(df, ["enquiry id"], "A").astype(str),
        "Lead Date": to_dt(col(df, ["date", "enquiry time"], "B")),
        "Customer": col(df, ["customer", "name"], "C"),
        "Phone": col(df, ["mobile", "phone", "contact"]).map(phone),
        "Salesperson": col(df, ["salesperson"], "G").map(person),
        "Original Status": col(df, ["qualification status"], "P").fillna("").astype(str),
        "Remark": col(df, ["stage"]).fillna("").astype(str)})


def _pipeline(df):
    return pd.DataFrame({
        "Source": col(df, ["source"], "A").map(src_name),
        "Record Id": col(df, ["record id", "id"], "B").astype(str),
        "Lead Date": to_dt(col(df, ["date", "lead date"], "D")),
        "Qualified On": to_dt(col(df, ["qualified on"], "C")),
        "Customer": col(df, ["name", "customer"], "E"),
        "Contact": col(df, ["contact", "phone", "mobile"], "F").map(phone),
        "Customer Type": col(df, ["customer type"], "H").fillna("Not Captured").replace("", "Not Captured"),
        "Purchase Qty": pd.to_numeric(col(df, ["purchase qty", "qty", "quantity"], "I"), errors="coerce"),
        "Value": pd.to_numeric(col(df, ["value"], "J"), errors="coerce"),
        "Salesperson": col(df, ["salesperson"], "K").map(person),
        "Original Status": col(df, ["status"], "L").fillna("Qualified").astype(str)})


BUILDERS = {"meta": _meta, "indiamart": _indiamart, "justdial": _justdial, "pipeline": _pipeline}


def inspect_file(name: str, data: bytes) -> dict:
    """Light-weight peek used by the UI right after upload."""
    try:
        df = read_table(name, data)
    except Exception as exc:  # noqa: BLE001
        return {"name": name, "kind": None, "label": "Unreadable", "rows": 0, "error": str(exc)[:160]}
    kind = detect(df, name)
    return {"name": name, "kind": kind, "label": KIND_LABELS.get(kind, "Not recognised"),
            "rows": int(len(df)), "columns": [str(c) for c in df.columns[:20]]}


def load(files):
    """files: iterable of (name, bytes). Returns leads, pipe, found, issues."""
    leads, pipe, found, issues = [], [], {}, []
    for name, data in files:
        base = os.path.basename(name)
        try:
            df = read_table(base, data)
        except Exception as exc:  # noqa: BLE001
            issues.append(f"{base}: could not read ({str(exc)[:120]})")
            continue
        kind = detect(df, base)
        if kind is None:
            issues.append(f"{base}: file type not recognised – skipped")
            continue
        found.setdefault(kind, []).append({"name": base, "rows": int(len(df))})
        (pipe if kind == "pipeline" else leads).append(BUILDERS[kind](df))

    L_ = pd.concat(leads, ignore_index=True) if leads else pd.DataFrame(columns=LEAD_COLS)
    P_ = pd.concat(pipe, ignore_index=True) if pipe else pd.DataFrame(columns=PIPE_COLS)
    L_["Lead Date"] = pd.to_datetime(L_["Lead Date"], errors="coerce")
    for c in ("Lead Date", "Qualified On"):
        P_[c] = pd.to_datetime(P_[c], errors="coerce")
    for c in ("Purchase Qty", "Value"):
        P_[c] = pd.to_numeric(P_[c], errors="coerce")
    L_["Status"] = L_["Original Status"].map(lead_bucket)
    P_["Stage"] = P_["Original Status"].map(pipe_stage)
    return L_, P_, found, issues
