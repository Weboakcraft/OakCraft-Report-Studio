"""
Analytics: turns the standardised tables into every number the dashboard
and the Excel report show, plus auto-generated findings and data checks.
"""
from __future__ import annotations

import math

import pandas as pd

from .loader import KIND_LABELS

SRC_ORDER = ["Meta", "IndiaMART", "JustDial"]
LEAD_STATUSES = ["Qualified", "Not Qualified", "Call Not Pick", "Pending (Not Worked)"]
PENDING = "Pending (Not Worked)"
AGE_BUCKETS = [("0–2 days", 0, 2), ("3–4 days", 3, 4), ("5–7 days", 5, 7), ("8–14 days", 8, 14), ("15+ days", 15, 10**6)]


def pct(a, b):
    return float(a) / float(b) if b else 0.0


def ordered(values, order=SRC_ORDER):
    vals = set(values)
    return [v for v in order if v in vals] + sorted(vals - set(order))


def _d(x):
    return None if x is None or pd.isna(x) else pd.Timestamp(x).strftime("%Y-%m-%d")


def _n(x):
    if x is None:
        return None
    try:
        if pd.isna(x):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(x, float) and x.is_integer():
        return int(x)
    return x.item() if hasattr(x, "item") else x


def clean(o):
    """Make any nested structure JSON-safe."""
    if isinstance(o, dict):
        return {str(k): clean(v) for k, v in o.items()}
    if isinstance(o, (list, tuple)):
        return [clean(v) for v in o]
    if isinstance(o, float):
        return None if math.isnan(o) or math.isinf(o) else round(o, 6)
    if hasattr(o, "item"):
        return clean(o.item())
    if isinstance(o, pd.Timestamp):
        return _d(o)
    return o


# ----------------------------------------------------------------- context
class Ctx:
    def __init__(self, leads, pipe, found, issues, asof):
        self.asof = pd.Timestamp(asof).normalize()
        self.leads = leads.sort_values(["Source", "Lead Date"]).reset_index(drop=True)
        pipe = pipe.sort_values(["Salesperson", "Qualified On"]).reset_index(drop=True)
        pipe["Days"] = (self.asof - pipe["Qualified On"]).dt.days
        self.pipe = pipe
        self.found, self.issues = found, issues
        self.lsrc = ordered(self.leads.Source)
        self.psrc = ordered(self.pipe.Source)
        self.people = sorted((set(self.leads.Salesperson) | set(self.pipe.Salesperson)) - {"Unassigned"})
        self.lpeople = self.people + (["Unassigned"] if (self.leads.Salesperson == "Unassigned").any() else [])
        self.lstat = LEAD_STATUSES + (["Other"] if (self.leads.Status == "Other").any() else [])
        mids = sorted(s for s in set(self.pipe.Stage) if s not in ("No Action", "Won", "Lost"))
        self.mids = mids
        self.pstages = ["No Action"] + mids + ["Won", "Lost"]


def lead_block(d: pd.DataFrame) -> dict:
    n = len(d)
    st = d.Status.value_counts()
    q, pend = int(st.get("Qualified", 0)), int(st.get(PENDING, 0))
    worked = n - pend
    return {"leads": n, "qualified": q, "not_qualified": int(st.get("Not Qualified", 0)),
            "call_not_pick": int(st.get("Call Not Pick", 0)), "pending": pend, "other": int(st.get("Other", 0)),
            "worked": worked, "worked_pct": pct(worked, n), "conv": pct(q, n), "conv_worked": pct(q, worked)}


def pipe_block(d: pd.DataFrame, stages) -> dict:
    n = len(d)
    st = d.Stage.value_counts()
    na = int(st.get("No Action", 0))
    na_d = d[d.Stage == "No Action"]
    return {"total": n, "stages": {s: int(st.get(s, 0)) for s in stages}, "no_action": na,
            "actioned": n - na, "action_rate": pct(n - na, n), "won": int(st.get("Won", 0)),
            "lost": int(st.get("Lost", 0)), "win_rate": pct(st.get("Won", 0), n),
            "oldest": _n(na_d.Days.max()) if len(na_d) else None, "na5": int((na_d.Days >= 5).sum())}


# ----------------------------------------------------------------- sections
def kpis(c: Ctx) -> dict:
    lb = lead_block(c.leads)
    pb = pipe_block(c.pipe, c.pstages)
    return {**lb, "unassigned": int((c.leads.Salesperson == "Unassigned").sum()),
            "pipe_total": pb["total"], "no_action": pb["no_action"], "actioned": pb["actioned"],
            "action_rate": pb["action_rate"], "won": pb["won"], "lost": pb["lost"],
            "win_rate": pb["win_rate"], "na5": pb["na5"]}


def funnel(k: dict) -> list:
    rows = [("Leads received", k["leads"], "lead"), ("Leads worked", k["worked"], "lead"),
            ("Leads qualified", k["qualified"], "lead"), ("Qualified in pipeline", k["pipe_total"], "pipe"),
            ("Actioned after qualification", k["actioned"], "pipe"), ("Won", k["won"], "pipe")]
    out = []
    for i, (name, v, grp) in enumerate(rows):
        base = rows[0][1] if grp == "lead" else rows[3][1]
        prev = None if i in (0, 3) else rows[i - 1][1]
        out.append({"stage": name, "value": v, "group": grp, "of_top": pct(v, base),
                    "of_prev": None if prev is None else pct(v, prev),
                    "leak": None if prev is None else prev - v})
    return out


def scorecard(c: Ctx) -> dict:
    rows = []
    for p in c.lpeople:
        lb = lead_block(c.leads[c.leads.Salesperson == p])
        pb = pipe_block(c.pipe[c.pipe.Salesperson == p], c.pstages)
        rows.append({"name": p, **lb, "pipe": pb["total"], "no_action": pb["no_action"], "actioned": pb["actioned"],
                     "action_rate": pb["action_rate"], "won": pb["won"], "lost": pb["lost"],
                     "win_rate": pb["win_rate"], "na5": pb["na5"], "oldest": pb["oldest"]})
    lb, pb = lead_block(c.leads), pipe_block(c.pipe, c.pstages)
    team = {"name": "Team Total", **lb, "pipe": pb["total"], "no_action": pb["no_action"], "actioned": pb["actioned"],
            "action_rate": pb["action_rate"], "won": pb["won"], "lost": pb["lost"], "win_rate": pb["win_rate"],
            "na5": pb["na5"], "oldest": pb["oldest"]}
    return {"rows": rows, "team": team}


def person_source(c: Ctx) -> list:
    out = []
    for p in c.lpeople:
        d = c.leads[c.leads.Salesperson == p]
        if not len(d):
            continue
        children = []
        for s in c.lsrc:
            ds = d[d.Source == s]
            if len(ds):
                children.append({"source": s, **lead_block(ds), "share": pct(len(ds), len(d))})
        out.append({"name": p, "total": lead_block(d), "sources": children})
    return out


def sources(c: Ctx) -> dict:
    rows = []
    total_q = int((c.leads.Status == "Qualified").sum())
    for s in ordered(list(c.lsrc) + list(c.psrc)):
        d = c.leads[c.leads.Source == s]
        pb = pipe_block(c.pipe[c.pipe.Source == s], c.pstages)
        lb = lead_block(d)
        rows.append({"source": s, **lb, "share": pct(len(d), len(c.leads)), "share_q": pct(lb["qualified"], total_q),
                     "unassigned": int((d.Salesperson == "Unassigned").sum()), "pipe": pb["total"],
                     "actioned": pb["actioned"], "action_rate": pb["action_rate"], "won": pb["won"],
                     "win_rate": pb["win_rate"]})
    meta = c.leads[c.leads.Source == "Meta"]
    plats = []
    for sub in sorted(set(meta["Sub-Source"]) - {"", "Meta"}):
        d = meta[meta["Sub-Source"] == sub]
        plats.append({"platform": sub, **lead_block(d), "share": pct(len(d), len(meta))})
    return {"rows": rows, "platforms": plats}


def pipeline(c: Ctx) -> dict:
    people = []
    for p in sorted(set(c.pipe.Salesperson)):
        d = c.pipe[c.pipe.Salesperson == p]
        kids = [{"source": s, **pipe_block(d[d.Source == s], c.pstages)} for s in c.psrc if (d.Source == s).any()]
        people.append({"name": p, "total": pipe_block(d, c.pstages), "sources": kids})
    na = c.pipe[c.pipe.Stage == "No Action"]
    age = []
    for p in sorted(set(c.pipe.Salesperson)):
        d = na[na.Salesperson == p]
        age.append({"name": p, "counts": [int(((d.Days >= lo) & (d.Days <= hi)).sum()) for _, lo, hi in AGE_BUCKETS],
                    "total": int(len(d))})
    ctypes = []
    for t, d in c.pipe.groupby("Customer Type"):
        qty = d["Purchase Qty"].fillna(0).sum()
        ctypes.append({"type": t, "total": len(d), "no_action": int((d.Stage == "No Action").sum()),
                       "actioned": int((d.Stage != "No Action").sum()),
                       "action_rate": pct((d.Stage != "No Action").sum(), len(d)),
                       "qty": _n(qty), "avg_qty": pct(qty, len(d))})
    ctypes.sort(key=lambda x: -x["total"])
    return {"stages": c.pstages, "mid_stages": c.mids, "people": people, "team": pipe_block(c.pipe, c.pstages),
            "age_buckets": [b[0] for b in AGE_BUCKETS], "ageing": age, "customer_types": ctypes}


def trend(c: Ctx) -> dict:
    d = c.leads.dropna(subset=["Lead Date"]).copy()
    if not len(d):
        return {"dates": [], "series": [], "qualified": []}
    d["day"] = d["Lead Date"].dt.normalize()
    days = pd.date_range(d.day.min(), d.day.max(), freq="D")
    series = [{"source": s, "values": [int(((d.Source == s) & (d.day == x)).sum()) for x in days]} for s in c.lsrc]
    qual = [int(((d.Status == "Qualified") & (d.day == x)).sum()) for x in days]
    return {"dates": [x.strftime("%Y-%m-%d") for x in days], "series": series, "qualified": qual}


def next_step(qty, ctype) -> str:
    big = (qty is not None and not pd.isna(qty) and qty >= 20) or str(ctype) not in ("Individual", "Not Captured")
    return "Call today + send quotation" if big else "Call + share catalogue / price"


def priority(days) -> str:
    if days is None or pd.isna(days):
        return "Normal"
    return "Critical" if days >= 5 else "High" if days >= 3 else "Normal"


def lists(c: Ctx) -> dict:
    na = c.pipe[c.pipe.Stage == "No Action"].sort_values(["Days", "Purchase Qty"], ascending=[False, False])
    follow = [{"salesperson": r.Salesperson, "customer": r.Customer, "source": r.Source, "id": r._2,
               "contact": r.Contact, "customer_type": r._7, "qty": _n(r._8), "lead_date": _d(r._3),
               "qualified_on": _d(r._4), "days": _n(r.Days), "priority": priority(r.Days),
               "next_step": next_step(r._8, r._7)}
              for r in na.itertuples()]
    act = c.pipe[c.pipe.Stage != "No Action"].sort_values(["Stage", "Salesperson"])
    actioned = [{"salesperson": r.Salesperson, "customer": r.Customer, "source": r.Source, "id": r._2,
                 "stage": r.Stage, "customer_type": r._7, "qty": _n(r._8), "lead_date": _d(r._3),
                 "qualified_on": _d(r._4), "contact": r.Contact} for r in act.itertuples()]
    un = c.leads[c.leads.Salesperson == "Unassigned"].sort_values("Lead Date")
    unassigned = [{"source": r.Source, "id": r._3, "customer": r.Customer, "phone": r.Phone,
                   "lead_date": None if pd.isna(r._4) else pd.Timestamp(r._4).strftime("%Y-%m-%d %H:%M"),
                   "status": r._8} for r in un.itertuples()]
    return {"followup": follow, "actioned": actioned, "unassigned": unassigned}


# ----------------------------------------------------------------- findings
def insights(c: Ctx) -> list:
    leads, pipe, out = c.leads, c.pipe, []
    N = len(leads)
    if N:
        Q = int((leads.Status == "Qualified").sum()); P = int((leads.Status == PENDING).sum())
        out.append({"finding": f"{N} leads me se {Q} Qualified ({pct(Q, N):.1%}). {P} leads ({pct(P, N):.0%}) par abhi tak koi kaam nahi hua (Created / New).",
                    "action": "Roz subah pending leads ka review; 24 ghante me first call ka rule lagayen.",
                    "owner": "Sales Head", "severity": "high" if pct(P, N) > 0.3 else "medium", "area": "Leads"})
        rows = []
        for p in c.people:
            d = leads[leads.Salesperson == p]
            if len(d):
                b = lead_block(d)
                rows.append((p, b["leads"], b["conv"], b["conv_worked"], b["pending"]))
        if len(rows) > 1:
            best = max(rows, key=lambda x: x[2]); worst = min(rows, key=lambda x: x[2])
            msg = f"{best[0]} ka Lead→Qualified conversion sabse zyada ({best[2]:.1%}); {worst[0]} ka sabse kam ({worst[2]:.1%})."
            if worst[4] / max(worst[1], 1) > 0.4:
                msg += (f" {worst[0]} ki {worst[4]} of {worst[1]} leads pending hain – worked leads par conversion "
                        f"{worst[3]:.1%} hai, yaani problem calling speed ki hai, quality ki nahi.")
                act = f"{worst[0]} ke pending leads 48 ghante me clear karwayen; zarurat ho to kuch leads re-assign karein."
            else:
                act = f"{worst[0]} ke calls ka quality review karein; {best[0]} ki pitch / script team ke saath share karein."
            out.append({"finding": msg, "action": act, "owner": "Sales Head", "severity": "high", "area": "Salespeople"})
        sc = []
        for s in c.lsrc:
            b = lead_block(leads[leads.Source == s]); sc.append((s, b["leads"], b["conv"], b["conv_worked"]))
        if sc:
            vol = max(sc, key=lambda x: x[1]); qual = max(sc, key=lambda x: x[3])
            msg = (f"{vol[0]} se sabse zyada leads ({pct(vol[1], N):.0%} volume) aati hain, conversion {vol[2]:.1%}. "
                   f"Worked leads par best source {qual[0]} ({qual[3]:.0%}).")
            small = [s for s in sc if s[1] < 20]
            if small:
                msg += " " + ", ".join(f"{s[0]} ({s[1]} leads)" for s in small) + " ka sample chhota hai – trend ke liye 3–4 hafte dekhein."
            out.append({"finding": msg, "action": f"{qual[0]} leads ko sabse pehle call karein; {vol[0]} ke ad targeting / form questions tighten karein.",
                        "owner": "Marketing", "severity": "medium", "area": "Sources"})
        un = leads[leads.Salesperson == "Unassigned"]
        if len(un):
            src = un.Source.value_counts()
            out.append({"finding": f"{len(un)} leads kisi salesperson ko assign nahi hain ({', '.join(f'{k} {v}' for k, v in src.items())}).",
                        "action": "Aaj hi assign karein; CRM me auto-assignment (round-robin) on karein.",
                        "owner": "MIS Team", "severity": "high", "area": "Leads"})
        cnp = leads[leads.Status == "Call Not Pick"]
        if len(cnp):
            out.append({"finding": f"{len(cnp)} leads Call Not Pick par hain ({', '.join(f'{k.split()[0]} {v}' for k, v in cnp.Salesperson.value_counts().items())}).",
                        "action": "3 attempts (alag time slots) + WhatsApp message ke baad hi lead close karein.",
                        "owner": "Sales Team", "severity": "medium", "area": "Leads"})
    if len(pipe):
        T = len(pipe); NA = int((pipe.Stage == "No Action").sum()); A = T - NA
        out.append({"finding": f"Pipeline me {T} qualified leads; sirf {A} ({pct(A, T):.1%}) par qualification ke baad kaam hua. {NA} leads abhi bhi 'Qualified' par ruki hain.",
                    "action": "Har qualified lead par 48 ghante me quotation / visit / follow-up update karna mandatory karein.",
                    "owner": "Sales Head", "severity": "high" if pct(A, T) < 0.3 else "medium", "area": "Pipeline"})
        zero = [p for p in sorted(set(pipe.Salesperson)) if ((pipe.Salesperson == p) & (pipe.Stage != "No Action")).sum() == 0]
        if zero:
            n0 = int(pipe.Salesperson.isin(zero).sum())
            out.append({"finding": f"{', '.join(zero)} ki {n0} qualified leads me se ek par bhi aage ka kaam record nahi hua.",
                        "action": "Daily pipeline review meeting me har lead ka next step aur date likhwayen.",
                        "owner": "Sales Head", "severity": "high", "area": "Pipeline"})
        old = pipe[(pipe.Stage == "No Action") & (pipe.Days >= 5)]
        if len(old):
            out.append({"finding": f"{len(old)} qualified leads 5+ din se bina kaam ke padi hain ({', '.join(f'{k.split()[0]} {v}' for k, v in old.Salesperson.value_counts().items())}).",
                        "action": "Follow-up list ki Critical leads aaj hi call karein – lead jitni purani, close hone ka chance utna kam.",
                        "owner": "Sales Team", "severity": "high", "area": "Pipeline"})
        b2b = pipe[~pipe["Customer Type"].isin(["Individual", "Not Captured"])]
        if len(b2b):
            bna = b2b[b2b.Stage == "No Action"]
            tail = "ek par bhi action nahi hua" if len(bna) == len(b2b) else f"{len(bna)} par abhi tak action nahi hua"
            out.append({"finding": f"B2B / trade qualified leads ({', '.join(b2b['Customer Type'].value_counts().index[:4])}…): {len(b2b)} leads, {b2b['Purchase Qty'].fillna(0).sum():.0f} pcs demand – {tail}.",
                        "action": "B2B leads senior salesperson ko dein; 24 ghante me quotation bhejein.",
                        "owner": "Sales Head", "severity": "high" if len(bna) else "low", "area": "Pipeline"})
        if (pipe.Stage == "Lost").sum() == 0:
            out.append({"finding": "Pipeline me ek bhi lead Lost mark nahi hai – dead leads close nahi ho rahi, pipeline asli se badi dikh rahi hai.",
                        "action": "Jo lead 3 follow-up ke baad bhi respond na kare, use Lost + reason ke saath close karein.",
                        "owner": "Sales Team", "severity": "medium", "area": "Pipeline"})
        if pipe["Value"].isna().all():
            out.append({"finding": "Value column poora khaali hai – ₹ pipeline aur won value nahi nikal sakti.",
                        "action": "Quotation bhejte waqt Value field bharna mandatory karein.",
                        "owner": "MIS Team", "severity": "low", "area": "Data"})
    return out


def quality(c: Ctx) -> list:
    leads, pipe, dq = c.leads, c.pipe, []
    for k in ("meta", "indiamart", "justdial", "pipeline"):
        if k in c.found:
            files = ", ".join(f"{f['name']} ({f['rows']} rows)" for f in c.found[k])
            dq.append({"level": "ok", "text": f"{KIND_LABELS[k]} file detected: {files}"})
        else:
            dq.append({"level": "issue", "text": f"{KIND_LABELS[k]} file nahi mili – uske sections khaali / zero dikhenge."})
    dq += [{"level": "issue", "text": i} for i in c.issues]
    dq.append({"level": "note", "text": "Lead status mapping: Qualified / Quotation / Won → Qualified; Not Qualified / Lost / Junk → Not Qualified; "
                                        "Call Not Pick / CNP / Busy → Call Not Pick; Created / New → Pending (Not Worked)."})
    dq.append({"level": "note", "text": "Pipeline: 'Qualified' = qualify hone ke baad koi kaam nahi (No Action); Won = crack; Lost = kaam ki nahi; "
                                        "baaki statuses (e.g. Quotation Sent) = in-progress kaam. 'Enquiry' source = JustDial."})
    oth = leads[leads.Status == "Other"]
    if len(oth):
        dq.append({"level": "issue", "text": f"{len(oth)} lead statuses map nahi hue (Other): {', '.join(oth['Original Status'].unique()[:8])}."})
    ph = leads[leads.Phone != ""]
    dup = ph[ph.duplicated(["Source", "Phone"], keep=False)]
    if len(dup):
        names = ", ".join(f"{a} ({b})" for a, b in dup.groupby("Phone").agg({"Customer": "first", "Source": "first"}).itertuples(index=False))
        dq.append({"level": "issue", "text": f"{dup.Phone.nunique()} phone numbers ek hi source me 2+ baar aaye ({len(dup)} rows): {names[:300]}. Alag enquiry gina hai."})
    multi = ph[ph.duplicated("Phone", keep=False)].groupby("Phone").Source.nunique()
    if (multi > 1).any():
        dq.append({"level": "note", "text": f"{int((multi > 1).sum())} customers ne ek se zyada source se enquiry ki (cross-source duplicates)."})
    jd = leads[(leads.Source == "JustDial") & leads.Remark.astype(str).str.contains("pick", case=False) & (leads.Status == "Qualified")]
    if len(jd):
        dq.append({"level": "issue", "text": f"JustDial: {len(jd)} lead ka Stage 'Call Not Pick' hai par Qualification Status 'Qualified' ({', '.join(jd['Lead ID'])}). Col P ke hisaab se Qualified gina – verify karein."})
    groups = {}
    for p in set(leads.Salesperson) | set(pipe.Salesperson):
        groups.setdefault(p.replace(" ", "").lower()[:6], set()).add(p)
    near = [v for v in groups.values() if len(v) > 1]
    if near:
        dq.append({"level": "issue", "text": "Milte-julte salesperson naam: " + "; ".join(" / ".join(sorted(v)) for v in near) + " – spelling ek karein."})
    dq.append({"level": "note", "text": "Salesperson naam Title Case me standardise kiye (e.g. 'UJALA RAJPUT' = 'Ujala Rajput'). Blank = 'Unassigned'."})
    if len(pipe):
        nc = int((pipe["Customer Type"] == "Not Captured").sum())
        if nc:
            dq.append({"level": "issue", "text": f"Pipeline: {nc} leads me Customer Type khaali."})
        if pipe["Value"].isna().all():
            dq.append({"level": "issue", "text": "Pipeline: Value column poora khaali – ₹ value analysis possible nahi."})
        nq = int(pipe["Purchase Qty"].isna().sum() + (pipe["Purchase Qty"] == 0).sum())
        if nq:
            dq.append({"level": "note", "text": f"Pipeline: {nq} leads me Purchase Qty khaali ya 0."})
    nd = int(leads["Lead Date"].isna().sum())
    if nd:
        dq.append({"level": "issue", "text": f"{nd} leads ki date read nahi hui."})
    return dq


# ----------------------------------------------------------------- entry
def summarize(c: Ctx) -> dict:
    k = kpis(c)
    lead_dates = c.leads["Lead Date"].dropna()
    q_dates = c.pipe["Qualified On"].dropna()
    files = [{"name": f["name"], "rows": f["rows"], "kind": kind, "label": KIND_LABELS[kind]}
             for kind, fs in c.found.items() for f in fs]
    return clean({
        "meta": {"asof": c.asof.strftime("%Y-%m-%d"),
                 "lead_from": _d(lead_dates.min()) if len(lead_dates) else None,
                 "lead_to": _d(lead_dates.max()) if len(lead_dates) else None,
                 "qual_from": _d(q_dates.min()) if len(q_dates) else None,
                 "qual_to": _d(q_dates.max()) if len(q_dates) else None,
                 "sources": c.lsrc, "people": c.people, "files": files,
                 "has_leads": bool(len(c.leads)), "has_pipeline": bool(len(c.pipe))},
        "kpi": k, "funnel": funnel(k), "scorecard": scorecard(c), "person_source": person_source(c),
        "sources": sources(c), "pipeline": pipeline(c), "trend": trend(c), **lists(c),
        "insights": insights(c), "quality": quality(c)})
