"""
Excel builder: one-sheet "Lead Report" dashboard + two data tabs.
Every number on the report is a live COUNTIFS formula over the data tabs,
so a status corrected in Data_Leads / Data_Pipeline updates the report.
"""
from __future__ import annotations

import io
import re

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.comments import Comment
from openpyxl.formatting.rule import DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.properties import CalcProperties

from .analytics import AGE_BUCKETS, Ctx, next_step

FONT = "Arial"; NAVY = "1F3864"; TEAL = "2E75B6"; GREEN = "548235"; LIGHT = "DDEBF7"; PALE = "F3F7FC"
GOOD = "C6EFCE"; BAD = "FFC7CE"; WARN = "FFEB9C"
thin = Side(style="thin", color="BFBFBF"); BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
CEN = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
PCT = '0.0%;-0.0%;"-"'; INT = '#,##0;-#,##0;"-"'; DAYS = '0;-0;"-"'
LEVEL = {"ok": ("✔ OK", GOOD), "issue": ("⚠ Issue", WARN), "note": ("ℹ Note", LIGHT)}


def fnt(b=False, s=10, c="000000", i=False):
    return Font(name=FONT, bold=b, size=s, color=c, italic=i)


def fill(c):
    return PatternFill("solid", start_color=c, end_color=c)


def _v(v):
    if v is None or v is pd.NaT:
        return None
    if isinstance(v, float) and pd.isna(v):
        return None
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()
    return v


class Sheet:
    def __init__(self, ws):
        self.ws = ws
        add = ws.conditional_formatting.add

        def safe(rng, rule):  # skip empty ranges (e.g. no salespeople)
            try:
                a, b = rng.split(":")
                if int(re.sub(r"\D", "", b)) < int(re.sub(r"\D", "", a)):
                    return
            except ValueError:
                return
            add(rng, rule)
        self.cf = safe

    def c(self, r, c, v, fmt=None, b=False, bg=None, al=CEN, color="000000", size=10, border=True, i=False):
        x = self.ws.cell(r, c, _v(v)); x.font = fnt(b, size, color, i); x.alignment = al
        if border: x.border = BORDER
        if fmt: x.number_format = fmt
        if bg: x.fill = fill(bg)
        return x

    def hdr(self, r, c0, labels, color=NAVY, h=40):
        for k, t in enumerate(labels):
            self.c(r, c0 + k, t, b=True, bg=color, color="FFFFFF")
        self.ws.row_dimensions[r].height = h

    def section(self, r, text, sub=None):
        self.ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=16)
        x = self.ws.cell(r, 1, text); x.font = fnt(True, 13, "FFFFFF"); x.fill = fill(NAVY)
        x.alignment = Alignment(vertical="center", indent=1)
        self.ws.row_dimensions[r].height = 24
        if sub:
            self.ws.cell(r + 1, 1, sub).font = fnt(False, 9, "595959", True)
            return r + 2
        return r + 1

    def merge(self, r, c1, c2):
        self.ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)

    def total(self, r, label, cols_sum, pct_cells, first, last, bg=NAVY, color="FFFFFF", label2=None):
        self.c(r, 1, label, b=True, bg=bg, color=color, al=LEFT)
        if label2: self.c(r, 2, label2, b=True, bg=bg, color=color, al=LEFT)
        for cc in cols_sum:
            self.c(r, cc, f"=SUM({L(cc)}{first}:{L(cc)}{last})", INT, True, bg, color=color)
        for cc, fml in pct_cells:
            self.c(r, cc, fml, PCT, True, bg, color=color)


def _data_tabs(wb, c: Ctx):
    DL = wb.create_sheet("Data_Leads"); DP = wb.create_sheet("Data_Pipeline")
    lcols = ["Source", "Sub-Source", "Lead ID", "Lead Date", "Customer", "Phone", "Salesperson", "Original Status", "Status", "Remark"]
    S = Sheet(DL); S.hdr(1, 1, lcols, h=30)
    for i, row in enumerate(c.leads[lcols].itertuples(index=False), 2):
        for j, v in enumerate(row, 1):
            x = DL.cell(i, j, _v(v)); x.font = fnt(); x.border = BORDER
            if j == 4: x.number_format = "dd-mmm-yyyy hh:mm"
    for j, w in enumerate([11, 11, 14, 17, 26, 13, 16, 17, 20, 18], 1):
        DL.column_dimensions[L(j)].width = w
    nL = len(c.leads) + 1; DL.freeze_panes = "A2"; DL.auto_filter.ref = f"A1:J{max(nL, 2)}"

    pcols = ["Source", "Record Id", "Lead Date", "Qualified On", "Customer", "Contact", "Customer Type",
             "Purchase Qty", "Value", "Salesperson", "Original Status", "Stage"]
    S = Sheet(DP); S.hdr(1, 1, pcols + ["Days Since Qualified"], h=30)
    for i, row in enumerate(c.pipe[pcols].itertuples(index=False), 2):
        for j, v in enumerate(row, 1):
            x = DP.cell(i, j, _v(v)); x.font = fnt(); x.border = BORDER
            if j in (3, 4): x.number_format = "dd-mmm-yyyy"
        x = DP.cell(i, 13, f"=IF(D{i}=\"\",\"\",'Lead Report'!$O$3-D{i})"); x.font = fnt(); x.border = BORDER; x.number_format = "0"
    for j, w in enumerate([11, 13, 12, 12, 24, 13, 22, 10, 10, 15, 15, 15, 11], 1):
        DP.column_dimensions[L(j)].width = w
    nP = len(c.pipe) + 1; DP.freeze_panes = "A2"; DP.auto_filter.ref = f"A1:M{max(nP, 2)}"
    for w in (DL, DP):
        w.sheet_properties.tabColor = "A6A6A6"
    return nL, nP


def build_workbook(c: Ctx, insights: list, quality: list) -> bytes:
    wb = Workbook()
    ws = wb.active; ws.title = "Lead Report"; R = Sheet(ws)
    nL, nP = _data_tabs(wb, c)

    rng = lambda sh, col, n: f"{sh}!${col}$2:${col}${max(n, 2)}"
    LS, LSP, LST, LSUB = (rng("Data_Leads", x, nL) for x in "AGIB")
    PS, PSP, PSTG, PDAY, PCTY, PQ = (rng("Data_Pipeline", x, nP) for x in "AJLMGH")
    q = lambda t: '"' + t + '"'

    def lc(p=None, s=None, st=None, sub=None):
        a = []
        for crit, rg in ((p, LSP), (s, LS), (sub, LSUB), (st, LST)):
            if crit: a += [rg, crit]
        return "COUNTIFS(" + ",".join(a) + ")" if a else f"COUNTA({LS})"

    def pc(p=None, s=None, st=None, extra=()):
        a = []
        for crit, rg in ((p, PSP), (s, PS), (st, PSTG)):
            if crit: a += [rg, crit]
        a += list(extra)
        return "COUNTIFS(" + ",".join(a) + ")" if a else f"COUNTA({PS})"

    PEND = q("Pending (Not Worked)"); QUAL = q("Qualified"); NA = q("No Action")
    for j, w in enumerate([24] + [11.5] * 15 + [3], 1):
        ws.column_dimensions[L(j)].width = w
    ws.sheet_view.showGridLines = False

    # ---- title
    R.merge(1, 1, 13); ws["A1"] = "OakCraft – Lead & Sales Pipeline Master Report"; ws["A1"].font = fnt(True, 18, NAVY)
    R.merge(2, 1, 13)
    ld, qd = c.leads["Lead Date"].dropna(), c.pipe["Qualified On"].dropna()
    sub = (f"Lead period: {ld.min():%d %b %Y} – {ld.max():%d %b %Y}" if len(ld) else "Lead period: n/a")
    if len(qd): sub += f"   |   Qualified on: {qd.min():%d %b} – {qd.max():%d %b %Y}"
    ws["A2"] = sub + f"   |   Sources: {', '.join(c.lsrc) or '-'}"; ws["A2"].font = fnt(False, 10, "595959", True)
    R.c(3, 14, "Report as-of", b=True, al=Alignment(horizontal="right", vertical="center"), border=False)
    x = R.c(3, 15, c.asof, "dd-mmm-yyyy", b=True, bg="FFFF00", color="0000FF")
    x.comment = Comment("Input: report date. Ageing (days since qualified) is measured from this date.", "Report Studio")
    ws.merge_cells("O3:P3")

    # ---- A. KPIs
    r = R.section(5, "A.  HEADLINE KPIs")
    tiles1 = [("Total Leads", f"={lc()}", INT), ("Worked Leads %", f"=IFERROR(1-{lc(st=PEND)}/{lc()},0)", PCT),
              ("Qualified Leads", f"={lc(st=QUAL)}", INT), ("Lead → Qualified %", f"=IFERROR({lc(st=QUAL)}/{lc()},0)", PCT),
              ("Pending (Not Worked)", f"={lc(st=PEND)}", INT), ("Unassigned Leads", f"={lc(p=q('Unassigned'))}", INT),
              ("Call Not Pick", f"={lc(st=q('Call Not Pick'))}", INT), ("Not Qualified", f"={lc(st=q('Not Qualified'))}", INT)]
    tiles2 = [("Pipeline: Qualified", f"={pc()}", INT), ("No Action after Qual.", f"={pc(st=NA)}", INT),
              ("Actioned after Qual.", f"={pc()}-{pc(st=NA)}", INT), ("Action Rate %", f"=IFERROR(({pc()}-{pc(st=NA)})/{pc()},0)", PCT),
              ("Won", f"={pc(st=q('Won'))}", INT), ("Lost", f"={pc(st=q('Lost'))}", INT),
              ("Win Rate % (Won ÷ Qual.)", f"=IFERROR({pc(st=q('Won'))}/{pc()},0)", PCT),
              ("No Action 5+ days", f"={pc(st=NA, extra=[PDAY, q('>=5')])}", INT)]
    for tiles, color in ((tiles1, TEAL), (tiles2, GREEN)):
        for k, (t, v, fm) in enumerate(tiles):
            c0 = 1 + 2 * k
            R.merge(r, c0, c0 + 1); R.merge(r + 1, c0, c0 + 1)
            for cc in (c0, c0 + 1):
                for rr, bg in ((r, color), (r + 1, PALE)):
                    ws.cell(rr, cc).fill = fill(bg); ws.cell(rr, cc).border = BORDER
            a = ws.cell(r, c0, t); a.font = fnt(True, 9, "FFFFFF"); a.alignment = CEN
            b = ws.cell(r + 1, c0, v); b.font = fnt(True, 18, NAVY); b.alignment = CEN; b.number_format = fm
        ws.row_dimensions[r].height = 26; ws.row_dimensions[r + 1].height = 34
        r += 2
    ws.cell(r, 1, "Blue tiles = lead handling (Meta / IndiaMART / JustDial).  Green tiles = post-qualification pipeline (Qualified file).").font = fnt(False, 9, "808080", True)
    r += 2

    # ---- B. Funnel
    r = R.section(r, "B.  END-TO-END CONVERSION FUNNEL", "Stage 1–3 from lead files, stage 4–6 from the qualified pipeline file. % of previous stage shows where leads leak.")
    R.hdr(r, 1, ["Funnel Stage", "Leads", "% of Previous Stage", "% of Top of Funnel", "Leakage at this Stage"]); fr0 = r + 1
    stages = [("1. Leads Received", f"={lc()}"), ("2. Leads Worked (contacted)", f"={lc()}-{lc(st=PEND)}"),
              ("3. Leads Qualified", f"={lc(st=QUAL)}"), ("4. Qualified in Pipeline", f"={pc()}"),
              ("5. Actioned after Qualification", f"={pc()}-{pc(st=NA)}"), ("6. Won", f"={pc(st=q('Won'))}")]
    for k, (n, v) in enumerate(stages):
        rr = fr0 + k
        R.c(rr, 1, n, b=True, al=LEFT); R.c(rr, 2, v, INT, b=True)
        R.c(rr, 3, "-" if k == 0 else ("new base" if k == 3 else f"=IFERROR(B{rr}/B{rr-1},0)"), PCT)
        R.c(rr, 4, f"=IFERROR(B{rr}/$B${fr0 if k < 3 else fr0 + 3},0)", PCT)
        R.c(rr, 5, "-" if k in (0, 3) else f"=B{rr-1}-B{rr}", INT)
    R.cf(f"B{fr0}:B{fr0+5}", DataBarRule(start_type="num", start_value=0, end_type="max", color="5B9BD5"))
    R.c(fr0 + 6, 1, "Stage 4 is a separate base: the Qualified file is the CRM pipeline (it includes leads qualified from older enquiries).",
        al=LEFT, border=False, size=8, i=True, color="808080")
    R.merge(fr0 + 6, 1, 8)
    funnel_rows = (fr0, fr0 + 5)
    r = fr0 + 8

    # ---- C. Scorecard
    r = R.section(r, "C.  SALESPERSON SCORECARD  (lead handling + after-qualification work)",
                  "Green = better than team average, Red = worse. Conv. on Worked = Qualified ÷ (Leads − Pending) → true calling quality.")
    R.hdr(r, 1, ["Salesperson", "Leads Assigned", "Worked %", "Qualified", "Lead → Qual. %", "Conv. on Worked %", "Call Not Pick",
                 "Pending (Not Worked)", "Pipeline Qualified", "No Action", "Actioned", "Action Rate %", "Won", "Lost", "Win Rate %", "No Action 5+ days"])
    sc0 = r + 1
    for k, p in enumerate(c.lpeople):
        rr = sc0 + k; A = f"$A{rr}"
        R.c(rr, 1, p, b=True, al=LEFT)
        vals = {2: (f"={lc(p=A)}", INT), 3: (f"=IFERROR(1-H{rr}/B{rr},0)", PCT), 4: (f"={lc(p=A, st=QUAL)}", INT),
                5: (f"=IFERROR(D{rr}/B{rr},0)", PCT), 6: (f"=IFERROR(D{rr}/(B{rr}-H{rr}),0)", PCT),
                7: (f"={lc(p=A, st=q('Call Not Pick'))}", INT), 8: (f"={lc(p=A, st=PEND)}", INT), 9: (f"={pc(p=A)}", INT),
                10: (f"={pc(p=A, st=NA)}", INT), 11: (f"=I{rr}-J{rr}", INT), 12: (f"=IFERROR(K{rr}/I{rr},0)", PCT),
                13: (f"={pc(p=A, st=q('Won'))}", INT), 14: (f"={pc(p=A, st=q('Lost'))}", INT),
                15: (f"=IFERROR(M{rr}/I{rr},0)", PCT), 16: (f"={pc(p=A, st=NA, extra=[PDAY, q('>=5')])}", INT)}
        for cc, (fml, fm) in vals.items():
            R.c(rr, cc, fml, fm)
    st = sc0 + len(c.lpeople)
    R.total(st, "TEAM TOTAL", (2, 4, 7, 8, 9, 10, 11, 13, 14, 16),
            ((3, f"=IFERROR(1-H{st}/B{st},0)"), (5, f"=IFERROR(D{st}/B{st},0)"), (6, f"=IFERROR(D{st}/(B{st}-H{st}),0)"),
             (12, f"=IFERROR(K{st}/I{st},0)"), (15, f"=IFERROR(M{st}/I{st},0)")), sc0, st - 1)
    if c.people:
        last = sc0 + len(c.people) - 1
        for cl in "CEFLO":
            R.cf(f"{cl}{sc0}:{cl}{last}", FormulaRule(formula=[f"{cl}{sc0}>={cl}${st}"], fill=fill(GOOD)))
            R.cf(f"{cl}{sc0}:{cl}{last}", FormulaRule(formula=[f"{cl}{sc0}<{cl}${st}"], fill=fill(BAD)))
        for cl in "HP":
            R.cf(f"{cl}{sc0}:{cl}{last}", FormulaRule(formula=[f"AND({cl}{sc0}>0,{cl}{sc0}>{cl}${st}/{len(c.people)})"], fill=fill(BAD)))
    score_last = sc0 + len(c.people) - 1
    r = st + 1
    R.c(r, 1, "Check (should be 0):  scorecard − data", al=LEFT, border=False, size=8, i=True, color="808080")
    R.c(r, 2, f"=B{st}-{lc()}", INT, border=False, size=8, i=True, color="808080")
    R.c(r, 9, f"=I{st}-{pc()}", INT, border=False, size=8, i=True, color="808080")
    r += 2

    # ---- D. Salesperson x Source
    r = R.section(r, "D.  SALESPERSON × SOURCE  –  kis source se kitni leads uthayi, kitni Qualified / Not Qualified",
                  "Har salesperson ki har source ki leads ka status aur conversion.")
    oth = "Other" in c.lstat
    R.hdr(r, 1, ["Salesperson", "Source", "Leads"] + c.lstat + ["Lead → Qual. %", "Worked Leads", "Conv. on Worked %", "Share of Person's Leads"])
    stc = {s: 4 + k for k, s in enumerate(c.lstat)}
    cP, cW, cWP, cSh = (x + int(oth) for x in (8, 9, 10, 11))
    rr = r + 1; subs = []
    for p in [x for x in c.lpeople if (c.leads.Salesperson == x).any()]:
        s0 = rr
        for s in [x for x in c.lsrc if ((c.leads.Salesperson == p) & (c.leads.Source == x)).any()]:
            R.c(rr, 1, p, al=LEFT); R.c(rr, 2, s, al=LEFT)
            R.c(rr, 3, f"={lc(p=f'$A{rr}', s=f'$B{rr}')}", INT)
            for stn, cc in stc.items():
                R.c(rr, cc, f"={lc(p=f'$A{rr}', s=f'$B{rr}', st=q(stn))}", INT)
            R.c(rr, cP, f"=IFERROR(D{rr}/C{rr},0)", PCT); R.c(rr, cW, f"=C{rr}-G{rr}", INT)
            R.c(rr, cWP, f"=IFERROR(D{rr}/{L(cW)}{rr},0)", PCT)
            rr += 1
        tot = rr
        for x_ in range(s0, tot):
            R.c(x_, cSh, f"=IFERROR(C{x_}/C${tot},0)", PCT)
        R.total(tot, p, [3] + list(stc.values()) + [cW],
                ((cP, f"=IFERROR(D{tot}/C{tot},0)"), (cWP, f"=IFERROR(D{tot}/{L(cW)}{tot},0)"), (cSh, f"=IFERROR(C{tot}/C{tot},0)")),
                s0, tot - 1, bg=LIGHT, color="000000", label2="Total")
        subs.append(tot); rr += 1
    R.c(rr, 1, "GRAND TOTAL", b=True, bg=NAVY, color="FFFFFF", al=LEFT); R.c(rr, 2, "All Sources", b=True, bg=NAVY, color="FFFFFF", al=LEFT)
    for cc in [3] + list(stc.values()) + [cW]:
        R.c(rr, cc, ("=" + "+".join(f"{L(cc)}{x_}" for x_ in subs)) if subs else 0, INT, True, NAVY, color="FFFFFF")
    R.c(rr, cP, f"=IFERROR(D{rr}/C{rr},0)", PCT, True, NAVY, color="FFFFFF")
    R.c(rr, cWP, f"=IFERROR(D{rr}/{L(cW)}{rr},0)", PCT, True, NAVY, color="FFFFFF"); R.c(rr, cSh, "", bg=NAVY)
    r = rr + 2

    # ---- E. Sources
    r = R.section(r, "E.  SOURCE PERFORMANCE  –  kis source se kitni leads aayi, kitni qualify hui, aur pipeline me kya hua",
                  "Lead → Qual. % = Qualified ÷ Leads.  Conv. on Worked = Qualified ÷ worked leads.  Pipeline columns from the Qualified file.")
    R.hdr(r, 1, ["Source", "Leads", "Share of Leads", "Qualified", "Not Qualified", "Call Not Pick", "Pending (Not Worked)", "Lead → Qual. %",
                 "Conv. on Worked %", "Share of Qualified", "Unassigned", "Pipeline Qualified", "Actioned", "Action Rate %", "Won", "Win Rate %"])
    so0 = r + 1
    allsrc = c.lsrc + [s for s in c.psrc if s not in c.lsrc]
    tr = so0 + len(allsrc)
    for k, s in enumerate(allsrc):
        rr = so0 + k; A = f"$A{rr}"
        R.c(rr, 1, s, b=True, al=LEFT)
        R.c(rr, 2, f"={lc(s=A)}", INT); R.c(rr, 3, f"=IFERROR(B{rr}/$B${tr},0)", PCT)
        for j, stn in enumerate(["Qualified", "Not Qualified", "Call Not Pick", "Pending (Not Worked)"]):
            R.c(rr, 4 + j, f"={lc(s=A, st=q(stn))}", INT)
        R.c(rr, 8, f"=IFERROR(D{rr}/B{rr},0)", PCT); R.c(rr, 9, f"=IFERROR(D{rr}/(B{rr}-G{rr}),0)", PCT)
        R.c(rr, 10, f"=IFERROR(D{rr}/$D${tr},0)", PCT); R.c(rr, 11, f"={lc(p=q('Unassigned'), s=A)}", INT)
        R.c(rr, 12, f"={pc(s=A)}", INT); R.c(rr, 13, f"=L{rr}-{pc(s=A, st=NA)}", INT)
        R.c(rr, 14, f"=IFERROR(M{rr}/L{rr},0)", PCT); R.c(rr, 15, f"={pc(s=A, st=q('Won'))}", INT)
        R.c(rr, 16, f"=IFERROR(O{rr}/L{rr},0)", PCT)
    R.total(tr, "TOTAL", (2, 4, 5, 6, 7, 11, 12, 13, 15),
            ((3, f"=IFERROR(B{tr}/B{tr},0)"), (8, f"=IFERROR(D{tr}/B{tr},0)"), (9, f"=IFERROR(D{tr}/(B{tr}-G{tr}),0)"),
             (10, f"=IFERROR(D{tr}/D{tr},0)"), (14, f"=IFERROR(M{tr}/L{tr},0)"), (16, f"=IFERROR(O{tr}/L{tr},0)")), so0, tr - 1)
    for cl in "HI":
        R.cf(f"{cl}{so0}:{cl}{tr-1}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="70AD47"))
    src_last = so0 + len(c.lsrc) - 1
    r = tr + 2
    meta = c.leads[c.leads.Source == "Meta"]
    plats = sorted(set(meta["Sub-Source"]) - {"", "Meta"})
    if plats:
        R.c(r, 1, "Meta platform split", b=True, al=LEFT, border=False, color=NAVY); r += 1
        R.hdr(r, 1, ["Platform", "Leads", "Share of Meta", "Qualified", "Not Qualified", "Call Not Pick", "Pending (Not Worked)",
                     "Lead → Qual. %", "Conv. on Worked %"], TEAL, h=32)
        for k, s in enumerate(plats):
            rr = r + 1 + k; A = f"$A{rr}"
            R.c(rr, 1, s, al=LEFT); R.c(rr, 2, f"={lc(sub=A)}", INT); R.c(rr, 3, f"=IFERROR(B{rr}/{lc(s=q('Meta'))},0)", PCT)
            for j, stn in enumerate(["Qualified", "Not Qualified", "Call Not Pick", "Pending (Not Worked)"]):
                R.c(rr, 4 + j, f"={lc(sub=A, st=q(stn))}", INT)
            R.c(rr, 8, f"=IFERROR(D{rr}/B{rr},0)", PCT); R.c(rr, 9, f"=IFERROR(D{rr}/(B{rr}-G{rr}),0)", PCT)
        r = r + 2 + len(plats)
    r += 1

    # ---- F/G. Pipeline
    pp = sorted(set(c.pipe.Salesperson)); age_hdr = None
    if not len(c.pipe):
        r = R.section(r, "F–G.  AFTER-QUALIFICATION PIPELINE")
        R.c(r, 1, "Qualified pipeline file upload nahi hui – is section ke liye OakCraft-qualified file bhi upload karein.",
            al=LEFT, border=False, i=True, color="C00000")
        r += 2
    else:
        r = R.section(r, "F.  AFTER QUALIFICATION – kis salesperson ne kis lead par kya kaam kiya",
                      "No Action = status abhi bhi 'Qualified' (qualify hone ke baad kuch nahi hua). Won = lead crack. Lost = lead kaam ki nahi.")
        R.hdr(r, 1, (["Salesperson", "Source", "Qualified Leads", "No Action\n(still Qualified)"] + c.mids +
                     ["Won", "Lost", "Actioned", "Action Rate %", "Win Rate %", "Oldest No-Action (days)", "No Action 5+ days"])[:16])
        stc = {s: 4 + k for k, s in enumerate(c.pstages)}
        cA = 4 + len(c.pstages); cAR, cWR, cOld, c5 = cA + 1, cA + 2, cA + 3, cA + 4
        wonC = L(stc["Won"])
        rr = r + 1; ptot = []
        for p in pp:
            s0 = rr
            for s in [x for x in c.psrc if ((c.pipe.Salesperson == p) & (c.pipe.Source == x)).any()]:
                A, B_ = f"$A{rr}", f"$B{rr}"
                R.c(rr, 1, p, al=LEFT); R.c(rr, 2, s, al=LEFT); R.c(rr, 3, f"={pc(p=A, s=B_)}", INT)
                for stn, cc in stc.items():
                    R.c(rr, cc, f"={pc(p=A, s=B_, st=q(stn))}", INT)
                R.c(rr, cA, f"=C{rr}-D{rr}", INT); R.c(rr, cAR, f"=IFERROR({L(cA)}{rr}/C{rr},0)", PCT)
                R.c(rr, cWR, f"=IFERROR({wonC}{rr}/C{rr},0)", PCT)
                R.c(rr, cOld, f"=_xlfn.MAXIFS({PDAY},{PSP},{A},{PS},{B_},{PSTG},{NA})", DAYS)
                R.c(rr, c5, f"={pc(p=A, s=B_, st=NA, extra=[PDAY, q('>=5')])}", INT)
                rr += 1
            tot = rr
            R.total(tot, p, [3] + list(stc.values()) + [cA, c5],
                    ((cAR, f"=IFERROR({L(cA)}{tot}/C{tot},0)"), (cWR, f"=IFERROR({wonC}{tot}/C{tot},0)")),
                    s0, tot - 1, bg=LIGHT, color="000000", label2="Total")
            R.c(tot, cOld, f"=MAX({L(cOld)}{s0}:{L(cOld)}{tot-1})", DAYS, True, LIGHT)
            ptot.append(tot); rr += 1
        R.c(rr, 1, "GRAND TOTAL", b=True, bg=NAVY, color="FFFFFF", al=LEFT); R.c(rr, 2, "All Sources", b=True, bg=NAVY, color="FFFFFF", al=LEFT)
        for cc in [3] + list(stc.values()) + [cA, c5]:
            R.c(rr, cc, "=" + "+".join(f"{L(cc)}{x_}" for x_ in ptot), INT, True, NAVY, color="FFFFFF")
        R.c(rr, cAR, f"=IFERROR({L(cA)}{rr}/C{rr},0)", PCT, True, NAVY, color="FFFFFF")
        R.c(rr, cWR, f"=IFERROR({wonC}{rr}/C{rr},0)", PCT, True, NAVY, color="FFFFFF")
        R.c(rr, cOld, "=MAX(" + ",".join(f"{L(cOld)}{x_}" for x_ in ptot) + ")", DAYS, True, NAVY, color="FFFFFF")
        r = rr + 2

        r = R.section(r, "G.  NO-ACTION AGEING  &  CUSTOMER-TYPE PIPELINE",
                      "Kitne din se qualified lead par koi kaam nahi hua (as-of date se). Right: kis customer type ki lead par kitna kaam hua.")
        R.hdr(r, 1, ["Salesperson"] + [b[0] for b in AGE_BUCKETS] + ["Total No Action"], TEAL); age_hdr = r
        for k, p in enumerate(pp + ["TOTAL"]):
            rr = r + 1 + k; is_t = p == "TOTAL"; bg = LIGHT if is_t else None
            R.c(rr, 1, p, b=True, al=LEFT, bg=bg)
            for j, (_, lo, hi) in enumerate(AGE_BUCKETS):
                cc = 2 + j
                if is_t:
                    R.c(rr, cc, f"=SUM({L(cc)}{r+1}:{L(cc)}{rr-1})", INT, True, bg)
                else:
                    R.c(rr, cc, f"={pc(p=f'$A{rr}', st=NA, extra=[PDAY, q('>=' + str(lo)), PDAY, q('<=' + str(hi))])}", INT)
            R.c(rr, 7, f"=SUM(B{rr}:F{rr})", INT, True, bg)
        R.cf(f"D{r+1}:F{r+len(pp)}", FormulaRule(formula=[f"D{r+1}>0"], fill=fill(BAD)))
        R.hdr(r, 9, ["Customer Type", "", "Qualified", "No Action", "Actioned", "Action Rate %", "Purchase Qty (captured)", "Avg Qty / Lead"], TEAL)
        R.merge(r, 9, 10)
        types = c.pipe["Customer Type"].value_counts().index.tolist()
        for k, t in enumerate(types):
            rr = r + 1 + k; R.merge(rr, 9, 10)
            R.c(rr, 9, t, al=LEFT); R.c(rr, 10, None)
            R.c(rr, 11, f"=COUNTIFS({PCTY},$I{rr})", INT); R.c(rr, 12, f"=COUNTIFS({PCTY},$I{rr},{PSTG},{NA})", INT)
            R.c(rr, 13, f"=K{rr}-L{rr}", INT); R.c(rr, 14, f"=IFERROR(M{rr}/K{rr},0)", PCT)
            R.c(rr, 15, f"=SUMIFS({PQ},{PCTY},$I{rr})", INT); R.c(rr, 16, f"=IFERROR(O{rr}/K{rr},0)", '0.0;-0.0;"-"')
        rr = r + 1 + len(types); R.merge(rr, 9, 10)
        R.c(rr, 9, "TOTAL", b=True, bg=LIGHT, al=LEFT); R.c(rr, 10, None, bg=LIGHT)
        for cc in (11, 12, 13, 15):
            R.c(rr, cc, f"=SUM({L(cc)}{r+1}:{L(cc)}{rr-1})", INT, True, LIGHT)
        R.c(rr, 14, f"=IFERROR(M{rr}/K{rr},0)", PCT, True, LIGHT); R.c(rr, 16, f"=IFERROR(O{rr}/K{rr},0)", "0.0", True, LIGHT)
        r = max(r + 2 + len(pp), rr + 1) + 1

    # ---- H. Charts
    r = R.section(r, "H.  VISUAL SUMMARY")
    ch1 = r + 1; ch2 = ch1 + 15

    def bar(title, data, cats, pct=False, stacked=False, horiz=False, w=15.5, legend=False):
        ch = BarChart(); ch.type = "bar" if horiz else "col"; ch.title = title; ch.style = 10
        if stacked: ch.grouping = "stacked"; ch.overlap = 100
        ch.add_data(data, titles_from_data=True); ch.set_categories(cats)
        ch.y_axis.majorGridlines = None; ch.height = 7.2; ch.width = w
        if pct: ch.y_axis.numFmt = "0%"
        if horiz: ch.x_axis.scaling.orientation = "maxMin"
        if not legend and not stacked: ch.legend = None
        if not stacked:
            ch.dataLabels = DataLabelList(); ch.dataLabels.showVal = True
            for a in ("showSerName", "showCatName", "showLegendKey"): setattr(ch.dataLabels, a, False)
        return ch
    ref = lambda c1, r1, r2, c2=None: Reference(ws, min_col=c1, max_col=c2 or c1, min_row=r1, max_row=r2)
    ws.add_chart(bar("Conversion Funnel", ref(2, funnel_rows[0] - 1, funnel_rows[1]), ref(1, funnel_rows[0], funnel_rows[1]), horiz=True), f"A{ch1}")
    if c.people:
        ws.add_chart(bar("Lead → Qual. % vs Conv. on Worked %", ref(5, sc0 - 1, score_last, 6), ref(1, sc0, score_last), pct=True, w=11.2, legend=True), f"G{ch1}")
        ws.add_chart(bar("After-Qualification: No Action vs Actioned", ref(10, sc0 - 1, score_last, 11), ref(1, sc0, score_last), stacked=True, w=11.2), f"G{ch2}")
        ws.add_chart(bar("Call Not Pick vs Pending", ref(7, sc0 - 1, st - 1, 8), ref(1, sc0, st - 1), stacked=True, w=11.2), f"L{ch2}")
    if pp and age_hdr:
        ws.add_chart(bar("No-Action Ageing by Salesperson", ref(2, age_hdr, age_hdr + len(pp), 6), ref(1, age_hdr + 1, age_hdr + len(pp)), stacked=True, w=11.2), f"L{ch1}")
    if c.lsrc:
        ws.add_chart(bar("Lead → Qualified % by Source", ref(8, so0 - 1, src_last), ref(1, so0, src_last), pct=True), f"A{ch2}")
    r = ch2 + 16

    # ---- I. Insights
    r = R.section(r, "I.  KEY FINDINGS & ACTION PLAN", "Data se auto-generate hue findings. Har finding ke saath next step aur owner.")
    R.hdr(r, 1, ["#", "Finding"] + [""] * 8 + ["Recommended Action"] + [""] * 4 + ["Owner"], h=24)
    R.merge(r, 2, 10); R.merge(r, 11, 15)
    for k, it in enumerate(insights):
        rr = r + 1 + k
        R.merge(rr, 2, 10); R.merge(rr, 11, 15)
        R.c(rr, 1, k + 1, b=True); R.c(rr, 2, it["finding"], al=LEFT); R.c(rr, 11, it["action"], al=LEFT)
        R.c(rr, 16, it["owner"], b=True)
        for cc in list(range(3, 11)) + list(range(12, 16)):
            ws.cell(rr, cc).border = BORDER
        ws.row_dimensions[rr].height = max(32, 15 * (1 + max(len(it["finding"]) // 95, len(it["action"]) // 52)))
        if k % 2:
            for cc in range(1, 17): ws.cell(rr, cc).fill = fill(PALE)
    r = r + 2 + len(insights)

    # ---- J. Data quality
    r = R.section(r, "J.  DATA QUALITY CHECKS & ASSUMPTIONS")
    for k, it in enumerate(quality):
        rr = r + k; lab, bg = LEVEL.get(it["level"], ("ℹ Note", LIGHT))
        R.c(rr, 1, lab, b=True, bg=bg); R.merge(rr, 2, 16); R.c(rr, 2, it["text"], al=LEFT)
        for cc in range(3, 17): ws.cell(rr, cc).border = BORDER
        ws.row_dimensions[rr].height = max(18, 15 * (1 + len(it["text"]) // 170))
    r = r + len(quality) + 1

    # ---- K. Follow-up list
    r = R.section(r, "K.  PRIORITY FOLLOW-UP LIST  –  qualified leads par abhi tak koi kaam nahi (sabse purani pehle)",
                  "Critical = 5+ din, High = 3–4 din, Normal = 0–2 din. Ye list seedha salesperson ko assign karein.")
    R.hdr(r, 1, ["Salesperson", "Customer", "", "Source", "Record Id", "Contact", "Customer Type", "", "Purchase Qty",
                 "Lead Date", "Qualified On", "Days Pending", "Priority", "Next Step", "", ""], TEAL, h=30)
    R.merge(r, 2, 3); R.merge(r, 7, 8); R.merge(r, 14, 16)
    na = c.pipe[c.pipe.Stage == "No Action"].sort_values(["Days", "Purchase Qty"], ascending=[False, False])
    fl0 = r + 1
    for k, rw in enumerate(na.itertuples(index=False)):
        rr = fl0 + k
        R.merge(rr, 2, 3); R.merge(rr, 7, 8); R.merge(rr, 14, 16)
        vals = {1: rw.Salesperson, 2: rw.Customer, 4: rw.Source, 5: rw[1], 6: rw.Contact, 7: rw[6], 9: rw[7], 10: rw[2], 11: rw[3]}
        for cc in range(1, 17):
            R.c(rr, cc, vals.get(cc), "dd-mmm-yy" if cc in (10, 11) else None, al=LEFT if cc in (1, 2, 7, 14) else CEN)
        R.c(rr, 12, f"=IF(K{rr}=\"\",\"\",$O$3-K{rr})", "0")
        R.c(rr, 13, f"=IF(L{rr}=\"\",\"\",IF(L{rr}>=5,\"Critical\",IF(L{rr}>=3,\"High\",\"Normal\")))")
        R.c(rr, 14, next_step(rw[7], rw[6]), al=LEFT)
    if len(na):
        lr = fl0 + len(na) - 1
        R.cf(f"M{fl0}:M{lr}", FormulaRule(formula=[f"$L{fl0}>=5"], fill=fill(BAD), font=Font(name=FONT, bold=True, color="9C0006")))
        R.cf(f"M{fl0}:M{lr}", FormulaRule(formula=[f"AND($L{fl0}>=3,$L{fl0}<5)"], fill=fill(WARN)))
        r = lr + 2
    else:
        R.c(fl0, 1, "Sabhi qualified leads par kaam ho chuka hai.", al=LEFT, border=False, i=True); r = fl0 + 2

    # ---- L. Actioned
    act = c.pipe[c.pipe.Stage != "No Action"].sort_values(["Stage", "Salesperson"])
    r = R.section(r, f"L.  LEADS JIN PAR QUALIFICATION KE BAAD KAAM HUA ({len(act)})")
    R.hdr(r, 1, ["Salesperson", "Customer", "", "Source", "Record Id", "Stage", "Customer Type", "", "Purchase Qty",
                 "Lead Date", "Qualified On", "Contact"], TEAL, h=26)
    R.merge(r, 2, 3); R.merge(r, 7, 8)
    for k, rw in enumerate(act.itertuples(index=False)):
        rr = r + 1 + k; R.merge(rr, 2, 3); R.merge(rr, 7, 8)
        vals = {1: rw.Salesperson, 2: rw.Customer, 4: rw.Source, 5: rw[1], 6: rw.Stage, 7: rw[6], 9: rw[7], 10: rw[2], 11: rw[3], 12: rw.Contact}
        for cc in range(1, 13):
            R.c(rr, cc, vals.get(cc), "dd-mmm-yy" if cc in (10, 11) else None, al=LEFT if cc in (1, 2, 7) else CEN,
                bg={"Won": GOOD, "Lost": "D9D9D9"}.get(rw.Stage, WARN) if cc == 6 else None)
    r = r + 2 + max(len(act), 1)

    # ---- M. Unassigned
    un = c.leads[c.leads.Salesperson == "Unassigned"].sort_values("Lead Date")
    if len(un):
        r = R.section(r, f"M.  UNASSIGNED LEADS ({len(un)})  –  turant kisi salesperson ko assign karein")
        R.hdr(r, 1, ["Source", "Lead ID", "Customer", "", "Phone", "Lead Date", "Status"], TEAL, h=26); R.merge(r, 3, 4)
        for k, rw in enumerate(un.itertuples(index=False)):
            rr = r + 1 + k; R.merge(rr, 3, 4)
            for cc, v in {1: rw.Source, 2: rw[2], 3: rw.Customer, 4: None, 5: rw.Phone, 6: rw[3], 7: rw[7]}.items():
                R.c(rr, cc, v, "dd-mmm-yy hh:mm" if cc == 6 else None, al=LEFT if cc == 3 else CEN)
        r = r + 2 + len(un)

    R.c(r, 1, "Generated by OakCraft Report Studio  •  All numbers are live formulas on Data_Leads / Data_Pipeline – correct a status there and the report updates.",
        border=False, size=8, i=True, color="808080", al=LEFT)
    ws.freeze_panes = "A5"
    ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True; ws.print_title_rows = "1:2"
    ws.sheet_properties.tabColor = NAVY
    wb.calculation = CalcProperties(fullCalcOnLoad=True)

    out = io.BytesIO(); wb.save(out)
    return out.getvalue()
