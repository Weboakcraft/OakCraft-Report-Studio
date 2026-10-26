"""
Synthetic demo exports in the exact OakCraft formats (no real customer data).
Used by the "Try demo data" button and by the test-suite.
"""
from __future__ import annotations

import datetime as dt
import io
import random

import pandas as pd

TEAM = ["Amira Saikh", "Mayank Mittal", "Ujala Rajput"]
FIRST = ["Aarav", "Vivaan", "Aditya", "Neha", "Priya", "Rohan", "Karan", "Sneha", "Ishaan", "Pooja", "Rahul", "Ananya",
         "Manish", "Kavya", "Arjun", "Divya", "Sahil", "Meera", "Nikhil", "Tanvi"]
LAST = ["Sharma", "Verma", "Gupta", "Singh", "Mehta", "Jain", "Patel", "Reddy", "Nair", "Kapoor", "Bansal", "Chopra"]
PRODUCTS = ["Mesh Office Chair", "Executive Chair", "Recliner Sofa", "Gaming Chair", "Visitor Chair", "Dining Chair", "Staff Chair"]
CTYPES = ["Individual"] * 6 + ["Corporate / MNC", "Reseller", "Architect", "Hotel / Resort", "Contractor / Turnkey", None, None]


def _xlsx(df: pd.DataFrame, sheet: str) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name=sheet, index=False)
    return buf.getvalue()


def _name(rnd):
    return f"{rnd.choice(FIRST)} {rnd.choice(LAST)}"


def _phone(rnd):
    return int(f"{rnd.choice('6789')}{rnd.randint(100000000, 999999999)}")


def generate(end: dt.date | None = None, seed: int = 7):
    """Returns list of (filename, bytes) – Meta, IndiaMART, JustDial and Qualified exports."""
    rnd = random.Random(seed)
    end = end or dt.date.today() - dt.timedelta(days=1)
    start = end - dt.timedelta(days=6)
    stamp = end.strftime("%Y%m%d")

    def when():
        d = start + dt.timedelta(days=rnd.randint(0, 6))
        return dt.datetime.combine(d, dt.time(rnd.randint(8, 21), rnd.randint(0, 59), rnd.randint(0, 59)))

    meta = []
    for i in range(150):
        sp = rnd.choice(TEAM)
        w = {"Amira Saikh": [12, 25, 33, 30], "Mayank Mittal": [8, 10, 12, 70], "Ujala Rajput": [24, 22, 30, 24]}[sp]
        st = rnd.choices(["QUALIFIED", "NOT_QUALIFIED", "CALL_NOT_PICK", "CREATED"], w)[0]
        meta.append({"enquiry_id": f"ML-{rnd.randint(0, 0xFFFFFF):06X}", "created_time_ist": when().strftime("%d-%m-%Y %H:%M:%S"),
                     "platform": rnd.choice(["fb", "ig", "ig"]), "full_name": _name(rnd), "phone_number": _phone(rnd),
                     "city_state": None, "lead_status": st, "assigned_to": sp})

    im = []
    for i in range(40):
        sp = rnd.choice(TEAM + [None, None])
        st = "New" if sp is None else rnd.choices(["Qualified", "Not Qualified", "Call Not Pick", "New", "Quotation Sent"], [45, 15, 8, 25, 7])[0]
        qty = rnd.choice([1, 5, 10, 20, 25, 50, 100])
        prod = rnd.choice(PRODUCTS)
        im.append({"Query Time": when().strftime("%d-%m-%Y %H:%M:%S"), "Buyer": _name(rnd), "Company": None,
                   "Mobile": _phone(rnd), "Email": None, "City / State": rnd.choice(["Delhi", "Noida", "Pune", "Jaipur", "Bengaluru"]),
                   "Enquiry No": str(rnd.randint(1_300_000_000, 3_199_999_999)), "Product / Requirement": f"Requirement for {prod}",
                   "Quantity": qty, "Query Type": "India Mart", "Status": st, "Assigned To": sp,
                   "Note": f"Subject: Requirement for {prod} | Qty: {qty}", "Enquiry Id": None})

    jd = []
    for i in range(12):
        q = rnd.random() < 0.35
        jd.append({"Enquiry ID": f"ENQ-{rnd.randint(0, 0xFFFFFF):06X}", "Date": when().strftime("%Y-%m-%d"), "Customer": _name(rnd),
                   "Product": rnd.choice(PRODUCTS), "Qty": rnd.choice([1, 2, 10, 30]), "Source": "JustDial",
                   "Salesperson": rnd.choice(["AMIRA SAIKH", "UJALA RAJPUT"]), "Value": 0, "Priority": "Low",
                   "Stage": "Qualified" if q else "Not Qualified", "Note": 0, "Enquiry Id": None, "Enquiry Time": None,
                   "DISCUSSION": rnd.choice(["share catalogue", "fake lead", "call back", 0]),
                   "Lead's Quality": "Hot Lead" if q else "Not Interested",
                   "Qualification Status": "Qualified" if q else "Not Qualified", "Final Remark's": None})

    qual = []
    for i in range(60):
        sp = rnd.choice(TEAM)
        src = rnd.choices(["Meta", "IndiaMART", "Enquiry"], [70, 24, 6])[0]
        qd = end - dt.timedelta(days=rnd.randint(0, 9))
        ld = qd - dt.timedelta(days=rnd.randint(0, 3))
        st = "Qualified"
        if sp == "Ujala Rajput":
            st = rnd.choices(["Qualified", "Quotation Sent", "Won", "Lost"], [70, 18, 7, 5])[0]
        elif rnd.random() < 0.08:
            st = "Quotation Sent"
        pref = {"Meta": "ML", "IndiaMART": "IM", "Enquiry": "ENQ"}[src]
        qual.append({"Source": src, "Record Id": f"{pref}-{rnd.randint(0, 0xFFFFFF):06X}", "Qualified On": qd.isoformat(),
                     "Date": ld.isoformat(), "Name": _name(rnd), "Contact": float(_phone(rnd)), "Product": rnd.choice(PRODUCTS),
                     "Customer Type": rnd.choice(CTYPES), "Purchase Qty": rnd.choice([0, 0, 2, 5, 10, 15, 25, 50, 120]),
                     "Value": None, "Salesperson": sp, "Status": st})

    tag = f"{start:%Y-%m-%d}_to_{end:%Y-%m-%d}"
    return [(f"meta_leads_{stamp}.xlsx", _xlsx(pd.DataFrame(meta), "Meta")),
            (f"OakCraft-indiamart-{tag}.xlsx", _xlsx(pd.DataFrame(im), "IndiaMART")),
            (f"OakCraft-enquiries-{tag}.xlsx", _xlsx(pd.DataFrame(jd), "Enquiries")),
            (f"OakCraft-qualified-{stamp}.xlsx", _xlsx(pd.DataFrame(qual), "Qualified"))]
