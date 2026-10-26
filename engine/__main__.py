"""Command line: python -m engine FILE [FILE ...] -o report.xlsx [--asof YYYY-MM-DD] [--json summary.json]"""
import argparse
import json
import os
import sys

from . import process


def main(argv=None):
    ap = argparse.ArgumentParser(prog="python -m engine", description="OakCraft master lead report")
    ap.add_argument("files", nargs="+", help="Meta / IndiaMART / JustDial / Qualified exports (.xlsx or .csv)")
    ap.add_argument("-o", "--out", default="OakCraft_Master_Lead_Report.xlsx")
    ap.add_argument("--asof", help="Report date YYYY-MM-DD (default: today)")
    ap.add_argument("--json", help="Also write the dashboard summary JSON here")
    a = ap.parse_args(argv)

    files = []
    for p in a.files:
        with open(p, "rb") as fh:
            files.append((os.path.basename(p), fh.read()))
    summary, xlsx = process(files, a.asof)
    if xlsx is None:
        print("No recognised report files – nothing to build.", file=sys.stderr)
        return 1
    with open(a.out, "wb") as fh:
        fh.write(xlsx)
    if a.json:
        with open(a.json, "w", encoding="utf-8") as fh:
            json.dump(summary, fh, ensure_ascii=False, indent=1)
    k = summary["kpi"]
    print(f"Saved {a.out}")
    print("Detected:", ", ".join(f"{f['label']} ← {f['name']}" for f in summary["meta"]["files"]))
    print(f"Leads {k['leads']} | Qualified {k['qualified']} ({k['conv']:.1%}) | Pipeline {k['pipe_total']} | "
          f"Actioned {k['actioned']} | Won {k['won']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
