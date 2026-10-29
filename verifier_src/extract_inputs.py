"""Extract the 50 DCF inputs from the Nike filings and write inputs.json.

Sources
  * three 10-K text files (printed page numbers are taken from the page footers)
  * the 10-Q PDF (converted with ``pdftotext -layout``; page = 1-based PDF page)

Rule: every figure comes from the NEWEST filing that contains it (restatements win).
Also logs, for every figure, whether older filings that contain it agree.

Usage:
  python extract_inputs.py --filings-dir DIR --out inputs.json
File names are matched by suffix: *10K_2024*, *10K_2025*, *2026-10K*, *10Q_2027*.
"""
import argparse
import json
import re
import subprocess
import tempfile
from pathlib import Path

NUM = re.compile(r"\(?\$?\s*\d[\d,]*(?:\.\d+)?\)?|—")


def parse_line(line):
    """'Revenues $ 46,398 $ 46,309 $ 51,362' -> ('Revenues', [46398.0, 46309.0, 51362.0])"""
    line = line.replace("$", " ")
    first = NUM.search(line)
    if not first:
        return None
    label = line[: first.start()].strip()
    nums = []
    for tok in NUM.findall(line[first.start():]):
        tok = tok.strip()
        if tok == "—":
            nums.append(0.0)
            continue
        neg = tok.startswith("(") and tok.endswith(")")
        v = float(tok.strip("()").replace(",", "").strip())
        nums.append(-v if neg else v)
    return label, nums


def statement_block(lines, heading, start_at=0):
    """Lines from the exact heading to the 'accompanying Notes' sentence, plus printed page."""
    idx = next(i for i in range(start_at, len(lines)) if lines[i].strip() == heading)
    end = next(i for i in range(idx, len(lines)) if lines[i].startswith("The accompanying Notes"))
    page = None
    for j in range(end, min(end + 6, len(lines))):
        m = re.search(r"FORM 10-K (\d+)\s*$", lines[j]) or re.match(r"^(\d+) NIKE, INC", lines[j])
        if m:
            page = int(m.group(1))
            break
    assert page, f"no page footer after {heading}"
    return lines[idx:end], page


def rows(block):
    out = {}
    for ln in block:
        p = parse_line(ln)
        if p and p[0] and p[1]:
            out.setdefault(p[0], p[1])  # first occurrence wins
    return out


def load_10k(path):
    lines = Path(path).read_text().splitlines()
    res = {}
    for key, heading in [("IS", "CONSOLIDATED STATEMENTS OF INCOME"),
                         ("BS", "CONSOLIDATED BALANCE SHEETS"),
                         ("CF", "CONSOLIDATED STATEMENTS OF CASH FLOWS")]:
        block, page = statement_block(lines, heading)
        res[key] = {"rows": rows(block), "page": page}
    return res


def load_10q(pdf):
    work = Path(tempfile.mkdtemp(prefix="q10_"))
    txt = work / "q.txt"
    subprocess.run(["pdftotext", "-layout", str(pdf), str(txt)], check=True)
    pages = txt.read_text().split("\f")
    res = {}
    for key, needle in [("IS", "UNAUDITED CONDENSED CONSOLIDATED STATEMENTS OF INCOME"),
                        ("BS", "UNAUDITED CONDENSED CONSOLIDATED BALANCE SHEETS")]:
        # statement pages: heading present AND a data row present (skips the index page)
        hits = [i + 1 for i, p in enumerate(pages)
                if needle in p and ("Revenues" in p or "Cash and equivalents" in p)]
        assert len(hits) == 1, (key, hits)
        res[key] = {"rows": rows(pages[hits[0] - 1].splitlines()), "page": hits[0]}
    return res


def find(files, pattern):
    m = [p for p in Path(files).iterdir() if re.search(pattern, p.name)]
    assert len(m) == 1, (pattern, m)
    return m[0]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--filings-dir", required=True)
    ap.add_argument("--source-filings-dir")
    ap.add_argument("--out", required=True)
    a = ap.parse_args()

    f24 = find(a.filings_dir, r"10K_2024")
    f25 = find(a.filings_dir, r"10K_2025")
    f26 = find(a.filings_dir, r"2026-10K")
    fq = find(a.filings_dir, r"10Q_2027")
    source_aliases = {}
    if a.source_filings_dir:
        for pattern in (r"10K_2024", r"10K_2025", r"2026-10K", r"10Q_2027"):
            temporary = find(a.filings_dir, pattern)
            original = find(a.source_filings_dir, pattern)
            source_aliases[temporary.name] = original.name
    k24, k25, k26, q = load_10k(f24), load_10k(f25), load_10k(f26), load_10q(fq)

    # statement line items: (id key, statement, label in newest filing, label in older filings)
    IS = [("revenue", "Revenues"), ("pretax_income", "Income before income taxes"),
          ("tax_expense", "Income tax expense")]
    values, log = {}, []

    def put(id_, value, src, page, note=None):
        src = source_aliases.get(src, src)
        values[id_] = {"value": value, "source_file": src, "page": page}
        if note:
            values[id_]["note"] = note

    def pick(filing, stmt, label, col):
        return filing[stmt]["rows"][label][col]

    # ---- income statement + cash flow: FY0/FY-1/FY-2 all from the FY2026 10-K ------------
    fy_cols = {"FY0": 0, "FY-1": 1, "FY-2": 2}
    older = {"FY-1": [(k25, f25, 0)], "FY-2": [(k25, f25, 1), (k24, f24, 0)]}
    for p, c in fy_cols.items():
        for key, label in IS:
            put(f"{key}|{p}", pick(k26, "IS", label, c), f26.name, k26["IS"]["page"])
        # interest label differs by year-of-filing; match on prefix
        ilabel26 = next(l for l in k26["IS"]["rows"] if l.startswith("Interest"))
        put(f"interest_net|{p}", pick(k26, "IS", ilabel26, c), f26.name, k26["IS"]["page"],
            "reported sign: negative = net interest income")
        dlabel = next(l for l in k26["CF"]["rows"] if l.startswith("Depreciation"))
        put(f"depreciation|{p}", pick(k26, "CF", dlabel, c), f26.name, k26["CF"]["page"],
            "cash-flow line 'Depreciation and amortization' (labelled 'Depreciation' in the FY2024 10-K; same values)")
        put(f"capex|{p}", -pick(k26, "CF", "Additions to property, plant and equipment", c),
            f26.name, k26["CF"]["page"], "shown as a positive magnitude")

    # ---- balance sheet --------------------------------------------------------------------
    bs_items = [("receivables", "Accounts receivable, net"), ("inventories", "Inventories"),
                ("prepaid", "Prepaid expenses and other current assets"),
                ("payables", "Accounts payable"), ("accrued", "Accrued liabilities")]
    for key, label in bs_items:
        put(f"{key}|FY0", pick(k26, "BS", label, 0), f26.name, k26["BS"]["page"])
        put(f"{key}|FY-1", pick(k26, "BS", label, 1), f26.name, k26["BS"]["page"])
        # FY2024 balance sheet exists only in the FY2025 and FY2024 10-Ks -> newest is FY2025
        put(f"{key}|FY-2", pick(k25, "BS", label, 1), f25.name, k25["BS"]["page"])
    values["accrued|FY-2"]["note"] = (
        "FY2025 10-K basis; excludes $6M notes payable that the FY2026 10-K folds into accrued "
        "liabilities from FY2025 onward; FY2024 not restated by the company")
    values["accrued|FY-1"]["note"] = "restated in FY2026 10-K (5,916 vs 5,911 in the FY2025 10-K)"

    put("cash|FY0", pick(k26, "BS", "Cash and equivalents", 0), f26.name, k26["BS"]["page"])
    put("short_term_investments|FY0", pick(k26, "BS", "Short-term investments", 0),
        f26.name, k26["BS"]["page"])
    put("current_ltd|FY0", pick(k26, "BS", "Current portion of long-term debt", 0),
        f26.name, k26["BS"]["page"])
    put("long_term_debt|FY0", pick(k26, "BS", "Long-term debt", 0), f26.name, k26["BS"]["page"])
    assert "Notes payable" not in k26["BS"]["rows"]
    put("notes_payable|FY0", 0.0, f26.name, k26["BS"]["page"],
        "not presented separately in the FY2026 10-K balance sheet (inferred folded into accrued liabilities)")

    # ---- 10-Q ---------------------------------------------------------------------------
    for p, c in [("YTD_CUR", 0), ("YTD_PRI", 1)]:
        put(f"revenue|{p}", pick(q, "IS", "Revenues", c), fq.name, q["IS"]["page"])
        put(f"pretax_income|{p}", pick(q, "IS", "Income before income taxes", c), fq.name, q["IS"]["page"])
        put(f"tax_expense|{p}", pick(q, "IS", "Income tax expense", c), fq.name, q["IS"]["page"])
        il = next(l for l in q["IS"]["rows"] if l.startswith("Interest"))
        put(f"interest_net|{p}", pick(q, "IS", il, c), fq.name, q["IS"]["page"],
            "reported sign: negative = net interest income")
    # 'Diluted' appears twice: the EPS row, then the weighted-average shares row (second)
    txt =subprocess.run(["pdftotext", "-layout", "-f", str(q["IS"]["page"]), "-l", str(q["IS"]["page"]),
                          str(fq), "-"], capture_output=True, text=True, check=True).stdout
    dil = [parse_line(l) for l in txt.splitlines() if l.strip().startswith("Diluted")]
    dil = [d for d in dil if d]
    assert len(dil) == 2, dil  # EPS row then weighted-average shares row
    put("diluted_shares|YTD_CUR", dil[1][1][0], fq.name, q["IS"]["page"],
        "weighted-average diluted shares, three months ended Aug 31, 2026 (millions)")
    put("current_ltd|Q_END", pick(q, "BS", "Current portion of long-term debt", 0), fq.name, q["BS"]["page"])
    put("long_term_debt|Q_END", pick(q, "BS", "Long-term debt", 0), fq.name, q["BS"]["page"])
    assert "Notes payable" not in q["BS"]["rows"]
    put("notes_payable|Q_END", 0.0, fq.name, q["BS"]["page"],
        "not presented separately in the 10-Q balance sheet")

    # ---- cross-filing consistency log (older filings that contain the same figure) --------
    def cmp(id_, newest, others):
        for name, v in others:
            log.append({"id": id_, "newest": newest,
                        "older_file": source_aliases.get(name, name), "older_value": v,
                        "differs": abs(newest - v) > 1e-9})
    for key, label in IS:
        for p, c25, c24 in [("FY-1", 1, None), ("FY-2", 2, None)]:
            pass
    for key, label in IS:
        cmp(f"{key}|FY-1", values[f"{key}|FY-1"]["value"], [(f25.name, pick(k25, "IS", label, 0))])
        cmp(f"{key}|FY-2", values[f"{key}|FY-2"]["value"],
            [(f25.name, pick(k25, "IS", label, 1)), (f24.name, pick(k24, "IS", label, 0))])
    for key, label in bs_items:
        cmp(f"{key}|FY-1", values[f"{key}|FY-1"]["value"], [(f25.name, pick(k25, "BS", label, 0))])
        cmp(f"{key}|FY-2", values[f"{key}|FY-2"]["value"], [(f24.name, pick(k24, "BS", label, 0))])
    cmp("capex|FY-1", values["capex|FY-1"]["value"], [(f25.name, -pick(k25, "CF", "Additions to property, plant and equipment", 0))])
    cmp("capex|FY-2", values["capex|FY-2"]["value"],
        [(f25.name, -pick(k25, "CF", "Additions to property, plant and equipment", 1)),
         (f24.name, -pick(k24, "CF", "Additions to property, plant and equipment", 0))])
    d25 = next(l for l in k25["CF"]["rows"] if l.startswith("Depreciation"))
    d24 = next(l for l in k24["CF"]["rows"] if l.startswith("Depreciation"))
    cmp("depreciation|FY-1", values["depreciation|FY-1"]["value"], [(f25.name, pick(k25, "CF", d25, 0))])
    cmp("depreciation|FY-2", values["depreciation|FY-2"]["value"],
        [(f25.name, pick(k25, "CF", d25, 1)), (f24.name, pick(k24, "CF", d24, 0))])

    ordered = {}
    import sys
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "tests"))
    import reference_model as rm
    for i in rm.required_ids():
        ordered[i] = values[i]
    assert set(values) == set(rm.required_ids()), set(values) ^ set(rm.required_ids())

    doc = {
        "_meta": {
            "illustrative": False,
            "company": "Nike, Inc. (NKE)",
            "units": "millions of USD; diluted_shares in millions of shares",
            "periods": {"FY-2": "fiscal year ended May 31, 2024", "FY-1": "fiscal year ended May 31, 2025",
                        "FY0": "fiscal year ended May 31, 2026",
                        "YTD_CUR": "three months ended Aug 31, 2026", "YTD_PRI": "three months ended Aug 31, 2025",
                        "Q_END": "balance sheet at Aug 31, 2026"},
            "page_basis": "10-K files are text conversions with no PDF page breaks, so 'page' is the printed 10-K page number from the page footer; the 10-Q 'page' is the 1-based page of the supplied PDF.",
            "sign_conventions": "capex is a positive magnitude; interest_net keeps the reported sign (negative = net interest income)",
            "restatements": "each figure comes from the newest filing that contains it",
            "restatement_differences": [x for x in log if x["differs"]],
            "cross_filing_checks": len(log),
        },
        "values": ordered,
    }
    Path(a.out).write_text(json.dumps(doc, indent=2))
    print(f"wrote {a.out}: {len(ordered)} ids, {len(log)} cross-filing comparisons, "
          f"{sum(x['differs'] for x in log)} differences")
    for x in log:
        if x["differs"]:
            print("  DIFF", x)


if __name__ == "__main__":
    main()