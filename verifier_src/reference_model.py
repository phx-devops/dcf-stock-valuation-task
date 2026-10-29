"""Reference DCF model for the nke-dcf-valuation task (verifier side).

Pure Python, no third-party imports. It mirrors, step by step, the Excel formulas the
golden workbook must contain, and it defines the label contract the verifier uses to
locate output cells.

Conventions
-----------
* Money in millions of USD, shares in millions, per-share values in USD.
* Series are ordered oldest -> latest: index 0 = FY-2, 1 = FY-1, 2 = FY0 (latest 10-K year).
* Input ids look like ``key|period`` (see ``required_ids``).
"""
import json

PERIODS = ["FY-2", "FY-1", "FY0"]
FY_KEYS = [
    "revenue", "pretax_income", "interest_net", "tax_expense", "depreciation", "capex",
    "receivables", "inventories", "prepaid", "payables", "accrued",
]
YTD_KEYS = ["revenue", "pretax_income", "interest_net", "tax_expense"]
BRIDGE_KEYS = ["cash", "short_term_investments", "notes_payable", "current_ltd", "long_term_debt"]
DEBT_KEYS = ["notes_payable", "current_ltd", "long_term_debt"]
N_YEARS = 10


def required_ids():
    """The 50 Inputs ids the solver must provide."""
    ids = [f"{k}|{p}" for p in PERIODS for k in FY_KEYS]
    ids += [f"{k}|FY0" for k in BRIDGE_KEYS]
    ids += [f"{k}|{p}" for p in ("YTD_CUR", "YTD_PRI") for k in YTD_KEYS]
    ids += ["diluted_shares|YTD_CUR"]
    ids += [f"{k}|Q_END" for k in DEBT_KEYS]
    return ids


# Label contract: (sheet, label) -> number of value columns (0 = scalar in column B,
# otherwise a series laid out in columns B.. of that many columns).
OUTPUT_SPEC = {}
for _l in ["Revenue", "EBIT", "Revenue growth", "EBIT margin", "Effective tax rate",
           "D&A % of revenue", "Capex % of revenue", "NWC", "NWC % of revenue",
           "Net investment"]:
    OUTPUT_SPEC[("Historicals", _l)] = 3
for _l in ["Avg EBIT margin", "Avg effective tax rate", "Avg D&A % of revenue",
           "Avg capex % of revenue", "Avg NWC % of revenue", "Revenue CAGR", "LTM revenue"]:
    OUTPUT_SPEC[("Historicals", _l)] = 0
for _l in ["Cost of equity", "After-tax cost of debt", "Market cap", "Debt",
           "Equity weight", "Debt weight", "WACC"]:
    OUTPUT_SPEC[("WACC", _l)] = 0
for _l in ["Revenue growth", "Revenue", "EBIT", "NOPAT", "Capex", "D&A", "NWC",
           "Change in NWC", "Net investment", "Free cash flow", "Discount factor", "PV of FCF"]:
    OUTPUT_SPEC[("DCF", _l)] = N_YEARS
for _l in ["Sum of PV of FCF", "Terminal value", "PV of terminal value", "Enterprise value",
           "Cash", "Debt", "Equity value", "Diluted shares", "Value per share",
           "Market price", "Upside"]:
    OUTPUT_SPEC[("DCF", _l)] = 0

# summary.json field -> (sheet, label) of the workbook cell it must equal
SUMMARY_MAP = {
    "wacc": ("WACC", "WACC"),
    "enterprise_value": ("DCF", "Enterprise value"),
    "equity_value": ("DCF", "Equity value"),
    "value_per_share": ("DCF", "Value per share"),
    "market_price": ("DCF", "Market price"),
    "upside_pct": ("DCF", "Upside"),
}


def assumptions_from_json(raw):
    """Flatten the pinned assumptions.json into the keys the model and workbook use."""
    return {
        "risk_free_rate": raw["wacc"]["risk_free_rate"],
        "equity_risk_premium": raw["wacc"]["equity_risk_premium"],
        "beta": raw["wacc"]["beta"],
        "pre_tax_cost_of_debt": raw["wacc"]["pre_tax_cost_of_debt"],
        "terminal_growth_rate": raw["terminal_value"]["growth_rate"],
        "explicit_forecast_years": raw["forecast"]["explicit_forecast_years"],
    }


def inputs_from_values(values, price):
    """Turn an ``{id: number}`` map into the model's input structure."""
    g = lambda key, p: values[f"{key}|{p}"]
    inp = {
        "rev": [g("revenue", p) for p in PERIODS],
        "pretax": [g("pretax_income", p) for p in PERIODS],
        "int_net": [g("interest_net", p) for p in PERIODS],
        "tax": [g("tax_expense", p) for p in PERIODS],
        "da": [g("depreciation", p) for p in PERIODS],
        "capex": [g("capex", p) for p in PERIODS],
        "ar": [g("receivables", p) for p in PERIODS],
        "inv": [g("inventories", p) for p in PERIODS],
        "prepaid": [g("prepaid", p) for p in PERIODS],
        "ap": [g("payables", p) for p in PERIODS],
        "accrued": [g("accrued", p) for p in PERIODS],
        "rev_ytd_cur": g("revenue", "YTD_CUR"),
        "rev_ytd_pri": g("revenue", "YTD_PRI"),
        "debt_q": sum(g(k, "Q_END") for k in DEBT_KEYS),
        "shares_q": g("diluted_shares", "YTD_CUR"),
        "price": price,
        "cash_k": g("cash", "FY0") + g("short_term_investments", "FY0"),
        "debt_k": sum(g(k, "FY0") for k in DEBT_KEYS),
    }
    return inp


def build_model(inp, a):
    n = 3
    avg = lambda xs: sum(xs) / len(xs)
    yrs = a["explicit_forecast_years"]
    assert yrs == N_YEARS, "the task fixes a 10-year explicit forecast"

    # ---- Historicals -------------------------------------------------------
    ebit = [inp["pretax"][i] + inp["int_net"][i] for i in range(n)]
    growth_h = [None] + [inp["rev"][i] / inp["rev"][i - 1] - 1 for i in (1, 2)]
    margin = [ebit[i] / inp["rev"][i] for i in range(n)]
    tax_rate = [inp["tax"][i] / inp["pretax"][i] for i in range(n)]
    da_pct = [inp["da"][i] / inp["rev"][i] for i in range(n)]
    capex_pct = [inp["capex"][i] / inp["rev"][i] for i in range(n)]
    nwc = [inp["ar"][i] + inp["inv"][i] + inp["prepaid"][i] - inp["ap"][i] - inp["accrued"][i]
           for i in range(n)]
    nwc_pct = [nwc[i] / inp["rev"][i] for i in range(n)]
    net_inv_h = [None] + [inp["capex"][i] - inp["da"][i] + (nwc[i] - nwc[i - 1]) for i in (1, 2)]
    m_avg, t_avg, da_avg, cx_avg, nwc_avg = (
        avg(margin), avg(tax_rate), avg(da_pct), avg(capex_pct), avg(nwc_pct))

    # ---- LTM base and growth path -----------------------------------------
    ltm_rev = inp["rev"][2] + inp["rev_ytd_cur"] - inp["rev_ytd_pri"]
    cagr = (inp["rev"][2] / inp["rev"][0]) ** 0.5 - 1
    gT = a["terminal_growth_rate"]
    growth = [cagr + (gT - cagr) * (t - 1) / (yrs - 1) for t in range(1, yrs + 1)]

    # ---- WACC (10-Q capital structure) ------------------------------------
    E = inp["price"] * inp["shares_q"]
    D = inp["debt_q"]
    ke = a["risk_free_rate"] + a["beta"] * a["equity_risk_premium"]
    kd_after = a["pre_tax_cost_of_debt"] * (1 - t_avg)
    we, wd = E / (D + E), D / (D + E)
    wacc = we * ke + wd * kd_after
    assert wacc > gT, "WACC must exceed terminal growth"

    # ---- 10-year projection -----------------------------------------------
    rows = {k: [] for k in ["rev", "ebit", "nopat", "capex", "da", "nwc", "dnwc",
                            "net_inv", "fcf", "df", "pv"]}
    prev_rev, prev_nwc = ltm_rev, nwc_avg * ltm_rev  # NWC_0 = avg NWC % * LTM revenue
    for t, g in enumerate(growth, start=1):
        r = prev_rev * (1 + g)
        e = m_avg * r
        nopat = e * (1 - t_avg)
        cx, da = cx_avg * r, da_avg * r
        nwc_t = nwc_avg * r
        dn = nwc_t - prev_nwc
        ni = cx - da + dn
        f = nopat - ni
        d = 1 / (1 + wacc) ** t
        for k, v in zip(rows, [r, e, nopat, cx, da, nwc_t, dn, ni, f, d, f * d]):
            rows[k].append(v)
        prev_rev, prev_nwc = r, nwc_t

    tv = rows["fcf"][-1] * (1 + gT) / (wacc - gT)  # at end of year 10
    pv_tv = tv * rows["df"][-1]
    sum_pv = sum(rows["pv"])
    ev = sum_pv + pv_tv
    equity = ev + inp["cash_k"] - inp["debt_k"]
    vps = equity / inp["shares_q"]
    return {
        "hist": dict(ebit=ebit, growth=growth_h, margin=margin, tax_rate=tax_rate,
                     da_pct=da_pct, capex_pct=capex_pct, nwc=nwc, nwc_pct=nwc_pct,
                     net_inv=net_inv_h, m_avg=m_avg, t_avg=t_avg, da_avg=da_avg,
                     cx_avg=cx_avg, nwc_avg=nwc_avg, cagr=cagr, ltm_rev=ltm_rev,
                     rev=list(inp["rev"])),
        "wacc": dict(ke=ke, kd_after=kd_after, E=E, D=D, we=we, wd=wd, wacc=wacc),
        "growth": growth, "rows": rows, "sum_pv": sum_pv, "tv": tv, "pv_tv": pv_tv,
        "ev": ev, "cash_k": inp["cash_k"], "debt_k": inp["debt_k"], "equity": equity,
        "shares": inp["shares_q"], "vps": vps, "price": inp["price"],
        "upside": vps / inp["price"] - 1,
    }


def expected_outputs(m):
    """Map every (sheet, label) in OUTPUT_SPEC to the reference value (list or scalar)."""
    h, w, r = m["hist"], m["wacc"], m["rows"]
    o = {
        ("Historicals", "Revenue"): h["rev"],
        ("Historicals", "EBIT"): h["ebit"],
        ("Historicals", "Revenue growth"): h["growth"],
        ("Historicals", "EBIT margin"): h["margin"],
        ("Historicals", "Effective tax rate"): h["tax_rate"],
        ("Historicals", "D&A % of revenue"): h["da_pct"],
        ("Historicals", "Capex % of revenue"): h["capex_pct"],
        ("Historicals", "NWC"): h["nwc"],
        ("Historicals", "NWC % of revenue"): h["nwc_pct"],
        ("Historicals", "Net investment"): h["net_inv"],
        ("Historicals", "Avg EBIT margin"): h["m_avg"],
        ("Historicals", "Avg effective tax rate"): h["t_avg"],
        ("Historicals", "Avg D&A % of revenue"): h["da_avg"],
        ("Historicals", "Avg capex % of revenue"): h["cx_avg"],
        ("Historicals", "Avg NWC % of revenue"): h["nwc_avg"],
        ("Historicals", "Revenue CAGR"): h["cagr"],
        ("Historicals", "LTM revenue"): h["ltm_rev"],
        ("WACC", "Cost of equity"): w["ke"],
        ("WACC", "After-tax cost of debt"): w["kd_after"],
        ("WACC", "Market cap"): w["E"],
        ("WACC", "Debt"): w["D"],
        ("WACC", "Equity weight"): w["we"],
        ("WACC", "Debt weight"): w["wd"],
        ("WACC", "WACC"): w["wacc"],
        ("DCF", "Revenue growth"): m["growth"],
        ("DCF", "Revenue"): r["rev"],
        ("DCF", "EBIT"): r["ebit"],
        ("DCF", "NOPAT"): r["nopat"],
        ("DCF", "Capex"): r["capex"],
        ("DCF", "D&A"): r["da"],
        ("DCF", "NWC"): r["nwc"],
        ("DCF", "Change in NWC"): r["dnwc"],
        ("DCF", "Net investment"): r["net_inv"],
        ("DCF", "Free cash flow"): r["fcf"],
        ("DCF", "Discount factor"): r["df"],
        ("DCF", "PV of FCF"): r["pv"],
        ("DCF", "Sum of PV of FCF"): m["sum_pv"],
        ("DCF", "Terminal value"): m["tv"],
        ("DCF", "PV of terminal value"): m["pv_tv"],
        ("DCF", "Enterprise value"): m["ev"],
        ("DCF", "Cash"): m["cash_k"],
        ("DCF", "Debt"): m["debt_k"],
        ("DCF", "Equity value"): m["equity"],
        ("DCF", "Diluted shares"): m["shares"],
        ("DCF", "Value per share"): m["vps"],
        ("DCF", "Market price"): m["price"],
        ("DCF", "Upside"): m["upside"],
    }
    assert set(o) == set(OUTPUT_SPEC)
    return o


def model_from(values, assumptions, price):
    return build_model(inputs_from_values(values, price), assumptions)


def summary_from(m):
    return {"wacc": m["wacc"]["wacc"], "enterprise_value": m["ev"], "equity_value": m["equity"],
            "value_per_share": m["vps"], "market_price": m["price"], "upside_pct": m["upside"]}


def _selfcheck():
    """Identity tests on illustrative numbers (not Nike data)."""
    A = dict(risk_free_rate=0.042, equity_risk_premium=0.05, beta=1.10,
             pre_tax_cost_of_debt=0.045, terminal_growth_rate=0.025, explicit_forecast_years=10)
    g = A["terminal_growth_rate"]
    inp = dict(
        rev=[46000, 49000, 51000], pretax=[5500, 5000, 4800], int_net=[-100, -50, 20],
        tax=[900, 850, 800], da=[700, 720, 740], capex=[800, 850, 900],
        ar=[4000, 4200, 4300], inv=[8000, 7500, 7400], prepaid=[2000, 2100, 2200],
        ap=[2500, 2600, 2700], accrued=[5000, 5200, 5400], rev_ytd_cur=12000,
        rev_ytd_pri=11800, debt_q=8000, shares_q=1480, price=70.0, cash_k=9000, debt_k=8000)
    m = build_model(inp, A)
    assert abs(m["growth"][0] - m["hist"]["cagr"]) < 1e-12 and abs(m["growth"][-1] - g) < 1e-12
    assert all(m["rows"]["df"][i] > m["rows"]["df"][i + 1] for i in range(9))
    assert abs(m["ev"] - (sum(m["rows"]["pv"]) + m["pv_tv"])) < 1e-6
    # closed form: constant growth at g from the base -> EV = FCF1 / (WACC - g)
    grow = lambda x: [x, x * (1 + g), x * (1 + g) ** 2]
    inp2 = dict(inp, rev=grow(100000), pretax=grow(10000), int_net=[0, 0, 0], tax=grow(1800),
                da=grow(1500), capex=grow(2000), ar=grow(5000), inv=grow(9000),
                prepaid=grow(2000), ap=grow(4000), accrued=grow(6000),
                rev_ytd_cur=0, rev_ytd_pri=0)
    m2 = build_model(inp2, A)
    closed = m2["rows"]["fcf"][0] / (m2["wacc"]["wacc"] - g)
    assert abs(m2["ev"] - closed) / closed < 1e-9
    m3 = build_model(dict(inp, cash_k=inp["cash_k"] + 100), A)
    assert abs((m3["equity"] - m["equity"]) - 100) < 1e-9
    assert set(expected_outputs(m)) == set(OUTPUT_SPEC)
    return m


if __name__ == "__main__":
    m = _selfcheck()
    print("ALL CHECKS PASSED")
    print(f"WACC={m['wacc']['wacc']:.4%} EV={m['ev']:.0f} Equity={m['equity']:.0f} "
          f"VPS={m['vps']:.2f} TV share={m['pv_tv'] / m['ev']:.1%}")
