"""Reference implementation of the valuation model, used to generate the
golden values that the VBA test runner checks itself against.

This exists because VBA has no usable unit testing story and no CI runner. The
arithmetic is specified once here, in a language that can be run in a pipeline,
and `tests/golden_values.csv` is the contract between the two. If somebody
changes a formula in `modDCF.bas`, the VBA test runner reports a mismatch
against a number that was computed independently rather than against another
copy of the same possibly-wrong formula.

Run:  python tools/generate_golden.py
"""

from __future__ import annotations

import csv
from dataclasses import dataclass, field
from pathlib import Path

OUT = Path(__file__).resolve().parents[1] / "tests" / "golden_values.csv"


@dataclass(frozen=True)
class Assumptions:
    """The base case. Figures in millions unless noted."""

    revenue_0: float = 4_820.0
    revenue_growth: tuple[float, ...] = (0.082, 0.071, 0.060, 0.050, 0.042)
    ebit_margin: tuple[float, ...] = (0.184, 0.189, 0.193, 0.196, 0.198)
    tax_rate: float = 0.24
    depreciation_pct_revenue: float = 0.041
    capex_pct_revenue: float = 0.052
    nwc_pct_revenue: float = 0.113

    # Capital structure and cost of capital
    risk_free_rate: float = 0.0428
    equity_risk_premium: float = 0.0510
    levered_beta: float = 1.18
    pretax_cost_of_debt: float = 0.0565
    market_cap: float = 11_430.0
    total_debt: float = 3_260.0
    cash: float = 742.0

    terminal_growth: float = 0.024
    exit_ebitda_multiple: float = 11.5
    diluted_shares: float = 214.6
    #: Mid-year convention: cash flows arrive through the year, not on 31 Dec.
    mid_year_convention: bool = True

    @property
    def net_debt(self) -> float:
        return self.total_debt - self.cash

    @property
    def enterprise_value_market(self) -> float:
        return self.market_cap + self.net_debt


def cost_of_equity(a: Assumptions) -> float:
    """CAPM."""
    return a.risk_free_rate + a.levered_beta * a.equity_risk_premium


def wacc(a: Assumptions) -> float:
    """Weighted average cost of capital on *market* values, not book.

    Book equity is an accounting residual and has nothing to do with what
    investors require; using it is the single most common error in a WACC.
    """
    equity = a.market_cap
    debt = a.total_debt
    total = equity + debt
    return (equity / total) * cost_of_equity(a) + (debt / total) * a.pretax_cost_of_debt * (1 - a.tax_rate)


def unlever_beta(levered: float, debt_to_equity: float, tax_rate: float) -> float:
    """Hamada. Strips the financing risk out of an observed equity beta."""
    return levered / (1.0 + (1.0 - tax_rate) * debt_to_equity)


def relever_beta(unlevered: float, debt_to_equity: float, tax_rate: float) -> float:
    return unlevered * (1.0 + (1.0 - tax_rate) * debt_to_equity)


@dataclass
class Projection:
    year: int
    revenue: float
    ebit: float
    nopat: float
    depreciation: float
    capex: float
    change_in_nwc: float
    fcff: float
    discount_factor: float
    present_value: float


def project(a: Assumptions) -> list[Projection]:
    """Unlevered free cash flow, year by year.

    FCFF = EBIT(1-t) + D&A - capex - change in net working capital.

    The working capital term is a *change*, not a level. Using the level is a
    mistake that silently understates cash flow by roughly a year of revenue
    times the NWC ratio, which on this company is about $545m a year.
    """
    rows: list[Projection] = []
    revenue = a.revenue_0
    nwc = a.nwc_pct_revenue * revenue
    discount = wacc(a)

    for i, (growth, margin) in enumerate(zip(a.revenue_growth, a.ebit_margin, strict=True), start=1):
        revenue = revenue * (1.0 + growth)
        ebit = revenue * margin
        nopat = ebit * (1.0 - a.tax_rate)
        depreciation = revenue * a.depreciation_pct_revenue
        capex = revenue * a.capex_pct_revenue

        new_nwc = revenue * a.nwc_pct_revenue
        change_in_nwc = new_nwc - nwc
        nwc = new_nwc

        fcff = nopat + depreciation - capex - change_in_nwc

        exponent = (i - 0.5) if a.mid_year_convention else float(i)
        factor = 1.0 / (1.0 + discount) ** exponent

        rows.append(
            Projection(i, revenue, ebit, nopat, depreciation, capex, change_in_nwc,
                       fcff, factor, fcff * factor)
        )
    return rows


def terminal_value_gordon(a: Assumptions, rows: list[Projection]) -> float:
    """Gordon growth on the final-year cash flow."""
    discount = wacc(a)
    if a.terminal_growth >= discount:
        raise ValueError("terminal growth must be below WACC or the value is infinite")
    return rows[-1].fcff * (1.0 + a.terminal_growth) / (discount - a.terminal_growth)


def terminal_value_exit_multiple(a: Assumptions, rows: list[Projection]) -> float:
    final = rows[-1]
    ebitda = final.ebit + final.depreciation
    return ebitda * a.exit_ebitda_multiple


def discount_terminal(a: Assumptions, terminal: float, rows: list[Projection]) -> float:
    """Terminal value is discounted a full N years, never N - 0.5.

    The mid-year convention applies to flows arriving *through* each year. The
    terminal value is a lump sum at the end of the horizon, and discounting it
    half a year less overstates it by about 6% at a 9% WACC.
    """
    return terminal / (1.0 + wacc(a)) ** len(rows)


@dataclass
class Valuation:
    method: str
    pv_explicit: float
    terminal_value: float
    pv_terminal: float
    enterprise_value: float
    equity_value: float
    value_per_share: float
    terminal_share: float
    implied_exit_multiple: float
    implied_perpetuity_growth: float


def value(a: Assumptions, method: str) -> Valuation:
    rows = project(a)
    pv_explicit = sum(r.present_value for r in rows)

    terminal = (terminal_value_gordon(a, rows) if method == "gordon"
                else terminal_value_exit_multiple(a, rows))
    pv_terminal = discount_terminal(a, terminal, rows)

    ev = pv_explicit + pv_terminal
    equity = ev - a.net_debt
    per_share = equity / a.diluted_shares

    final = rows[-1]
    final_ebitda = final.ebit + final.depreciation
    implied_multiple = terminal / final_ebitda

    # The cross-check every DCF should carry: what perpetuity growth rate is
    # the exit multiple actually assuming? If it is above nominal GDP the
    # multiple is doing the work, not the analysis.
    discount = wacc(a)
    implied_growth = (terminal * discount - final.fcff) / (terminal + final.fcff)

    return Valuation(
        method=method,
        pv_explicit=pv_explicit,
        terminal_value=terminal,
        pv_terminal=pv_terminal,
        enterprise_value=ev,
        equity_value=equity,
        value_per_share=per_share,
        terminal_share=pv_terminal / ev,
        implied_exit_multiple=implied_multiple,
        implied_perpetuity_growth=implied_growth,
    )


def sensitivity(a: Assumptions, wacc_deltas, growth_deltas) -> list[tuple[float, float, float]]:
    """Value per share across a WACC and terminal-growth grid.

    WACC is perturbed through the cost of equity so the grid stays internally
    consistent: bumping a stored WACC directly would leave the beta, the
    capital structure and the cost of debt all disagreeing with it.
    """
    base_wacc = wacc(a)
    grid = []
    for dw in wacc_deltas:
        target = base_wacc + dw
        # Solve for the levered beta that produces the target WACC.
        equity_weight = a.market_cap / (a.market_cap + a.total_debt)
        debt_leg = (a.total_debt / (a.market_cap + a.total_debt)) * a.pretax_cost_of_debt * (1 - a.tax_rate)
        required_coe = (target - debt_leg) / equity_weight
        beta = (required_coe - a.risk_free_rate) / a.equity_risk_premium

        for dg in growth_deltas:
            perturbed = Assumptions(
                **{**a.__dict__, "levered_beta": beta, "terminal_growth": a.terminal_growth + dg}
            )
            grid.append((target, perturbed.terminal_growth, value(perturbed, "gordon").value_per_share))
    return grid


def main() -> None:
    a = Assumptions()
    rows: list[dict[str, object]] = []

    def record(name: str, value_: float, tolerance: float = 1e-6) -> None:
        rows.append({"name": name, "expected": f"{value_:.10f}", "tolerance": tolerance})

    record("CostOfEquity", cost_of_equity(a), 1e-9)
    record("WACC", wacc(a), 1e-9)
    record("NetDebt", a.net_debt, 1e-6)

    debt_to_equity = a.total_debt / a.market_cap
    unlevered = unlever_beta(a.levered_beta, debt_to_equity, a.tax_rate)
    record("UnleveredBeta", unlevered, 1e-9)
    record("ReleveredBeta", relever_beta(unlevered, debt_to_equity, a.tax_rate), 1e-9)

    projections = project(a)
    for r in projections:
        record(f"Revenue_Y{r.year}", r.revenue, 1e-5)
        record(f"FCFF_Y{r.year}", r.fcff, 1e-5)
        record(f"DiscountFactor_Y{r.year}", r.discount_factor, 1e-9)
        record(f"PV_Y{r.year}", r.present_value, 1e-5)

    for method in ("gordon", "exit"):
        v = value(a, method)
        tag = "Gordon" if method == "gordon" else "Exit"
        record(f"{tag}_PVExplicit", v.pv_explicit, 1e-5)
        record(f"{tag}_TerminalValue", v.terminal_value, 1e-5)
        record(f"{tag}_PVTerminal", v.pv_terminal, 1e-5)
        record(f"{tag}_EnterpriseValue", v.enterprise_value, 1e-5)
        record(f"{tag}_EquityValue", v.equity_value, 1e-5)
        record(f"{tag}_ValuePerShare", v.value_per_share, 1e-7)
        record(f"{tag}_TerminalShare", v.terminal_share, 1e-9)
        record(f"{tag}_ImpliedExitMultiple", v.implied_exit_multiple, 1e-8)
        record(f"{tag}_ImpliedPerpetuityGrowth", v.implied_perpetuity_growth, 1e-9)

    # A handful of sensitivity cells, so the grid logic is pinned too.
    for dw, dg, per_share in sensitivity(a, (-0.01, 0.0, 0.01), (-0.005, 0.0, 0.005)):
        record(f"Sensitivity_W{dw * 10000:.0f}bp_G{dg * 10000:.0f}bp", per_share, 1e-6)

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with OUT.open("w", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=["name", "expected", "tolerance"])
        writer.writeheader()
        writer.writerows(rows)

    print(f"wrote {len(rows)} golden values to {OUT}")
    print()
    for method in ("gordon", "exit"):
        v = value(a, method)
        print(f"{method:8s} EV {v.enterprise_value:10,.1f}  equity {v.equity_value:10,.1f}  "
              f"per share {v.value_per_share:7.2f}  terminal {v.terminal_share:6.1%}  "
              f"implied g {v.implied_perpetuity_growth:6.2%}")
    print(f"\nWACC {wacc(a):.4%}   cost of equity {cost_of_equity(a):.4%}   "
          f"market EV {a.enterprise_value_market:,.0f}")


if __name__ == "__main__":
    main()
