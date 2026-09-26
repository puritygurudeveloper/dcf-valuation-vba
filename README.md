# DCF valuation model — VBA

A discounted cash flow and comparable-multiples valuation engine written in
VBA, with a test suite, a specification, and a build that keeps the macros in
version control as plain text.

```
                       Gordon growth    Exit multiple
Enterprise value         $11,910.8m       $14,627.0m
Equity value              $9,392.8m       $12,109.0m
Value per share              $43.77           $56.43
Terminal share of EV          74.5%            79.2%
Implied perpetuity growth      2.40%            3.87%

WACC 8.97%   cost of equity 10.30%   market enterprise value $13,948m
```

The two methods differ by **29%**, and the reason is in the last row. The 11.5x
exit multiple implies a perpetuity growth rate of 3.87% — above long-run
nominal GDP. The multiple is doing the work, not the analysis. A DCF that
reports only the headline number hides that; this one computes it on every run.

## Why VBA has a test suite

VBA has no test framework, no package manager and no CI runner, so financial
models written in it are typically verified by someone re-deriving a number by
hand once and never again. That is the actual risk in a valuation model: not
that the theory is wrong, but that a `ΔNWC` became an `NWC` in year three and
nobody noticed for a year.

The approach here:

1. `tools/generate_golden.py` implements the specification **independently**,
   in a language that runs in a pipeline.
2. It writes `tests/golden_values.csv` — 52 expected values with per-value
   tolerances, covering every intermediate as well as the headline.
3. `modTests.RunAllTests` reads that file and checks the VBA against it.

A VBA test asserting a VBA formula against another VBA formula proves only that
the two agree. Checking against a separately written reference catches the case
where the formula itself is wrong.

CI regenerates the golden values on every push and fails if they differ from
what is committed, so a VBA test failure always means the VBA is wrong rather
than the expectations being stale.

## The errors this model is built to avoid

**Working capital as a level, not a change.** `FCFF = NOPAT + D&A − capex −
ΔNWC`. Subtracting the *level* understates cash flow every year by roughly
revenue × the NWC ratio — about $545m a year here, which is more than the
entire year-one free cash flow. It is invisible unless somebody reconciles the
first forecast year to the last actual one.

**Half-year discounting the terminal value.** The mid-year convention applies to
flows arriving *through* each year. The terminal value is a lump sum at the
horizon. Discounting it `N − 0.5` years overstates it by about 4%, and since it
is three quarters of this valuation, roughly 3% of the answer.

**The tax shield on the wrong leg.** Interest is deductible, dividends are not.
The after-tax adjustment belongs on the debt leg of the WACC only.

**Book equity in the WACC.** Book equity is an accounting residual. Weights are
market values throughout.

**A sensitivity table that shows the base case in every cell.** Excel's native
data table recalculates the whole workbook per cell and depends on calculation
mode; leave it on manual and every cell is identical. `modSensitivity.BuildGrid`
constructs the grid in code from a cloned assumption object — deterministic,
testable, and a test explicitly asserts that no two corner cells are equal.

**A sensitivity grid that is internally inconsistent.** Bumping a stored WACC
directly leaves the beta, the capital structure and the cost of debt all
disagreeing with it — every cell is then a different, unstated company. WACC is
moved by solving for the levered beta that produces the target
(`modWACC.BetaForTargetWACC`), so the whole grid stays coherent.

## Two things VBA forces you to get right

**`Clone` must be a deep copy.** VBA's `Set` is a reference assignment, so
without an explicit clone every cell of the sensitivity grid would mutate and
read the same object, and the grid would collapse to whichever scenario ran
last. There is a test asserting the base case is unchanged after a scenario has
been perturbed.

**You cannot write a generic "assert this raises" helper.** VBA evaluates
arguments at the call site, so `CheckRaises "x", SomethingThatThrows` runs the
probe *first* and the error escapes past the assertion, aborting the run. The
first version of `modTests` did exactly that. Each negative case now traps
inline with `On Error Resume Next` / `Err.Number`. Verbose, and the only
construction that works.

## Layout

```
src/clsAssumptions.cls    inputs, validation, deep clone
src/modWACC.bas           CAPM, WACC, Hamada un/relevering, inverse solve
src/modDCF.bas            FCFF projection, terminal value, EV-to-share bridge
src/modSensitivity.bas    the WACC x growth grid
src/modTests.bas          test runner: golden values, invariants, validation

tools/generate_golden.py  the independent reference implementation
tools/build_workbook.py   builds the .xlsx shell with named ranges
tests/golden_values.csv   52 expected values with tolerances
docs/model_spec.md        the arithmetic, written down once
```

Macros live in `src/` as plain `.bas` and `.cls` text. A `.xlsm` is a binary —
its diff is one unreadable line — so the workbook is built from
`tools/build_workbook.py` and the modules are imported into it. Named ranges
rather than cell addresses, so inserting a row cannot silently repoint an
assumption at the wrong cell.

## Running it

```bash
python tools/generate_golden.py     # regenerate expected values
python tools/build_workbook.py      # build/dcf_model.xlsx
```

Then save the workbook as `.xlsm`, import `src/*.bas` and `src/*.cls` via the
VBA editor, and run `RunAllTests` from the Immediate window (Ctrl+G):

```
  ok   WACC
  ok   Gordon_ValuePerShare
  ...
  ok   mid-year convention raises the value
  ok   Clone does not alias the original
  ok   terminal growth above WACC is rejected
----------------------------------------------------------------------
passed: 71   failed: 0
ALL TESTS PASSED
```

## Scope

Single-scenario FCFF DCF with two terminal methods. No three-statement model —
the projection is driver-based, so the balance sheet does not balance and there
is no circularity between interest expense and the cash sweep, which is the
usual reason these models need iterative calculation enabled. No operating
lease adjustment, no NOL carryforwards, no minority interest or associates in
the bridge, and no currency translation.

The largest unmodelled risk is the one the output already flags: 75% of the
value sits beyond year five, so this is substantially a terminal value
assumption with a five-year preamble attached. That is true of every DCF, and
the reason the implied-growth cross-check is printed rather than optional.
