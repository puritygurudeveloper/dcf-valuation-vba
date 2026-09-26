# Model specification

The single source of truth for the arithmetic. `tools/generate_golden.py` and
`src/modDCF.bas` are two independent implementations of what follows; the test
runner checks they agree.

## Free cash flow to the firm

For each forecast year *t*:

```
revenue_t        = revenue_{t-1} × (1 + g_t)
EBIT_t           = revenue_t × margin_t
NOPAT_t          = EBIT_t × (1 − tax rate)
D&A_t            = revenue_t × depreciation %
capex_t          = revenue_t × capex %
NWC_t            = revenue_t × NWC %
ΔNWC_t           = NWC_t − NWC_{t−1}
FCFF_t           = NOPAT_t + D&A_t − capex_t − ΔNWC_t
```

`ΔNWC` is a **change**, not a level. Year 0 working capital is
`revenue_0 × NWC %`, so the first forecast year's change is measured against
the last actual year rather than against zero.

## Discounting

```
exponent_t       = t − 0.5   (mid-year convention)  or  t
discount_t       = (1 + WACC) ^ −exponent_t
PV_t             = FCFF_t × discount_t
```

The terminal value is discounted `(1 + WACC) ^ −N`, a **full** N years. It is a
lump sum at the horizon, not a flow arriving through the final year.

## Cost of capital

```
cost of equity   = r_f + β_L × ERP                          (CAPM)
WACC             = (E/(E+D)) × cost of equity
                 + (D/(E+D)) × cost of debt × (1 − tax rate)
```

`E` is market capitalisation, not book equity. Only the debt leg carries the
tax shield.

Hamada:

```
β_U              = β_L / (1 + (1 − t) × D/E)
β_L              = β_U × (1 + (1 − t) × D/E)
```

## Terminal value

```
Gordon:          TV = FCFF_N × (1 + g) / (WACC − g),  requires g < WACC
Exit multiple:   TV = (EBIT_N + D&A_N) × multiple
```

Cross-check, reported for both methods — the perpetuity growth rate the
terminal value implies:

```
g_implied        = (TV × WACC − FCFF_N) / (TV + FCFF_N)
```

## Bridge

```
enterprise value = Σ PV_t + PV(TV)
net debt         = total debt − cash
equity value     = enterprise value − net debt
value per share  = equity value / diluted shares
```

## Base case

| input | value | | input | value |
|---|---|---|---|---|
| Base revenue | $4,820m | | Risk-free rate | 4.28% |
| Revenue growth | 8.2 / 7.1 / 6.0 / 5.0 / 4.2% | | Equity risk premium | 5.10% |
| EBIT margin | 18.4 / 18.9 / 19.3 / 19.6 / 19.8% | | Levered beta | 1.18 |
| Tax rate | 24.0% | | Pre-tax cost of debt | 5.65% |
| D&A % revenue | 4.1% | | Market cap | $11,430m |
| Capex % revenue | 5.2% | | Total debt | $3,260m |
| NWC % revenue | 11.3% | | Cash | $742m |
| Terminal growth | 2.4% | | Exit EV/EBITDA | 11.5x |
| Diluted shares | 214.6m | | Mid-year convention | on |

## Results

| | Gordon growth | Exit multiple |
|---|---|---|
| Enterprise value | $11,910.8m | $14,627.0m |
| Equity value | $9,392.8m | $12,109.0m |
| **Value per share** | **$43.77** | **$56.43** |
| Terminal share of EV | 74.5% | 79.2% |
| Implied perpetuity growth | 2.40% | 3.87% |

WACC 8.97%, cost of equity 10.30%, market enterprise value $13,948m.
