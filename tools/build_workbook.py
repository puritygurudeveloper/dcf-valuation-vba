"""Assemble the .xlsm shell the VBA modules are imported into.

openpyxl cannot write a VBA project, so this builds the workbook -- named
ranges, input sheet, output sheet, formats -- and the modules are imported
once by hand or by the PowerShell script in tools/import_modules.ps1. The
split keeps the VBA under version control as plain text, which a .xlsm binary
would not be: a diff of a macro-enabled workbook is one unreadable line.
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "build" / "dcf_model.xlsx"

HEADER = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="1F3864")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
LABEL = Font(bold=True)

INPUTS: list[tuple[str, str, object, str]] = [
    ("Revenue0", "Base year revenue ($m)", 4820.0, "#,##0.0"),
    ("TaxRate", "Effective tax rate", 0.24, "0.0%"),
    ("DepreciationPct", "D&A as % of revenue", 0.041, "0.0%"),
    ("CapexPct", "Capex as % of revenue", 0.052, "0.0%"),
    ("NwcPct", "Net working capital as % of revenue", 0.113, "0.0%"),
    ("RiskFreeRate", "Risk-free rate", 0.0428, "0.00%"),
    ("EquityRiskPremium", "Equity risk premium", 0.0510, "0.00%"),
    ("LeveredBeta", "Levered beta", 1.18, "0.00"),
    ("PretaxCostOfDebt", "Pre-tax cost of debt", 0.0565, "0.00%"),
    ("MarketCap", "Market capitalisation ($m)", 11430.0, "#,##0.0"),
    ("TotalDebt", "Total debt ($m)", 3260.0, "#,##0.0"),
    ("Cash", "Cash and equivalents ($m)", 742.0, "#,##0.0"),
    ("TerminalGrowth", "Terminal growth rate", 0.024, "0.00%"),
    ("ExitEbitdaMultiple", "Exit EV/EBITDA multiple", 11.5, "0.0x"),
    ("DilutedShares", "Diluted shares (m)", 214.6, "#,##0.0"),
]


def main() -> None:
    wb = Workbook()

    inputs = wb.active
    inputs.title = "Inputs"
    inputs["A1"] = "Valuation inputs"
    inputs["A1"].font = Font(bold=True, size=14)
    inputs["A2"] = "Shaded cells are inputs. Everything else is computed in VBA."
    inputs["A2"].font = Font(italic=True, size=9)

    for row, (name, label, value, fmt) in enumerate(INPUTS, start=4):
        inputs.cell(row=row, column=1, value=label).font = LABEL
        cell = inputs.cell(row=row, column=2, value=value)
        cell.number_format = fmt
        cell.fill = INPUT_FILL
        # A defined name per input, so the VBA reads Range("TaxRate") rather
        # than Range("B5") -- inserting a row then cannot silently repoint an
        # assumption at the wrong cell.
        wb.defined_names.add(_defined_name(name, f"Inputs!$B${row}"))

    inputs.column_dimensions["A"].width = 38
    inputs.column_dimensions["B"].width = 16

    # Forecast drivers, one column per year.
    drivers = wb.create_sheet("Drivers")
    drivers["A1"] = "Forecast drivers"
    drivers["A1"].font = Font(bold=True, size=14)
    for col, year in enumerate(range(1, 6), start=2):
        header = drivers.cell(row=3, column=col, value=f"Year {year}")
        header.font = HEADER
        header.fill = HEADER_FILL
        header.alignment = Alignment(horizontal="center")

    for row, (label, values) in enumerate(
        [
            ("Revenue growth", (0.082, 0.071, 0.060, 0.050, 0.042)),
            ("EBIT margin", (0.184, 0.189, 0.193, 0.196, 0.198)),
        ],
        start=4,
    ):
        drivers.cell(row=row, column=1, value=label).font = LABEL
        for col, value in enumerate(values, start=2):
            cell = drivers.cell(row=row, column=col, value=value)
            cell.number_format = "0.0%"
            cell.fill = INPUT_FILL

    wb.defined_names.add(_defined_name("RevenueGrowth", "Drivers!$B$4:$F$4"))
    wb.defined_names.add(_defined_name("EbitMargin", "Drivers!$B$5:$F$5"))
    drivers.column_dimensions["A"].width = 24

    output = wb.create_sheet("Valuation")
    output["A1"] = "Valuation output"
    output["A1"].font = Font(bold=True, size=14)
    output["A3"] = "Run modDCF.Value or the RunModel macro to populate."
    output["A3"].font = Font(italic=True, size=9)
    output.column_dimensions["A"].width = 34
    for col in range(2, 8):
        output.column_dimensions[get_column_letter(col)].width = 14

    sensitivity = wb.create_sheet("Sensitivity")
    sensitivity["A1"] = "Value per share: WACC vs terminal growth"
    sensitivity["A1"].font = Font(bold=True, size=14)
    sensitivity["A3"] = "Built by modSensitivity.BuildGrid, not an Excel data table."
    sensitivity["A3"].font = Font(italic=True, size=9)
    sensitivity.column_dimensions["A"].width = 14

    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print(f"wrote {OUT}")
    print("next: open it, save as .xlsm, then import src/*.bas and src/*.cls")


def _defined_name(name: str, ref: str):
    from openpyxl.workbook.defined_name import DefinedName

    return DefinedName(name, attr_text=ref)


if __name__ == "__main__":
    main()
