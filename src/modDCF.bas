Attribute VB_Name = "modDCF"
Option Explicit

'==============================================================================
' modDCF
'
' Unlevered free cash flow projection, terminal value, and the bridge from
' enterprise value to value per share.
'==============================================================================

Private Const ERR_BASE As Long = vbObjectError + 5200

Public Enum TerminalMethod
    tmGordonGrowth = 1
    tmExitMultiple = 2
End Enum

' One projected year.
Public Type ProjectionYear
    Year As Long
    Revenue As Double
    Ebit As Double
    Nopat As Double
    Depreciation As Double
    Capex As Double
    ChangeInNwc As Double
    Fcff As Double
    DiscountFactor As Double
    PresentValue As Double
End Type

Public Type ValuationResult
    Method As TerminalMethod
    PvExplicit As Double
    TerminalValue As Double
    PvTerminal As Double
    EnterpriseValue As Double
    EquityValue As Double
    ValuePerShare As Double
    TerminalShare As Double
    ImpliedExitMultiple As Double
    ImpliedPerpetuityGrowth As Double
End Type

'------------------------------------------------------------------------------
' Project unlevered free cash flow.
'
'   FCFF = EBIT x (1 - t) + D&A - capex - change in net working capital
'
' The working capital term is a CHANGE, not a level. Subtracting the level is a
' mistake that understates cash flow every single year by roughly revenue times
' the working capital ratio -- about $545m a year on the base case -- and it is
' invisible unless somebody reconciles the first forecast year to the last
' actual one.
'------------------------------------------------------------------------------
Public Function ProjectCashFlows(a As clsAssumptions) As ProjectionYear()
    a.Validate

    Dim growth As Variant, margin As Variant
    growth = a.RevenueGrowth
    margin = a.EbitMargin

    Dim n As Long
    n = a.ForecastYears

    Dim rows() As ProjectionYear
    ReDim rows(1 To n)

    Dim discountRate As Double
    discountRate = modWACC.WACC(a)

    Dim revenue As Double, priorNwc As Double, newNwc As Double
    Dim exponent As Double
    Dim i As Long

    revenue = a.Revenue0
    priorNwc = a.NwcPctRevenue * revenue

    For i = 1 To n
        revenue = revenue * (1# + CDbl(growth(LBound(growth) + i - 1)))

        With rows(i)
            .Year = i
            .Revenue = revenue
            .Ebit = revenue * CDbl(margin(LBound(margin) + i - 1))
            .Nopat = .Ebit * (1# - a.TaxRate)
            .Depreciation = revenue * a.DepreciationPctRevenue
            .Capex = revenue * a.CapexPctRevenue

            newNwc = revenue * a.NwcPctRevenue
            .ChangeInNwc = newNwc - priorNwc
            priorNwc = newNwc

            .Fcff = .Nopat + .Depreciation - .Capex - .ChangeInNwc

            ' Mid-year convention: cash arrives through the year rather than on
            ' 31 December, so the average flow is discounted half a year less.
            If a.MidYearConvention Then
                exponent = CDbl(i) - 0.5
            Else
                exponent = CDbl(i)
            End If
            .DiscountFactor = 1# / ((1# + discountRate) ^ exponent)
            .PresentValue = .Fcff * .DiscountFactor
        End With
    Next i

    ProjectCashFlows = rows
End Function

' Gordon growth terminal value on the final-year cash flow.
Public Function TerminalValueGordon(a As clsAssumptions, rows() As ProjectionYear) As Double
    Dim discountRate As Double
    discountRate = modWACC.WACC(a)

    If a.TerminalGrowth >= discountRate Then
        Err.Raise ERR_BASE + 1, "modDCF", _
            "Terminal growth (" & Format$(a.TerminalGrowth, "0.00%") & ") must be below WACC (" & _
            Format$(discountRate, "0.00%") & ") or the value is infinite"
    End If

    TerminalValueGordon = rows(UBound(rows)).Fcff * (1# + a.TerminalGrowth) _
                        / (discountRate - a.TerminalGrowth)
End Function

Public Function TerminalValueExitMultiple(a As clsAssumptions, rows() As ProjectionYear) As Double
    Dim final As ProjectionYear
    final = rows(UBound(rows))
    TerminalValueExitMultiple = (final.Ebit + final.Depreciation) * a.ExitEbitdaMultiple
End Function

'------------------------------------------------------------------------------
' Discount the terminal value a FULL N years, never N - 0.5.
'
' The mid-year convention applies to flows arriving through each year. The
' terminal value is a lump sum representing everything after the horizon, and
' half-year discounting it overstates the answer by about 4% at a 9% WACC --
' which, since the terminal value is three quarters of this valuation, is 3% of
' the whole number.
'------------------------------------------------------------------------------
Public Function DiscountTerminalValue(a As clsAssumptions, terminalValue As Double, _
                                      horizonYears As Long) As Double
    DiscountTerminalValue = terminalValue / ((1# + modWACC.WACC(a)) ^ CDbl(horizonYears))
End Function

Public Function Value(a As clsAssumptions, method As TerminalMethod) As ValuationResult
    Dim rows() As ProjectionYear
    rows = ProjectCashFlows(a)

    Dim result As ValuationResult
    result.Method = method

    Dim i As Long
    For i = LBound(rows) To UBound(rows)
        result.PvExplicit = result.PvExplicit + rows(i).PresentValue
    Next i

    Select Case method
        Case tmGordonGrowth
            result.TerminalValue = TerminalValueGordon(a, rows)
        Case tmExitMultiple
            result.TerminalValue = TerminalValueExitMultiple(a, rows)
        Case Else
            Err.Raise ERR_BASE + 2, "modDCF", "Unknown terminal method"
    End Select

    result.PvTerminal = DiscountTerminalValue(a, result.TerminalValue, UBound(rows))
    result.EnterpriseValue = result.PvExplicit + result.PvTerminal
    result.EquityValue = result.EnterpriseValue - a.NetDebt
    result.ValuePerShare = result.EquityValue / a.DilutedShares
    result.TerminalShare = result.PvTerminal / result.EnterpriseValue

    Dim final As ProjectionYear
    final = rows(UBound(rows))
    result.ImpliedExitMultiple = result.TerminalValue / (final.Ebit + final.Depreciation)

    ' The cross-check every DCF should carry. Rearranging the Gordon formula
    ' for g answers: what perpetuity growth rate is this terminal value
    ' actually assuming? If an exit multiple implies growth above nominal GDP,
    ' the multiple is doing the work and the analysis is decoration.
    result.ImpliedPerpetuityGrowth = _
        (result.TerminalValue * modWACC.WACC(a) - final.Fcff) / (result.TerminalValue + final.Fcff)

    Value = result
End Function
