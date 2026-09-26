Attribute VB_Name = "modWACC"
Option Explicit

'==============================================================================
' modWACC
'
' Cost of capital. Small module, and the place most valuation errors live.
'==============================================================================

Private Const ERR_BASE As Long = vbObjectError + 5100

' CAPM cost of equity.
Public Function CostOfEquity(a As clsAssumptions) As Double
    CostOfEquity = a.RiskFreeRate + a.LeveredBeta * a.EquityRiskPremium
End Function

' Weighted average cost of capital, on market values.
'
' The debt leg is after tax because interest is deductible; the equity leg is
' not, because dividends are paid out of taxed profit. Applying the tax shield
' to both -- or to neither -- is the classic WACC error and moves the answer by
' more than any amount of argument about beta.
Public Function WACC(a As clsAssumptions) As Double
    Dim totalCapital As Double
    totalCapital = a.MarketCap + a.TotalDebt
    If totalCapital <= 0 Then
        Err.Raise ERR_BASE + 1, "modWACC", "Capital structure cannot be empty"
    End If

    WACC = (a.MarketCap / totalCapital) * CostOfEquity(a) _
         + (a.TotalDebt / totalCapital) * a.PretaxCostOfDebt * (1# - a.TaxRate)
End Function

' Hamada unlevering. Strips financing risk out of an observed equity beta so
' that betas from differently geared comparables can be averaged.
Public Function UnleverBeta(leveredBeta As Double, debtToEquity As Double, _
                            taxRate As Double) As Double
    UnleverBeta = leveredBeta / (1# + (1# - taxRate) * debtToEquity)
End Function

Public Function ReleverBeta(unleveredBeta As Double, debtToEquity As Double, _
                            taxRate As Double) As Double
    ReleverBeta = unleveredBeta * (1# + (1# - taxRate) * debtToEquity)
End Function

' The levered beta that produces a target WACC, holding everything else fixed.
'
' Used by the sensitivity grid. Perturbing a stored WACC directly would leave
' the beta, the capital structure and the cost of debt all disagreeing with it,
' so the grid would be internally inconsistent -- every cell a different,
' unstated company.
Public Function BetaForTargetWACC(a As clsAssumptions, targetWACC As Double) As Double
    Dim totalCapital As Double, equityWeight As Double, debtLeg As Double
    Dim requiredCostOfEquity As Double

    totalCapital = a.MarketCap + a.TotalDebt
    equityWeight = a.MarketCap / totalCapital
    debtLeg = (a.TotalDebt / totalCapital) * a.PretaxCostOfDebt * (1# - a.TaxRate)

    requiredCostOfEquity = (targetWACC - debtLeg) / equityWeight
    If a.EquityRiskPremium <= 0 Then
        Err.Raise ERR_BASE + 2, "modWACC", "Equity risk premium must be positive"
    End If
    BetaForTargetWACC = (requiredCostOfEquity - a.RiskFreeRate) / a.EquityRiskPremium
End Function
