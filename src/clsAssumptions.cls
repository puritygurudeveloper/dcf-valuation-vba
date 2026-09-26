VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAssumptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' clsAssumptions
'
' Every input the valuation depends on, in one object.
'
' A class rather than a sheet of named ranges because the model has to be run
' hundreds of times for a sensitivity grid, and each run needs its own isolated
' copy of the inputs. Perturbing cells on a worksheet and reading the answer
' back leaves the grid dependent on recalculation order, which is how a
' sensitivity table ends up showing the base case in every cell.
'
' Validation happens in Validate, called by the engine before any arithmetic.
' An invalid assumption set should fail with a sentence an analyst can act on,
' not produce #NUM! four columns to the right.
'==============================================================================

Private Const ERR_BASE As Long = vbObjectError + 5000

Private mRevenue0 As Double
Private mRevenueGrowth() As Double
Private mEbitMargin() As Double
Private mTaxRate As Double
Private mDepreciationPctRevenue As Double
Private mCapexPctRevenue As Double
Private mNwcPctRevenue As Double

Private mRiskFreeRate As Double
Private mEquityRiskPremium As Double
Private mLeveredBeta As Double
Private mPretaxCostOfDebt As Double
Private mMarketCap As Double
Private mTotalDebt As Double
Private mCash As Double

Private mTerminalGrowth As Double
Private mExitEbitdaMultiple As Double
Private mDilutedShares As Double
Private mMidYearConvention As Boolean

'--- Construction -------------------------------------------------------------

Private Sub Class_Initialize()
    ' Defaults are the documented base case; see docs/model_spec.md.
    mRevenue0 = 4820#
    mRevenueGrowth = Array(0.082, 0.071, 0.06, 0.05, 0.042)
    mEbitMargin = Array(0.184, 0.189, 0.193, 0.196, 0.198)
    mTaxRate = 0.24
    mDepreciationPctRevenue = 0.041
    mCapexPctRevenue = 0.052
    mNwcPctRevenue = 0.113

    mRiskFreeRate = 0.0428
    mEquityRiskPremium = 0.051
    mLeveredBeta = 1.18
    mPretaxCostOfDebt = 0.0565
    mMarketCap = 11430#
    mTotalDebt = 3260#
    mCash = 742#

    mTerminalGrowth = 0.024
    mExitEbitdaMultiple = 11.5
    mDilutedShares = 214.6
    mMidYearConvention = True
End Sub

' Deep copy. The sensitivity grid mutates a copy per cell, and VBA's Set is a
' reference assignment, so without this every cell would share one object.
Public Function Clone() As clsAssumptions
    Dim result As clsAssumptions
    Set result = New clsAssumptions

    result.Revenue0 = mRevenue0
    result.SetRevenueGrowth mRevenueGrowth
    result.SetEbitMargin mEbitMargin
    result.TaxRate = mTaxRate
    result.DepreciationPctRevenue = mDepreciationPctRevenue
    result.CapexPctRevenue = mCapexPctRevenue
    result.NwcPctRevenue = mNwcPctRevenue

    result.RiskFreeRate = mRiskFreeRate
    result.EquityRiskPremium = mEquityRiskPremium
    result.LeveredBeta = mLeveredBeta
    result.PretaxCostOfDebt = mPretaxCostOfDebt
    result.MarketCap = mMarketCap
    result.TotalDebt = mTotalDebt
    result.Cash = mCash

    result.TerminalGrowth = mTerminalGrowth
    result.ExitEbitdaMultiple = mExitEbitdaMultiple
    result.DilutedShares = mDilutedShares
    result.MidYearConvention = mMidYearConvention

    Set Clone = result
End Function

'--- Properties ---------------------------------------------------------------

Public Property Get Revenue0() As Double: Revenue0 = mRevenue0: End Property
Public Property Let Revenue0(value As Double): mRevenue0 = value: End Property

Public Property Get TaxRate() As Double: TaxRate = mTaxRate: End Property
Public Property Let TaxRate(value As Double): mTaxRate = value: End Property

Public Property Get DepreciationPctRevenue() As Double: DepreciationPctRevenue = mDepreciationPctRevenue: End Property
Public Property Let DepreciationPctRevenue(value As Double): mDepreciationPctRevenue = value: End Property

Public Property Get CapexPctRevenue() As Double: CapexPctRevenue = mCapexPctRevenue: End Property
Public Property Let CapexPctRevenue(value As Double): mCapexPctRevenue = value: End Property

Public Property Get NwcPctRevenue() As Double: NwcPctRevenue = mNwcPctRevenue: End Property
Public Property Let NwcPctRevenue(value As Double): mNwcPctRevenue = value: End Property

Public Property Get RiskFreeRate() As Double: RiskFreeRate = mRiskFreeRate: End Property
Public Property Let RiskFreeRate(value As Double): mRiskFreeRate = value: End Property

Public Property Get EquityRiskPremium() As Double: EquityRiskPremium = mEquityRiskPremium: End Property
Public Property Let EquityRiskPremium(value As Double): mEquityRiskPremium = value: End Property

Public Property Get LeveredBeta() As Double: LeveredBeta = mLeveredBeta: End Property
Public Property Let LeveredBeta(value As Double): mLeveredBeta = value: End Property

Public Property Get PretaxCostOfDebt() As Double: PretaxCostOfDebt = mPretaxCostOfDebt: End Property
Public Property Let PretaxCostOfDebt(value As Double): mPretaxCostOfDebt = value: End Property

Public Property Get MarketCap() As Double: MarketCap = mMarketCap: End Property
Public Property Let MarketCap(value As Double): mMarketCap = value: End Property

Public Property Get TotalDebt() As Double: TotalDebt = mTotalDebt: End Property
Public Property Let TotalDebt(value As Double): mTotalDebt = value: End Property

Public Property Get Cash() As Double: Cash = mCash: End Property
Public Property Let Cash(value As Double): mCash = value: End Property

Public Property Get TerminalGrowth() As Double: TerminalGrowth = mTerminalGrowth: End Property
Public Property Let TerminalGrowth(value As Double): mTerminalGrowth = value: End Property

Public Property Get ExitEbitdaMultiple() As Double: ExitEbitdaMultiple = mExitEbitdaMultiple: End Property
Public Property Let ExitEbitdaMultiple(value As Double): mExitEbitdaMultiple = value: End Property

Public Property Get DilutedShares() As Double: DilutedShares = mDilutedShares: End Property
Public Property Let DilutedShares(value As Double): mDilutedShares = value: End Property

Public Property Get MidYearConvention() As Boolean: MidYearConvention = mMidYearConvention: End Property
Public Property Let MidYearConvention(value As Boolean): mMidYearConvention = value: End Property

'--- Array-valued inputs ------------------------------------------------------

Public Property Get RevenueGrowth() As Variant: RevenueGrowth = mRevenueGrowth: End Property
Public Sub SetRevenueGrowth(values As Variant): mRevenueGrowth = values: End Sub

Public Property Get EbitMargin() As Variant: EbitMargin = mEbitMargin: End Property
Public Sub SetEbitMargin(values As Variant): mEbitMargin = values: End Sub

Public Property Get ForecastYears() As Long
    ForecastYears = UBound(mRevenueGrowth) - LBound(mRevenueGrowth) + 1
End Property

'--- Derived ------------------------------------------------------------------

Public Property Get NetDebt() As Double
    NetDebt = mTotalDebt - mCash
End Property

Public Property Get DebtToEquity() As Double
    ' Market values throughout. Book equity is an accounting residual and has
    ' nothing to do with what investors require.
    If mMarketCap <= 0 Then Err.Raise ERR_BASE + 1, "clsAssumptions", "MarketCap must be positive"
    DebtToEquity = mTotalDebt / mMarketCap
End Property

Public Property Get MarketEnterpriseValue() As Double
    MarketEnterpriseValue = mMarketCap + NetDebt
End Property

'--- Validation ---------------------------------------------------------------

Public Sub Validate()
    If mRevenue0 <= 0 Then
        Err.Raise ERR_BASE + 2, "clsAssumptions", "Base revenue must be positive"
    End If
    If mTaxRate < 0 Or mTaxRate >= 1 Then
        Err.Raise ERR_BASE + 3, "clsAssumptions", "Tax rate must lie in [0, 1)"
    End If
    If mDilutedShares <= 0 Then
        Err.Raise ERR_BASE + 4, "clsAssumptions", "Diluted share count must be positive"
    End If
    If UBound(mRevenueGrowth) <> UBound(mEbitMargin) Then
        Err.Raise ERR_BASE + 5, "clsAssumptions", _
            "Revenue growth and EBIT margin must cover the same number of years"
    End If
    If mMarketCap + mTotalDebt <= 0 Then
        Err.Raise ERR_BASE + 6, "clsAssumptions", "Capital structure cannot be empty"
    End If
End Sub
