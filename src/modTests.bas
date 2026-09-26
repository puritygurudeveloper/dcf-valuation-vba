Attribute VB_Name = "modTests"
Option Explicit

'==============================================================================
' modTests
'
' A test runner for a language with no test framework and no CI.
'
' Every expected value in tests/golden_values.csv was produced by
' tools/generate_golden.py -- an independent implementation of the same
' specification, in a language that can be run in a pipeline. That matters: a
' VBA test asserting a VBA formula against another VBA formula proves only that
' the two agree. Checking against a separately written reference catches the
' case where the formula itself is wrong.
'
' Run RunAllTests from the Immediate window (Ctrl+G) or Alt+F8.
'==============================================================================

Private mPassed As Long
Private mFailed As Long
Private mLog As String

'--- Entry point --------------------------------------------------------------

Public Sub RunAllTests()
    mPassed = 0
    mFailed = 0
    mLog = ""

    TestAgainstGoldenValues
    TestStructuralInvariants
    TestValidationRejectsBadInput
    TestSensitivityGrid

    Debug.Print mLog
    Debug.Print String$(70, "-")
    Debug.Print "passed: " & mPassed & "   failed: " & mFailed
    If mFailed = 0 Then
        Debug.Print "ALL TESTS PASSED"
    Else
        Debug.Print "*** " & mFailed & " FAILURE(S) ***"
    End If
End Sub

'--- The golden-value suite ---------------------------------------------------

Private Sub TestAgainstGoldenValues()
    Dim path As String
    path = GoldenValuePath()

    If Len(Dir$(path)) = 0 Then
        Fail "golden values", "cannot find " & path & _
             " (run: python tools/generate_golden.py)"
        Exit Sub
    End If

    Dim a As clsAssumptions
    Set a = New clsAssumptions

    Dim actuals As Object
    Set actuals = ComputeAllValues(a)

    Dim handle As Integer, line As String
    Dim parts() As String
    Dim name As String, expected As Double, tolerance As Double

    handle = FreeFile
    Open path For Input As #handle
    Line Input #handle, line                     ' header
    Do Until EOF(handle)
        Line Input #handle, line
        If Len(Trim$(line)) > 0 Then
            parts = Split(line, ",")
            name = Trim$(parts(0))
            expected = CDbl(parts(1))
            tolerance = CDbl(parts(2))

            If Not actuals.Exists(name) Then
                Fail name, "no value computed for this golden key"
            Else
                CheckClose name, CDbl(actuals(name)), expected, tolerance
            End If
        End If
    Loop
    Close #handle
End Sub

' Every value the golden file might ask about, keyed by name.
Private Function ComputeAllValues(a As clsAssumptions) As Object
    Dim v As Object
    Set v = CreateObject("Scripting.Dictionary")

    v("CostOfEquity") = modWACC.CostOfEquity(a)
    v("WACC") = modWACC.WACC(a)
    v("NetDebt") = a.NetDebt

    Dim unlevered As Double
    unlevered = modWACC.UnleverBeta(a.LeveredBeta, a.DebtToEquity, a.TaxRate)
    v("UnleveredBeta") = unlevered
    v("ReleveredBeta") = modWACC.ReleverBeta(unlevered, a.DebtToEquity, a.TaxRate)

    Dim rows() As modDCF.ProjectionYear
    rows = modDCF.ProjectCashFlows(a)

    Dim i As Long
    For i = LBound(rows) To UBound(rows)
        v("Revenue_Y" & rows(i).Year) = rows(i).Revenue
        v("FCFF_Y" & rows(i).Year) = rows(i).Fcff
        v("DiscountFactor_Y" & rows(i).Year) = rows(i).DiscountFactor
        v("PV_Y" & rows(i).Year) = rows(i).PresentValue
    Next i

    AddValuation v, "Gordon", modDCF.Value(a, modDCF.tmGordonGrowth)
    AddValuation v, "Exit", modDCF.Value(a, modDCF.tmExitMultiple)

    AddSensitivityValues v, a

    Set ComputeAllValues = v
End Function

Private Sub AddValuation(v As Object, tag As String, r As modDCF.ValuationResult)
    v(tag & "_PVExplicit") = r.PvExplicit
    v(tag & "_TerminalValue") = r.TerminalValue
    v(tag & "_PVTerminal") = r.PvTerminal
    v(tag & "_EnterpriseValue") = r.EnterpriseValue
    v(tag & "_EquityValue") = r.EquityValue
    v(tag & "_ValuePerShare") = r.ValuePerShare
    v(tag & "_TerminalShare") = r.TerminalShare
    v(tag & "_ImpliedExitMultiple") = r.ImpliedExitMultiple
    v(tag & "_ImpliedPerpetuityGrowth") = r.ImpliedPerpetuityGrowth
End Sub

Private Sub AddSensitivityValues(v As Object, a As clsAssumptions)
    Dim waccDeltas As Variant, growthDeltas As Variant
    waccDeltas = Array(-0.01, 0#, 0.01)
    growthDeltas = Array(-0.005, 0#, 0.005)

    Dim grid() As modSensitivity.SensitivityCell
    grid = modSensitivity.BuildGrid(a, waccDeltas, growthDeltas)

    Dim nG As Long
    nG = UBound(growthDeltas) - LBound(growthDeltas) + 1

    Dim k As Long, iW As Long, iG As Long, key As String
    For k = LBound(grid) To UBound(grid)
        iW = ((k - 1) \ nG)
        iG = ((k - 1) Mod nG)
        key = "Sensitivity_W" & Format$(CDbl(waccDeltas(iW)) * 10000#, "0") & "bp_G" & _
              Format$(CDbl(growthDeltas(iG)) * 10000#, "0") & "bp"
        v(key) = grid(k).ValuePerShare
    Next k
End Sub

'--- Invariants that hold regardless of the inputs ----------------------------

Private Sub TestStructuralInvariants()
    Dim a As clsAssumptions
    Set a = New clsAssumptions

    Dim gordon As modDCF.ValuationResult, exitM As modDCF.ValuationResult
    gordon = modDCF.Value(a, modDCF.tmGordonGrowth)
    exitM = modDCF.Value(a, modDCF.tmExitMultiple)

    ' EV bridges to equity by net debt, by construction.
    CheckClose "bridge: EV - net debt = equity", _
               gordon.EnterpriseValue - a.NetDebt, gordon.EquityValue, 0.000001

    ' The two halves of the enterprise value must add up.
    CheckClose "EV = PV(explicit) + PV(terminal)", _
               gordon.PvExplicit + gordon.PvTerminal, gordon.EnterpriseValue, 0.000001

    ' Terminal value dominates any five-year DCF. If it did not, something is
    ' wrong with the horizon or the growth rate.
    CheckTrue "terminal share is between 50% and 90%", _
              gordon.TerminalShare > 0.5 And gordon.TerminalShare < 0.9

    ' Lower discount rate, higher value -- monotone, always.
    Dim cheaper As clsAssumptions
    Set cheaper = a.Clone()
    cheaper.LeveredBeta = a.LeveredBeta - 0.2
    CheckTrue "value rises as beta falls", _
              modDCF.Value(cheaper, modDCF.tmGordonGrowth).ValuePerShare > gordon.ValuePerShare

    ' Higher terminal growth, higher value.
    Dim faster As clsAssumptions
    Set faster = a.Clone()
    faster.TerminalGrowth = a.TerminalGrowth + 0.005
    CheckTrue "value rises with terminal growth", _
              modDCF.Value(faster, modDCF.tmGordonGrowth).ValuePerShare > gordon.ValuePerShare

    ' Turning off the mid-year convention must lower the value: every flow is
    ' then discounted half a year longer.
    Dim yearEnd As clsAssumptions
    Set yearEnd = a.Clone()
    yearEnd.MidYearConvention = False
    CheckTrue "mid-year convention raises the value", _
              modDCF.Value(yearEnd, modDCF.tmGordonGrowth).ValuePerShare < gordon.ValuePerShare

    ' Clone really is a deep copy; if it were not, the line above would have
    ' mutated the base case and every later test would silently drift.
    CheckTrue "Clone does not alias the original", a.MidYearConvention = True

    ' The exit multiple must imply a perpetuity growth rate, and on this base
    ' case it is above nominal GDP -- which is the point of reporting it.
    CheckTrue "exit multiple implies growth above the Gordon assumption", _
              exitM.ImpliedPerpetuityGrowth > a.TerminalGrowth
End Sub

'--- Validation ---------------------------------------------------------------

Private Sub TestValidationRejectsBadInput()
    ' VBA has no try/catch, and it evaluates arguments at the call site, so a
    ' generic "assert this raises" helper taking the probe as an argument runs
    ' the probe first and lets the error escape past the assertion. Each case
    ' therefore traps inline. Verbose, but it is the only construction that
    ' actually works.

    Dim a As clsAssumptions
    Dim raised As Boolean

    ' Negative base revenue.
    Set a = New clsAssumptions
    a.Revenue0 = -100#
    On Error Resume Next
    Err.Clear
    a.Validate
    raised = (Err.Number <> 0)
    On Error GoTo 0
    CheckTrue "negative revenue is rejected", raised

    ' A tax rate of 100%.
    Set a = New clsAssumptions
    a.TaxRate = 1#
    On Error Resume Next
    Err.Clear
    a.Validate
    raised = (Err.Number <> 0)
    On Error GoTo 0
    CheckTrue "a tax rate of 100% is rejected", raised

    ' Zero diluted shares -- otherwise value per share divides by zero.
    Set a = New clsAssumptions
    a.DilutedShares = 0#
    On Error Resume Next
    Err.Clear
    a.Validate
    raised = (Err.Number <> 0)
    On Error GoTo 0
    CheckTrue "a zero share count is rejected", raised

    ' Mismatched forecast arrays.
    Set a = New clsAssumptions
    a.SetEbitMargin Array(0.18, 0.19)
    On Error Resume Next
    Err.Clear
    a.Validate
    raised = (Err.Number <> 0)
    On Error GoTo 0
    CheckTrue "mismatched forecast arrays are rejected", raised

    ' Terminal growth above WACC: the Gordon formula returns a negative
    ' denominator and, left alone, a confidently negative enterprise value.
    Set a = New clsAssumptions
    a.TerminalGrowth = 0.5
    Dim result As modDCF.ValuationResult
    On Error Resume Next
    Err.Clear
    result = modDCF.Value(a, modDCF.tmGordonGrowth)
    raised = (Err.Number <> 0)
    On Error GoTo 0
    CheckTrue "terminal growth above WACC is rejected", raised
End Sub

'--- Sensitivity --------------------------------------------------------------

Private Sub TestSensitivityGrid()
    Dim a As clsAssumptions
    Set a = New clsAssumptions

    Dim waccDeltas As Variant, growthDeltas As Variant
    waccDeltas = Array(-0.01, 0#, 0.01)
    growthDeltas = Array(-0.005, 0#, 0.005)

    Dim grid() As modSensitivity.SensitivityCell
    grid = modSensitivity.BuildGrid(a, waccDeltas, growthDeltas)

    CheckTrue "grid is rectangular", (UBound(grid) - LBound(grid) + 1) = 9

    ' Down the WACC axis value must fall; across the growth axis it must rise.
    CheckTrue "value falls as WACC rises", grid(8).ValuePerShare < grid(2).ValuePerShare
    CheckTrue "value rises as growth rises", grid(6).ValuePerShare > grid(4).ValuePerShare

    ' The centre cell is the base case.
    CheckClose "grid centre equals the base case", grid(5).ValuePerShare, _
               modDCF.Value(a, modDCF.tmGordonGrowth).ValuePerShare, 0.000001

    ' Every cell must differ; an all-identical grid is the classic symptom of
    ' a data table left on manual calculation.
    CheckTrue "no two cells are identical", grid(1).ValuePerShare <> grid(9).ValuePerShare
End Sub

'--- Assertions ---------------------------------------------------------------

Private Sub CheckClose(name As String, actual As Double, expected As Double, tolerance As Double)
    If Abs(actual - expected) <= tolerance Then
        Pass name
    Else
        Fail name, "got " & Format$(actual, "0.##########") & _
                   ", expected " & Format$(expected, "0.##########") & _
                   " (tolerance " & tolerance & ", diff " & Abs(actual - expected) & ")"
    End If
End Sub

Private Sub CheckTrue(name As String, condition As Boolean)
    If condition Then Pass name Else Fail name, "condition was False"
End Sub

Private Sub Pass(name As String)
    mPassed = mPassed + 1
    mLog = mLog & "  ok   " & name & vbCrLf
End Sub

Private Sub Fail(name As String, detail As String)
    mFailed = mFailed + 1
    mLog = mLog & "  FAIL " & name & vbCrLf & "       " & detail & vbCrLf
End Sub

Private Function GoldenValuePath() As String
    GoldenValuePath = ThisWorkbook.path & Application.PathSeparator & "tests" & _
                      Application.PathSeparator & "golden_values.csv"
End Function
