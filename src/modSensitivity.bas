Attribute VB_Name = "modSensitivity"
Option Explicit

'==============================================================================
' modSensitivity
'
' The WACC x terminal-growth grid, built in code rather than with an Excel data
' table.
'
' A native data table recalculates the whole workbook once per cell and depends
' on calculation mode, iterative settings and the order the sheet happens to
' evaluate in. On a model this size it is slow, and worse, it is the reason
' sensitivity tables so often show the base case in every cell: somebody left
' calculation on manual. Building the grid from a cloned assumption object is
' deterministic, testable, and about a thousand times faster.
'==============================================================================

Private Const ERR_BASE As Long = vbObjectError + 5300

Public Type SensitivityCell
    WaccValue As Double
    TerminalGrowth As Double
    ValuePerShare As Double
    EnterpriseValue As Double
End Type

'------------------------------------------------------------------------------
' Value per share across a grid of WACC and terminal growth offsets.
'
' WACC is moved by solving for the beta that produces the target, not by
' overwriting a stored rate. The cost of equity, the capital structure and the
' cost of debt therefore stay mutually consistent in every cell.
'------------------------------------------------------------------------------
Public Function BuildGrid(baseAssumptions As clsAssumptions, _
                          waccDeltas As Variant, _
                          growthDeltas As Variant) As SensitivityCell()
    baseAssumptions.Validate

    Dim nW As Long, nG As Long
    nW = UBound(waccDeltas) - LBound(waccDeltas) + 1
    nG = UBound(growthDeltas) - LBound(growthDeltas) + 1
    If nW < 1 Or nG < 1 Then
        Err.Raise ERR_BASE + 1, "modSensitivity", "Grid must have at least one row and column"
    End If

    Dim grid() As SensitivityCell
    ReDim grid(1 To nW * nG)

    Dim baseWacc As Double
    baseWacc = modWACC.WACC(baseAssumptions)

    Dim i As Long, j As Long, k As Long
    Dim targetWacc As Double, beta As Double
    Dim scenario As clsAssumptions
    Dim result As modDCF.ValuationResult

    k = 0
    For i = LBound(waccDeltas) To UBound(waccDeltas)
        targetWacc = baseWacc + CDbl(waccDeltas(i))
        beta = modWACC.BetaForTargetWACC(baseAssumptions, targetWacc)

        For j = LBound(growthDeltas) To UBound(growthDeltas)
            k = k + 1

            ' Clone per cell. VBA's Set is a reference assignment, so without a
            ' deep copy every cell would mutate and read the same object and
            ' the grid would collapse to whichever scenario ran last.
            Set scenario = baseAssumptions.Clone()
            scenario.LeveredBeta = beta
            scenario.TerminalGrowth = baseAssumptions.TerminalGrowth + CDbl(growthDeltas(j))

            grid(k).WaccValue = targetWacc
            grid(k).TerminalGrowth = scenario.TerminalGrowth

            ' A cell where growth has been pushed past WACC is genuinely
            ' undefined. Recording it as a sentinel keeps the grid rectangular
            ' instead of aborting the whole table for one corner.
            If scenario.TerminalGrowth >= targetWacc Then
                grid(k).ValuePerShare = 0#
                grid(k).EnterpriseValue = 0#
            Else
                result = modDCF.Value(scenario, modDCF.tmGordonGrowth)
                grid(k).ValuePerShare = result.ValuePerShare
                grid(k).EnterpriseValue = result.EnterpriseValue
            End If
        Next j
    Next i

    BuildGrid = grid
End Function

' Write a grid to a worksheet as a rectangular block with headers.
Public Sub WriteGridToSheet(target As Object, topLeft As String, _
                            grid() As SensitivityCell, _
                            waccDeltas As Variant, growthDeltas As Variant)
    Dim nG As Long
    nG = UBound(growthDeltas) - LBound(growthDeltas) + 1

    Dim anchor As Object
    Set anchor = target.Range(topLeft)

    anchor.Value = "WACC \ g"

    Dim j As Long
    For j = 1 To nG
        anchor.Offset(0, j).Value = grid(j).TerminalGrowth
        anchor.Offset(0, j).NumberFormat = "0.00%"
    Next j

    Dim k As Long, row As Long, col As Long
    For k = LBound(grid) To UBound(grid)
        row = ((k - 1) \ nG) + 1
        col = ((k - 1) Mod nG) + 1

        If col = 1 Then
            anchor.Offset(row, 0).Value = grid(k).WaccValue
            anchor.Offset(row, 0).NumberFormat = "0.00%"
        End If
        anchor.Offset(row, col).Value = grid(k).ValuePerShare
        anchor.Offset(row, col).NumberFormat = "0.00"
    Next k
End Sub
