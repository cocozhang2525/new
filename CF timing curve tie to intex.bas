Attribute VB_Name = "Module1"
'=============================================================
' CNL Loss Curve Module
' Reads cumulative CNL curve from col X, derives timing curve
' in col W, writes monthly MDR to col U
'=============================================================

Sub ApplyCNLLossCurve()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("AmortizationModel")
    
    '-----------------------------------------------------------
    ' HEADERS
    '-----------------------------------------------------------
    ws.Cells(10, 24).Value = "CNL Curve"       ' X10
    ws.Cells(10, 23).Value = "Timing Curve"    ' W10
    ws.Cells(10, 24).Font.Bold = True
    ws.Cells(10, 23).Font.Bold = True
    
    '-----------------------------------------------------------
    ' INPUTS
    '-----------------------------------------------------------
    Dim numPeriods As Long
    numPeriods = ws.Cells(3, 3).Value          ' C3 = Amortization term (max 360)
    
    If numPeriods < 1 Or numPeriods > 360 Then
        MsgBox "Amortization term in C3 must be between 1 and 360.", vbCritical, "Input Error"
        Exit Sub
    End If
    
    '-----------------------------------------------------------
    ' READ CUMULATIVE CNL CURVE FROM COL X (X12:X(11+numPeriods))
    ' Find terminal CNL = last non-zero value in the curve
    '-----------------------------------------------------------
    Dim cnlCurve() As Double
    ReDim cnlCurve(1 To numPeriods)
    
    Dim terminalCNL As Double
    terminalCNL = 0
    
    Dim i As Long
    Dim cellVal As Double
    
    For i = 1 To numPeriods
        cellVal = ws.Cells(11 + i, 24).Value   ' Col X = col 24, starts row 12
        
        If IsEmpty(ws.Cells(11 + i, 24)) Or cellVal = 0 Then
            cnlCurve(i) = terminalCNL
        Else
            cnlCurve(i) = cellVal
            If cellVal > terminalCNL Then
                terminalCNL = cellVal
            End If
        End If
    Next i
    
    ' Validate at least one value was entered
    If terminalCNL = 0 Then
        MsgBox "No CNL curve values found in column X (X12:X" & 11 + numPeriods & "). Please enter your cumulative CNL curve.", _
               vbCritical, "Input Error"
        Exit Sub
    End If
    
    ' Write terminal CNL to I2
    ws.Range("I2").Value = terminalCNL
    
    '-----------------------------------------------------------
    ' DERIVE TIMING CURVE IN COL W
    ' W[t] = (CNL[t] - CNL[t-1]) / terminalCNL
    ' Sums to 100% by construction
    '-----------------------------------------------------------
    Dim timingCurve() As Double
    ReDim timingCurve(1 To numPeriods)
    
    Dim curveSum As Double
    Dim priorCNL As Double
    curveSum = 0
    
    For i = 1 To numPeriods
        If i = 1 Then
            priorCNL = 0
        Else
            priorCNL = cnlCurve(i - 1)
        End If
        timingCurve(i) = (cnlCurve(i) - priorCNL) / terminalCNL
        curveSum = curveSum + timingCurve(i)
        ws.Cells(11 + i, 23).Value = timingCurve(i)   ' Write to col W
    Next i
    
    ' Clear any leftover values below the used range
    If numPeriods < 360 Then
        ws.Range(ws.Cells(12 + numPeriods, 23), ws.Cells(372, 23)).ClearContents  ' Col W
    End If
    
    ' Warn if timing curve doesn't sum to ~100%
    If Abs(curveSum - 1) > 0.001 Then
        MsgBox "Warning: Timing curve sums to " & Format(curveSum, "0.00%") & " (expected 100.00%)." & vbCrLf & _
               "This usually means your CNL curve in col X does not reach its terminal value by period " & numPeriods & "." & vbCrLf & _
               "MDR will be applied as-is but total losses may be understated.", _
               vbExclamation, "Timing Curve Check"
    End If
    
    '-----------------------------------------------------------
    ' WRITE MONTHLY MDR TO COL U (col 21), rows 12 to 11+numPeriods
    ' Loss $[t] = terminalCNL * originalUPB * timingCurve[t]
    ' MDR[t]    = Loss $[t] / prior month ending balance (col D)
    ' Period 1 (row 12): uses row 11 ending balance as denominator
    ' Col T (annual CDR) is left untouched
    '-----------------------------------------------------------
    Dim originalUPB As Double
    originalUPB = ws.Cells(1, 3).Value         ' C1 = Current UPB

    Dim priorEndBal As Double
    Dim lossAmt As Double
    Dim mdr As Double
    
    For i = 1 To numPeriods
        priorEndBal = ws.Cells(11 + i - 1, 4).Value   ' Col D = Ending Balance, prior row
        ' i=1 ? row 11 (period 0 ending balance)
        ' i=2 ? row 12 (period 1 ending balance), etc.
        
        If priorEndBal > 0 Then
            lossAmt = terminalCNL * originalUPB * timingCurve(i)
            mdr = lossAmt / priorEndBal
        Else
            mdr = 0
        End If
        
        ws.Cells(11 + i, 21).Value = mdr        ' Write to col U (col 21)
    Next i
    
    ' Clear MDR below used range
    If numPeriods < 360 Then
        ws.Range(ws.Cells(12 + numPeriods, 21), ws.Cells(372, 21)).ClearContents  ' Col U
    End If
    
    MsgBox "Done." & vbCrLf & _
           "Terminal CNL: " & Format(terminalCNL, "0.000%") & vbCrLf & _
           "Timing curve sum: " & Format(curveSum, "0.000%") & vbCrLf & _
           "Periods updated: " & numPeriods, _
           vbInformation, "CNL Loss Curve Applied"

End Sub

