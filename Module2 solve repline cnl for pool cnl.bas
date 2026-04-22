Attribute VB_Name = "Module2"
Sub Generate_Repline_CNL_From_Overallv2()
    ' This macro generates CNL for each repline based on overall pool CNL in C14
    ' Using regression relationships:
    ' - Repayment type adjustments: full = -2.25% (1% better than IO), IO = -1.25%, partial = baseline, defer = +2%
    ' - Tier increments: +1.5% per tier (from tier_3 baseline) - TIER DOMINATES
    ' - Term increments: term_5 = -0.67%, term_7 = baseline, term_10 = +0.67%, term_15 = +1.0%
    ' - CNL floor: 0.75% minimum (no repline can go below this)
    ' - Iterative calibration ensures weighted average CNL = target in C14
    
    Dim ws As Worksheet
    Dim targetCNL As Double
    Dim lastRow As Long
    Dim row As Long
    
    ' Repline components
    Dim replineName As String
    Dim repaymentType As String
    Dim tier As Integer
    Dim term As Integer
    
    ' CNL calculation
    Dim baseCNL As Double
    Dim repaymentAdj As Double
    Dim tierAdj As Double
    Dim termAdj As Double
    Dim calculatedCNL As Double
    
    ' Iterative adjustment variables
    Dim replineWeights() As Variant
    Dim replineCNLs() As Variant
    Dim weightedAvg As Double
    Dim adjustmentFactor As Double
    Dim iteration As Integer
    Dim tolerance As Double
    Dim maxIterations As Integer
    
    ' Arrays for batch output
    Dim cnlOutputArr() As Variant
    Dim i As Long
    
    Set ws = ActiveSheet
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Get target CNL from C14
    targetCNL = ws.Range("C14").Value
    
    If Not IsNumeric(targetCNL) Or targetCNL = 0 Then
        MsgBox "Please enter a valid target CNL in cell C14.", vbExclamation
        GoTo Cleanup
    End If
    
    ' Find last row with data in column D
    lastRow = ws.Cells(ws.Rows.Count, 4).End(xlUp).row
    
    ' Count actual replines (skip header)
    Dim replineCount As Long
    replineCount = 0
    For row = 31 To lastRow
        If IsNumeric(ws.Cells(row, 4).Value) Then
            replineCount = replineCount + 1
        End If
    Next row
    
    If replineCount = 0 Then
        MsgBox "No replines found starting from row 31.", vbExclamation
        GoTo Cleanup
    End If
    
    ' Initialize arrays
    ReDim replineWeights(1 To replineCount)
    ReDim replineCNLs(1 To replineCount)
    ReDim cnlOutputArr(31 To lastRow, 1 To 1)
    
    ' === STEP 1: Calculate initial CNL for each repline ===
    
    ' Use tier_3 term_7 partial as baseline (reference point)
    ' From your example: partial tier_3 term_7 should equal targetCNL
    baseCNL = targetCNL
    
    i = 0
    For row = 31 To lastRow
        If IsNumeric(ws.Cells(row, 4).Value) Then
            i = i + 1
            
            replineName = Trim(ws.Cells(row, 5).Value)
            
            ' Parse repline name: "repayment tier_X term_Y"
            Call ParseReplineName(replineName, repaymentType, tier, term)
            
            ' === Calculate adjustments ===
            
            ' Repayment type adjustment (relative to "partial")
            ' Full is 1% better than IO, IO is 1.25% better than partial, defer is 2% worse than partial
            Select Case LCase(repaymentType)
                Case "full"
                    repaymentAdj = -0.0225  ' 2.25% better than partial (1% better than IO)
                Case "io"
                    repaymentAdj = -0.0125  ' 1.25% better than partial
                Case "partial"
                    repaymentAdj = 0        ' Baseline
                Case "defer"
                    repaymentAdj = 0.02     ' 2% worse than partial
                Case Else
                    repaymentAdj = 0
            End Select
            
            ' Tier adjustment (tier_3 is baseline, +1.5% per tier increase)
            ' Tier impact dominates term impact
            tierAdj = (tier - 3) * 0.015
            
            ' Term adjustment (term_7 is baseline)
            ' Term adjustments are LESS than tier adjustment (tier dominates)
            ' Max term adjustment = 1.0% (vs tier = 1.5%)
            Select Case term
                Case 5
                    termAdj = -0.0067       ' 0.67% better than term_7
                Case 7
                    termAdj = 0             ' Baseline
                Case 10
                    termAdj = 0.0067        ' 0.67% worse than term_7
                Case 15
                    termAdj = 0.01          ' 1.0% worse than term_7 (less than tier 1.5%)
                Case Else
                    termAdj = 0
            End Select
            
            ' Calculate CNL for this repline
            calculatedCNL = baseCNL + repaymentAdj + tierAdj + termAdj
            
            ' Note: CNL floor of 0.75% will be applied AFTER iterative adjustment
            ' to preserve relative ordering between replines
            
            ' Store calculated CNL and weight
            replineCNLs(i) = calculatedCNL
            replineWeights(i) = ws.Cells(row, 12).Value  ' Column L (%)
            
        End If
    Next row
    
    ' === STEP 2: Iterative adjustment to match target weighted average ===
    
    tolerance = 0.00001     ' 0.001% tolerance
    maxIterations = 100
    iteration = 0
    
    Do While iteration < maxIterations
        ' Calculate weighted average
        weightedAvg = 0
        For i = 1 To replineCount
            weightedAvg = weightedAvg + (replineCNLs(i) * replineWeights(i))
        Next i
        
        ' Check if we're within tolerance
        If Abs(weightedAvg - targetCNL) < tolerance Then
            Exit Do
        End If
        
        ' Calculate adjustment factor
        adjustmentFactor = targetCNL - weightedAvg
        
        ' Apply adjustment to all replines
        For i = 1 To replineCount
            replineCNLs(i) = replineCNLs(i) + adjustmentFactor
        Next i
        
        iteration = iteration + 1
    Loop
    
    ' === STEP 3: Apply CNL floor of 0.75% ===
    ' Applied AFTER iterative adjustment to preserve relative ordering
    
    For i = 1 To replineCount
        If replineCNLs(i) < 0.0075 Then
            replineCNLs(i) = 0.0075
        End If
    Next i
    
    ' === STEP 4: Output calculated CNL to column G ===
    
    i = 0
    For row = 31 To lastRow
        If IsNumeric(ws.Cells(row, 4).Value) Then
            i = i + 1
            cnlOutputArr(row, 1) = replineCNLs(i)
        Else
            cnlOutputArr(row, 1) = ws.Cells(row, 7).Value  ' Keep existing value
        End If
    Next row
    
    ' Batch write to column G
    ws.Range("G31:G" & lastRow).Value = cnlOutputArr
    ws.Range("G31:G" & lastRow).NumberFormat = "0.00%"
    
    ' Calculate final weighted average for verification
    weightedAvg = 0
    For i = 1 To replineCount
        weightedAvg = weightedAvg + (replineCNLs(i) * replineWeights(i))
    Next i
    
Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ' Find example replines for validation
    Dim fullT1T7_CNL As Double, ioT1T7_CNL As Double
    Dim validationMsg As String
    
    ' Search for full tier_1 term_7 and IO tier_1 term_7
    For row = 31 To lastRow
        If IsNumeric(ws.Cells(row, 4).Value) Then
            replineName = Trim(ws.Cells(row, 5).Value)
            If InStr(LCase(replineName), "full tier_1 term_7") > 0 Then
                fullT1T7_CNL = ws.Cells(row, 7).Value
            ElseIf InStr(LCase(replineName), "io tier_1 term_7") > 0 Then
                ioT1T7_CNL = ws.Cells(row, 7).Value
            End If
        End If
    Next row
    
    ' Build validation message
    If fullT1T7_CNL > 0 And ioT1T7_CNL > 0 Then
        validationMsg = vbCrLf & vbCrLf & "Validation:" & vbCrLf & _
                       "full tier_1 term_7: " & Format(fullT1T7_CNL, "0.00%") & vbCrLf & _
                       "IO tier_1 term_7: " & Format(ioT1T7_CNL, "0.00%") & vbCrLf & _
                       "Difference: " & Format(ioT1T7_CNL - fullT1T7_CNL, "0.00%") & " (should be ~1.0%)"
    Else
        validationMsg = ""
    End If
    
    MsgBox "CNL generation complete!" & vbCrLf & vbCrLf & _
           "Target CNL: " & Format(targetCNL, "0.00%") & vbCrLf & _
           "Achieved Weighted Avg: " & Format(weightedAvg, "0.0000%") & vbCrLf & _
           "Iterations: " & iteration & vbCrLf & _
           "Difference: " & Format(Abs(weightedAvg - targetCNL), "0.0000%") & validationMsg, vbInformation

End Sub

Sub ParseReplineName(ByVal replineName As String, _
                     ByRef repaymentType As String, _
                     ByRef tier As Integer, _
                     ByRef term As Integer)
    ' Parse repline name format: "repayment tier_X term_Y"
    ' Example: "partial tier_3 term_7"
    
    Dim parts() As String
    Dim tierPart As String
    Dim termPart As String
    
    ' Split by space
    parts = Split(Trim(replineName), " ")
    
    If UBound(parts) >= 2 Then
        ' Get repayment type (first word)
        repaymentType = LCase(Trim(parts(0)))
        
        ' Get tier (from "tier_X")
        tierPart = Trim(parts(1))
        If InStr(tierPart, "_") > 0 Then
            tier = CInt(Split(tierPart, "_")(1))
        Else
            tier = 3  ' Default
        End If
        
        ' Get term (from "term_Y")
        termPart = Trim(parts(2))
        If InStr(termPart, "_") > 0 Then
            term = CInt(Split(termPart, "_")(1))
        Else
            term = 7  ' Default
        End If
    Else
        ' Defaults if parsing fails
        repaymentType = "partial"
        tier = 3
        term = 7
    End If
    
End Sub
