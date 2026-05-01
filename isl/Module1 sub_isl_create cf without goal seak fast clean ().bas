Attribute VB_Name = "Module1"
Sub Generate_Repline_CNL_From_Overall_final()
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

Sub ISL_Create_CF_Without_GoalSeek_FAST_clean()
    Dim wsAssump As Worksheet, wsModel As Worksheet
    Dim wsPool As Worksheet, ws As Worksheet, wsRepline As Worksheet
    Dim replineSheets As Collection
    Dim replineNum As Variant
    Dim replineRow As Long, i As Long, col As Long, j As Long
    Dim scenarioNum As Variant, scenCol As Long, scenHeaderRow As Long
    Dim sheetName As String
    Dim lastRow As Long, cutoffMonth As Long, blankStartRow As Long
    Dim discRate As Double, thisMonth As Variant, thisCF As Variant

    Dim pool_initial_bal As Double
    Dim sumWeightedRate As Double, sumWeight As Double
    Dim poolDiscRate As Double

    Dim totalC1 As Double, waC2 As Double, waC3 As Double
    Dim wsReplineW As Worksheet, c1 As Double, c2 As Double, c3 As Double

    Dim wsAll As Worksheet
    Dim blockList As Collection
    Dim pasteRowHdr As Long, pasteRowLabel As Long, pasteRowId As Long
    Dim firstDataCol As Long, blockCols As Long, lastDataRow As Long, ii As Long
    Dim blockName As String
    Dim rngSrc As Range, rngDst As Range
    Dim sepCol As Long, cLabel As Long
    
    ' Arrays for batch operations
    Dim monthArr As Variant, cfArr As Variant, pvArr() As Variant
    Dim arrSize As Long
    Dim priceOutputArr() As Variant
    Dim weight As Double, rate As Double
    Dim cdrValue As Variant
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayStatusBar = False
    Application.DisplayAlerts = False
    
    On Error GoTo ErrorHandler

    Set wsAssump = ThisWorkbook.Sheets("Assumption")
    Set wsModel = ThisWorkbook.Sheets("AmortizationModel")
    Set replineSheets = New Collection
    
    ReDim priceOutputArr(39 To 299, 1 To 1)

    ' Delete old sheets
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "Repline * CF" Then
            ws.Delete
        End If
    Next ws

    ' MAIN REPLINE LOOP (rows 39 to 299)
    For replineRow = 39 To 299
        replineNum = wsAssump.Cells(replineRow, "C").Value
        scenarioNum = wsAssump.Cells(replineRow, "N").Value

        If IsNumeric(replineNum) And replineNum >= 1 And replineNum <= 299 _
           And IsNumeric(scenarioNum) Then

            scenHeaderRow = 22
            scenCol = 0
            For j = 1 To wsAssump.Cells(scenHeaderRow, wsAssump.Columns.Count).End(xlToLeft).Column
                If wsAssump.Cells(scenHeaderRow, j).Value = scenarioNum Then
                    scenCol = j
                    Exit For
                End If
            Next j
            If scenCol = 0 Then GoTo NextRepline

            ' Batch update all model inputs
            With wsModel
                .Range("R9").Value = wsAssump.Cells(scenHeaderRow + 4, scenCol).Value
                
                ' *** CORRECTED: Read CDR from column S (goal-seeked value) if it exists, otherwise use scenario default ***
                cdrValue = wsAssump.Cells(replineRow, "S").Value
                If IsNumeric(cdrValue) And cdrValue <> "" And cdrValue <> 0 Then
                    .Range("T9").Value = cdrValue
                Else
                    .Range("T9").Value = wsAssump.Cells(scenHeaderRow + 5, scenCol).Value
                End If
                
                .Range("N9").Value = wsAssump.Cells(scenHeaderRow + 8, scenCol).Value
                
                .Range("C1").Value = wsAssump.Cells(replineRow, "E").Value
                .Range("C2").Value = wsAssump.Cells(replineRow, "H").Value
                .Range("C3").Value = wsAssump.Cells(replineRow, "F").Value
                .Range("C4").Value = wsAssump.Cells(replineRow, "K").Value
                .Range("C5").Value = wsAssump.Cells(replineRow, "J").Value
                .Range("C6").Value = wsAssump.Cells(replineRow, "L").Value
                .Range("C7").Value = wsAssump.Cells(replineRow, "I").Value
                .Range("C8").Value = wsAssump.Cells(replineRow, "M").Value
            End With

            wsModel.Calculate

            ' Create repline CF sheet
            sheetName = "Repline " & replineNum & " CF"
            On Error Resume Next
            ThisWorkbook.Sheets(sheetName).Delete
            On Error GoTo ErrorHandler

            Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            ws.Name = sheetName
            replineSheets.Add ws

            ' Copy WITH FORMULAS
            wsModel.Cells.Copy
            ws.Cells.PasteSpecial xlPasteAll
            Application.CutCopyMode = False

            ' Cutoff tail after C3 + 70 months
            If IsNumeric(ws.Range("C3").Value) Then
                cutoffMonth = ws.Range("C3").Value + 70
                lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
                blankStartRow = 0
                For j = 12 To lastRow
                    If IsNumeric(ws.Cells(j, "B").Value) Then
                        If ws.Cells(j, "B").Value > cutoffMonth Then
                            blankStartRow = j
                            Exit For
                        End If
                    End If
                Next j
                If blankStartRow > 0 And blankStartRow <= lastRow Then
                    ws.Range("A" & blankStartRow & ":Z" & ws.Rows.Count).ClearContents
                End If
            End If

            ws.Columns("R:U").Hidden = True

            ' Repline-level discounting in column P, F1/F2
            discRate = wsAssump.Cells(replineRow, "M").Value
            ws.Range("P9").Value = discRate

            lastRow = ws.Cells(ws.Rows.Count, "O").End(xlUp).row
            
            ' Array-based PV calculations
            If lastRow >= 12 Then
                monthArr = ws.Range("B12:B" & lastRow).Value
                cfArr = ws.Range("O12:O" & lastRow).Value
                arrSize = UBound(monthArr, 1)
                
                ReDim pvArr(1 To arrSize, 1 To 1)
                
                For i = 1 To arrSize
                    If IsNumeric(monthArr(i, 1)) And IsNumeric(cfArr(i, 1)) Then
                        pvArr(i, 1) = cfArr(i, 1) / ((1 + discRate) ^ (monthArr(i, 1) / 12))
                    Else
                        pvArr(i, 1) = ""
                    End If
                Next i
                
                ws.Range("P12:P" & lastRow).Value = pvArr
                ws.Range("P12:P" & lastRow).NumberFormat = "#,##0.00"
            End If

            Dim pvSum As Double
            pvSum = Application.WorksheetFunction.Sum(ws.Range("P12:P" & lastRow))
            ws.Range("F1").Value = pvSum
            If ws.Range("C1").Value <> 0 Then
                ws.Range("F2").Value = pvSum / ws.Range("C1").Value
                ws.Range("F2").NumberFormat = "0.00%"
                priceOutputArr(replineRow, 1) = ws.Range("F2").Value
            Else
                ws.Range("F2").Value = ""
                priceOutputArr(replineRow, 1) = ""
            End If
        End If
NextRepline:
    Next replineRow

    ' Batch write all prices
    wsAssump.Range("P39:P299").Value = priceOutputArr
    wsAssump.Range("P39:P299").NumberFormat = "0.0000"

    ' SUM POOL INITIAL BALANCE (UPB)
    pool_initial_bal = 0
    For Each wsRepline In replineSheets
        If IsNumeric(wsRepline.Range("C1").Value) Then
            pool_initial_bal = pool_initial_bal + wsRepline.Range("C1").Value
        End If
    Next wsRepline

    ' --- POOL CF AGGREGATION ---
    On Error Resume Next
    ThisWorkbook.Sheets("Pool CF").Delete
    On Error GoTo ErrorHandler

    Set wsPool = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsPool.Name = "Pool CF"

    wsPool.Columns("R:U").Hidden = True
    wsModel.Range("R9,T9,N9,P9").ClearContents
    
    wsModel.Cells.Copy
    wsPool.Cells.PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    
    wsPool.Range("A10:Z10").Value = wsModel.Range("A10:Z10").Value

    ' Aggregate pool data
    Dim poolData() As Variant
    ReDim poolData(11 To 1999, 3 To 15)
    
    For i = 11 To 1999
        wsPool.Cells(i, 2).Value = wsModel.Cells(i, 2).Value
        For col = 3 To 15
            Dim total As Double: total = 0
            For Each ws In replineSheets
                If IsNumeric(ws.Cells(i, col).Value) Then
                    total = total + ws.Cells(i, col).Value
                End If
            Next ws
            poolData(i, col) = total
        Next col
    Next i
    
    wsPool.Range("C11:O1999").Value = poolData

    ' Assign the pool initial UPB
    wsPool.Range("C1").Value = pool_initial_bal

    ' Weighted avg discount
    sumWeightedRate = 0
    sumWeight = 0
    For replineRow = 39 To 299
        replineNum = wsAssump.Cells(replineRow, "C").Value
        If IsNumeric(replineNum) And replineNum >= 1 And replineNum <= 299 Then
            weight = wsAssump.Cells(replineRow, "E").Value
            rate = wsAssump.Cells(replineRow, "M").Value
            If IsNumeric(weight) And IsNumeric(rate) Then
                sumWeightedRate = sumWeightedRate + weight * rate
                sumWeight = sumWeight + weight
            End If
        End If
    Next replineRow

    If sumWeight <> 0 Then
        poolDiscRate = sumWeightedRate / sumWeight
    Else
        poolDiscRate = 0
    End If
    wsPool.Range("C8").Value = poolDiscRate

    ' Calculate weighted average for Pool CF C2:C3
    totalC1 = 0: waC2 = 0: waC3 = 0
    For Each wsReplineW In replineSheets
        If IsNumeric(wsReplineW.Range("C1").Value) Then c1 = wsReplineW.Range("C1").Value Else c1 = 0
        If IsNumeric(wsReplineW.Range("C2").Value) Then c2 = wsReplineW.Range("C2").Value Else c2 = 0
        If IsNumeric(wsReplineW.Range("C3").Value) Then c3 = wsReplineW.Range("C3").Value Else c3 = 0
        totalC1 = totalC1 + c1
        waC2 = waC2 + c2 * c1
        waC3 = waC3 + c3 * c1
    Next wsReplineW
    If totalC1 > 0 Then
        wsPool.Range("C2").Value = waC2 / totalC1
        wsPool.Range("C3").Value = waC3 / totalC1
    End If

    ' --- Build ALL CF ---
    Set blockList = New Collection
    On Error Resume Next
    ThisWorkbook.Sheets("ALL CF").Delete
    On Error GoTo ErrorHandler

    Set wsAll = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsAll.Name = "ALL CF"

    blockList.Add "Pool CF"
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "Repline * CF" Then blockList.Add ws.Name
    Next ws

    pasteRowHdr = 12
    pasteRowLabel = 8
    pasteRowId = 10
    firstDataCol = 2
    blockCols = 15

    For ii = 1 To blockList.Count
        blockName = blockList(ii)
        Set ws = ThisWorkbook.Sheets(blockName)

        lastDataRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
        If lastDataRow < pasteRowHdr Then lastDataRow = pasteRowHdr + 1

        wsAll.Cells(pasteRowLabel, firstDataCol).Value = blockName
        wsAll.Range(wsAll.Cells(pasteRowLabel, firstDataCol), wsAll.Cells(pasteRowLabel, firstDataCol + blockCols - 1)).Merge
        wsAll.Cells(pasteRowLabel, firstDataCol).Font.Bold = True

        cLabel = firstDataCol - 1
        wsAll.Cells(pasteRowId, cLabel).Value = blockName
        wsAll.Cells(pasteRowId, cLabel).Font.Bold = True

        wsAll.Range(wsAll.Cells(1, firstDataCol), wsAll.Cells(pasteRowHdr, firstDataCol + blockCols - 1)).Value = _
            ws.Range(ws.Cells(1, 2), ws.Cells(pasteRowHdr, 16)).Value

        If blockName = "Pool CF" Then
            wsAll.Cells(2, firstDataCol + 1).Value = wsPool.Range("C2").Value
            wsAll.Cells(3, firstDataCol + 1).Value = wsPool.Range("C3").Value
        End If

        If lastDataRow > pasteRowHdr Then
            wsAll.Range(wsAll.Cells(pasteRowHdr + 1, firstDataCol), wsAll.Cells(lastDataRow, firstDataCol + blockCols - 1)).Value = _
                ws.Range(ws.Cells(pasteRowHdr + 1, 2), ws.Cells(lastDataRow, 16)).Value
        End If

        Set rngSrc = ws.Range(ws.Cells(1, 2), ws.Cells(lastDataRow, 16))
        Set rngDst = wsAll.Range(wsAll.Cells(1, firstDataCol), wsAll.Cells(lastDataRow, firstDataCol + blockCols - 1))
        rngSrc.Copy
        rngDst.PasteSpecial xlPasteFormats
        Application.CutCopyMode = False

        sepCol = firstDataCol + blockCols + 1
        wsAll.Range(wsAll.Cells(1, sepCol + 1), wsAll.Cells(lastDataRow, sepCol + 1)).Interior.Color = RGB(200, 200, 200)
        firstDataCol = firstDataCol + blockCols + 3
    Next ii

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    
    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Else
        MsgBox "ALL CF created successfully! Processed " & replineSheets.Count & " replines.", vbInformation
    End If
End Sub


