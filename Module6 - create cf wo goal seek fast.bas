Attribute VB_Name = "Module6"

Sub Solve_Repline_CDR_From_CNL_GOALSEEK_V3()
    Dim wsAssump As Worksheet
    Dim wsCDR As Worksheet
    Dim wsRepline As Worksheet
    Dim replineRow As Long
    Dim replineNum As Long
    Dim targetCNL As Double
    Dim sheetName As String
    Dim solvedCDR As Double
    Dim lastRow As Long
    
    ' Arrays for batch output
    Dim arrCDROutput() As Variant
    Dim arrAssumpOutput() As Variant
    Dim cdrRowLookup As Object
    Dim outputRow As Long
    
    Set wsAssump = ThisWorkbook.Sheets("Assumption")
    Set wsCDR = ThisWorkbook.Sheets("CDR CPR")
    Set cdrRowLookup = CreateObject("Scripting.Dictionary")
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic  ' CRITICAL: Must be Automatic for Goal Seek
    Application.EnableEvents = False
    Application.DisplayStatusBar = True
    
    ' Find last row with repline number in column D of CDR CPR sheet
    lastRow = wsCDR.Cells(wsCDR.Rows.Count, "D").End(xlUp).Row
    
    ' Pre-build lookup dictionary (maps repline number to CDR CPR sheet row)
    For outputRow = 31 To lastRow
        If IsNumeric(wsCDR.Cells(outputRow, "D").Value) Then
            If Not cdrRowLookup.exists(CLng(wsCDR.Cells(outputRow, "D").Value)) Then
                cdrRowLookup.Add CLng(wsCDR.Cells(outputRow, "D").Value), outputRow
            End If
        End If
    Next outputRow
    
    ' Initialize arrays
    ReDim arrCDROutput(31 To lastRow, 1 To 1)
    ReDim arrAssumpOutput(39 To 340, 1 To 1)
    
    ' Load existing values
    For outputRow = 31 To lastRow
        arrCDROutput(outputRow, 1) = wsCDR.Cells(outputRow, "I").Value
    Next outputRow
    
    For outputRow = 39 To 340
        arrAssumpOutput(outputRow, 1) = wsAssump.Cells(outputRow, "S").Value
    Next outputRow
    
    ' Main loop through CDR CPR sheet rows 31 to last row
    For replineRow = 31 To lastRow
        
        ' Get repline number from column D
        If IsNumeric(wsCDR.Cells(replineRow, "D").Value) Then
            replineNum = CLng(wsCDR.Cells(replineRow, "D").Value)
        Else
            GoTo NextIteration
        End If
        
        ' Get target CNL from column G
        If IsNumeric(wsCDR.Cells(replineRow, "G").Value) Then
            targetCNL = wsCDR.Cells(replineRow, "G").Value
        Else
            GoTo NextIteration
        End If
        
        If replineNum >= 1 And replineNum <= 299 And targetCNL <> 0 Then
            
            sheetName = "Repline " & replineNum & " CF"
            
            On Error Resume Next
            Set wsRepline = ThisWorkbook.Sheets(sheetName)
            On Error GoTo 0
            
            If Not wsRepline Is Nothing Then
                
                Application.StatusBar = "Goal seeking Repline " & replineNum & " for CNL " & Format(targetCNL, "0.00%") & "..."
                
                ' Goal Seek: Set I1 (CNL) to target by changing T9 (CDR)
                On Error Resume Next
                wsRepline.Range("I1").GoalSeek _
                    Goal:=targetCNL, _
                    ChangingCell:=wsRepline.Range("T9")
                On Error GoTo 0
                
                solvedCDR = wsRepline.Range("T9").Value
                
                ' Store in arrays - output to CDR CPR column I
                arrCDROutput(replineRow, 1) = solvedCDR
                
                ' Also output to Assumption column S (find matching row in Assumption tab)
                Dim assumpRow As Long
                For assumpRow = 39 To 340
                    If IsNumeric(wsAssump.Cells(assumpRow, "C").Value) Then
                        If CLng(wsAssump.Cells(assumpRow, "C").Value) = replineNum Then
                            arrAssumpOutput(assumpRow, 1) = solvedCDR
                            Exit For
                        End If
                    End If
                Next assumpRow
                
            End If
            
            Set wsRepline = Nothing
        End If
        
NextIteration:
    Next replineRow
    
    ' Batch write outputs
    wsCDR.Range("I31:I" & lastRow).Value = arrCDROutput
    wsCDR.Range("I31:I" & lastRow).NumberFormat = "0.00%"
    
    wsAssump.Range("S39:S340").Value = arrAssumpOutput
    wsAssump.Range("S39:S340").NumberFormat = "0.00%"
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = False
    Application.StatusBar = False
    
    MsgBox "Goal Seek Complete! Processed rows 31 to " & lastRow, vbInformation
End Sub
Sub Solve_Repline_CDR_From_CNL_GOALSEEK_v2()
    Dim wsAssump As Worksheet
    Dim wsCDR As Worksheet
    Dim wsRepline As Worksheet
    Dim replineRow As Long
    Dim replineNum As Long
    Dim scenarioNum As Long
    Dim scenColCDR As Long
    Dim targetCNL As Double
    Dim sheetName As String
    Dim solvedCDR As Double
    
    ' Arrays for batch output
    Dim arrCDROutput() As Variant
    Dim arrAssumpOutput() As Variant
    Dim cdrRowLookup As Object
    Dim outputRow As Long
    
    Set wsAssump = ThisWorkbook.Sheets("Assumption")
    Set wsCDR = ThisWorkbook.Sheets("CDR CPR")
    Set cdrRowLookup = CreateObject("Scripting.Dictionary")
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic  ' CRITICAL: Must be Automatic for Goal Seek
    Application.EnableEvents = False
    Application.DisplayStatusBar = True
    
    ' Pre-build lookup dictionary
    For outputRow = 31 To 340
        If IsNumeric(wsCDR.Cells(outputRow, "B").Value) Then
            If Not cdrRowLookup.exists(CLng(wsCDR.Cells(outputRow, "B").Value)) Then
                cdrRowLookup.Add CLng(wsCDR.Cells(outputRow, "B").Value), outputRow
            End If
        End If
    Next outputRow
    
    ' Initialize arrays
    ReDim arrCDROutput(31 To 340, 1 To 1)
    ReDim arrAssumpOutput(39 To 340, 1 To 1)
    
    ' Load existing values
    For outputRow = 31 To 340
        arrCDROutput(outputRow, 1) = wsCDR.Cells(outputRow, "I").Value
    Next outputRow
    
    For outputRow = 39 To 340
        arrAssumpOutput(outputRow, 1) = wsAssump.Cells(outputRow, "S").Value
    Next outputRow
    
    ' Main loop
    For replineRow = 39 To 340
        
        If IsNumeric(wsAssump.Cells(replineRow, "C").Value) Then
            replineNum = CLng(wsAssump.Cells(replineRow, "C").Value)
        Else
            GoTo NextIteration
        End If
        
        If IsNumeric(wsAssump.Cells(replineRow, "N").Value) Then
            scenarioNum = CLng(wsAssump.Cells(replineRow, "N").Value)
        Else
            GoTo NextIteration
        End If
        
        If replineNum >= 1 And replineNum <= 299 And scenarioNum >= 1 Then
            
            ' Get target CNL from CDR CPR sheet row 25
            scenColCDR = 6 + (scenarioNum - 1)  ' Scenario 1 = col F (6), Scenario 2 = col G (7), etc.
            targetCNL = wsCDR.Cells(25, scenColCDR).Value
            
            If IsNumeric(targetCNL) And targetCNL <> 0 Then
                
                sheetName = "Repline " & replineNum & " CF"
                
                On Error Resume Next
                Set wsRepline = ThisWorkbook.Sheets(sheetName)
                On Error GoTo 0
                
                If Not wsRepline Is Nothing Then
                    
                    Application.StatusBar = "Goal seeking Repline " & replineNum & " for CNL " & Format(targetCNL, "0.00%") & "..."
                    
                    ' Goal Seek: Set I1 (CNL) to target by changing T9 (CDR)
                    On Error Resume Next
                    wsRepline.Range("I1").GoalSeek _
                        Goal:=targetCNL, _
                        ChangingCell:=wsRepline.Range("T9")
                    On Error GoTo 0
                    
                    solvedCDR = wsRepline.Range("T9").Value
                    
                    ' Store in arrays
                    If cdrRowLookup.exists(replineNum) Then
                        arrCDROutput(CLng(cdrRowLookup(replineNum)), 1) = solvedCDR
                    End If
                    
                    arrAssumpOutput(replineRow, 1) = solvedCDR
                    
                End If
                
                Set wsRepline = Nothing
            End If
        End If
        
NextIteration:
    Next replineRow
    
    ' Batch write outputs
    wsCDR.Range("I31:I340").Value = arrCDROutput
    wsCDR.Range("I31:I340").NumberFormat = "0.00%"
    
    wsAssump.Range("S39:S340").Value = arrAssumpOutput
    wsAssump.Range("S39:S340").NumberFormat = "0.00%"
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = False
    Application.StatusBar = False
    
    MsgBox "Goal Seek Complete!", vbInformation
End Sub
Sub Solve_Repline_CDR_From_CNL_GOALSEEK()
    Dim wsAssump As Worksheet
    Dim wsCDR As Worksheet
    Dim wsRepline As Worksheet
    Dim replineRow As Long
    Dim replineNum As Long
    Dim scenarioNum As Long
    Dim scenColCDR As Long
    Dim targetCNL As Double
    Dim sheetName As String
    Dim solvedCDR As Double
    
    ' Arrays for batch output
    Dim arrCDROutput() As Variant
    Dim arrAssumpOutput() As Variant
    Dim cdrRowLookup As Object
    Dim outputRow As Long
    
    Set wsAssump = ThisWorkbook.Sheets("Assumption")
    Set wsCDR = ThisWorkbook.Sheets("CDR CPR")
    Set cdrRowLookup = CreateObject("Scripting.Dictionary")
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic  ' CRITICAL: Must be Automatic for Goal Seek
    Application.EnableEvents = False
    Application.DisplayStatusBar = True
    
    ' Pre-build lookup dictionary
    For outputRow = 31 To 340
        If IsNumeric(wsCDR.Cells(outputRow, "B").Value) Then
            If Not cdrRowLookup.exists(CLng(wsCDR.Cells(outputRow, "B").Value)) Then
                cdrRowLookup.Add CLng(wsCDR.Cells(outputRow, "B").Value), outputRow
            End If
        End If
    Next outputRow
    
    ' Initialize arrays
    ReDim arrCDROutput(31 To 340, 1 To 1)
    ReDim arrAssumpOutput(39 To 340, 1 To 1)
    
    ' Load existing values
    For outputRow = 31 To 340
        arrCDROutput(outputRow, 1) = wsCDR.Cells(outputRow, "I").Value
    Next outputRow
    
    For outputRow = 39 To 340
        arrAssumpOutput(outputRow, 1) = wsAssump.Cells(outputRow, "S").Value
    Next outputRow
    
    ' Main loop
    For replineRow = 39 To 340
        
        If IsNumeric(wsAssump.Cells(replineRow, "C").Value) Then
            replineNum = CLng(wsAssump.Cells(replineRow, "C").Value)
        Else
            GoTo NextIteration
        End If
        
        If IsNumeric(wsAssump.Cells(replineRow, "N").Value) Then
            scenarioNum = CLng(wsAssump.Cells(replineRow, "N").Value)
        Else
            GoTo NextIteration
        End If
        
        If replineNum >= 1 And replineNum <= 299 And scenarioNum >= 1 Then
            
            ' Get target CNL from CDR CPR sheet row 25
            scenColCDR = 6 + (scenarioNum - 1)  ' Scenario 1 = col F (6), Scenario 2 = col G (7), etc.
            targetCNL = wsCDR.Cells(25, scenColCDR).Value
            
            If IsNumeric(targetCNL) And targetCNL <> 0 Then
                
                sheetName = "Repline " & replineNum & " CF"
                
                On Error Resume Next
                Set wsRepline = ThisWorkbook.Sheets(sheetName)
                On Error GoTo 0
                
                If Not wsRepline Is Nothing Then
                    
                    Application.StatusBar = "Goal seeking Repline " & replineNum & " for CNL " & Format(targetCNL, "0.00%") & "..."
                    
                    ' Goal Seek: Set I1 (CNL) to target by changing T9 (CDR)
                    On Error Resume Next
                    wsRepline.Range("I1").GoalSeek _
                        Goal:=targetCNL, _
                        ChangingCell:=wsRepline.Range("T9")
                    On Error GoTo 0
                    
                    solvedCDR = wsRepline.Range("T9").Value
                    
                    ' Store in arrays
                    If cdrRowLookup.exists(replineNum) Then
                        arrCDROutput(CLng(cdrRowLookup(replineNum)), 1) = solvedCDR
                    End If
                    
                    arrAssumpOutput(replineRow, 1) = solvedCDR
                    
                End If
                
                Set wsRepline = Nothing
            End If
        End If
        
NextIteration:
    Next replineRow
    
    ' Batch write outputs
    wsCDR.Range("I31:I340").Value = arrCDROutput
    wsCDR.Range("I31:I340").NumberFormat = "0.00%"
    
    wsAssump.Range("S39:S340").Value = arrAssumpOutput
    wsAssump.Range("S39:S340").NumberFormat = "0.00%"
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = False
    
    MsgBox "Goal Seek Complete!", vbInformation
End Sub

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
                .Range("T9").Value = wsAssump.Cells(scenHeaderRow + 5, scenCol).Value
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

            ' ***CRITICAL: Copy WITH FORMULAS using xlPasteAll***
            wsModel.Cells.Copy
            ws.Cells.PasteSpecial xlPasteAll
            Application.CutCopyMode = False

            ' Cutoff tail after C3 + 70 months
            If IsNumeric(ws.Range("C3").Value) Then
                cutoffMonth = ws.Range("C3").Value + 70
                lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
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

            lastRow = ws.Cells(ws.Rows.Count, "O").End(xlUp).Row
            
            ' Array-based PV calculations
            If lastRow >= 12 Then
                monthArr = ws.Range("B12:B" & lastRow).Value

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

        lastDataRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
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

Sub ISL_Create_CF_Without_GoalSeek_FAST()
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
                lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
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

            lastRow = ws.Cells(ws.Rows.Count, "O").End(xlUp).Row
            
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

        lastDataRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
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
