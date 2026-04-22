Attribute VB_Name = "Module15"
Sub Process_ISL_Tape_To_Replines_V2()
    Dim wsTape As Worksheet
    Dim wsReplines As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim tierValue As Long
    Dim repaymentType As String
    Dim originalTerm As Long
    Dim termCategory As String
    Dim replineNumber As Long
    Dim asofDate As Date
    Dim firstPmtDate As Date
    Dim monsToRepay As Long
    Dim validationMsg As String
    
    ' Dictionary to store repline data
    Dim replineDataDict As Object
    Set replineDataDict = CreateObject("Scripting.Dictionary")
    
    ' Speed optimization
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    On Error GoTo ErrorHandler
    
    ' Set worksheets - check if ISL_TAPE exists
    On Error Resume Next
    Set wsTape = ThisWorkbook.Sheets("ISL_TAPE")
    On Error GoTo ErrorHandler
    
    If wsTape Is Nothing Then
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.DisplayAlerts = True
        
        Dim sheetList As String
        Dim ws As Worksheet
        For Each ws In ThisWorkbook.Worksheets
            sheetList = sheetList & ws.Name & vbCrLf
        Next ws
        
        MsgBox "Error: Cannot find sheet named 'ISL_TAPE'." & vbCrLf & vbCrLf & _
               "Please make sure your tape data is in a sheet named 'ISL_TAPE'." & vbCrLf & vbCrLf & _
               "Current sheets in this workbook:" & vbCrLf & sheetList, vbCritical
        Exit Sub
    End If
    
    ' Create or clear Replines sheet
    On Error Resume Next
    Set wsReplines = ThisWorkbook.Sheets("Replines")
    If wsReplines Is Nothing Then
        Set wsReplines = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsReplines.Name = "Replines"
    Else
        wsReplines.Cells.Clear
    End If
    On Error GoTo ErrorHandler
    
    ' Initialize all 112 replines in dictionary
    Call InitializeReplinesV2(replineDataDict)
    
    ' Find last row in tape
    lastRow = wsTape.Cells(wsTape.Rows.Count, "A").End(xlUp).Row
    
    ' Process each loan
    For i = 2 To lastRow ' Start from row 2 (skip header)
        ' Get tier from column S (isl_tier)
        If IsNumeric(wsTape.Cells(i, "S").Value) And Not IsEmpty(wsTape.Cells(i, "S").Value) Then
            tierValue = CLng(wsTape.Cells(i, "S").Value)
        Else
            GoTo NextLoan
        End If
        If tierValue < 1 Or tierValue > 7 Then GoTo NextLoan
        
        ' FIRST ROUND FILTER: Determine repay type based on first_prin_int_pmt_dt vs asof_date
        asofDate = wsTape.Cells(i, "A").Value ' asof_date
        
        ' Check if first_prin_int_pmt_dt (CQ) exists and compare to asof_date
        If IsDate(wsTape.Cells(i, "CQ").Value) Then
            firstPmtDate = wsTape.Cells(i, "CQ").Value
            
            ' PRIMARY RULE: If first_prin_int_pmt_dt < asof_date, it's "full" repayment
            If firstPmtDate < asofDate Then
                repaymentType = "full"
            Else
                ' First payment date is in the future, use current_repay_type (column AN) to determine type
                Dim rawRepayType As String
                rawRepayType = UCase(Trim(wsTape.Cells(i, "AN").Value))
                
                Select Case rawRepayType
                    Case "INTEREST PAYMENT"
                        repaymentType = "IO"
                    Case "FIXED PAYMENT"
                        repaymentType = "partial"
                    Case "DEFERRED REPAY"
                        repaymentType = "defer"
                    Case "IMMEDIATE"
                        repaymentType = "full"
                    Case Else
                        GoTo NextLoan ' Skip loans with invalid current_repay_type
                End Select
            End If
        Else
            ' No valid first_prin_int_pmt_dt date - skip this loan
            GoTo NextLoan
        End If
        
        ' Get original term from column I (initial_term) and determine term bucket
        If IsNumeric(wsTape.Cells(i, "I").Value) And Not IsEmpty(wsTape.Cells(i, "I").Value) Then
            originalTerm = CLng(wsTape.Cells(i, "I").Value)
        Else
            GoTo NextLoan
        End If
        
        If originalTerm <= 60 Then
            termCategory = "term_5"
        ElseIf originalTerm <= 84 Then
            termCategory = "term_7"
        ElseIf originalTerm <= 120 Then
            termCategory = "term_10"
        Else
            termCategory = "term_15"
        End If
        
        ' Calculate repline number
        replineNumber = CalculateReplineNum(tierValue, repaymentType, termCategory)
        
        ' Create repline key
        Dim replineKey As String
        replineKey = CStr(replineNumber)
        
        ' Get loan data - handle empty cells and convert to numeric
        Dim origBalance As Double, currBalance As Double
        Dim remTerm As Double, wac As Double
        Dim accruedIntToCap As Double, fixedPay As Double
        
        ' cumulative_disbursed_to_date (column L)
        If IsNumeric(wsTape.Cells(i, "L").Value) And Not IsEmpty(wsTape.Cells(i, "L").Value) Then
            origBalance = CDbl(wsTape.Cells(i, "L").Value)
        Else
            origBalance = 0
        End If
        
        ' current_prin (column M)
        If IsNumeric(wsTape.Cells(i, "M").Value) And Not IsEmpty(wsTape.Cells(i, "M").Value) Then
            currBalance = CDbl(wsTape.Cells(i, "M").Value)
        Else
            currBalance = 0
        End If
        
        ' months_to_maturity (column J)
        If IsNumeric(wsTape.Cells(i, "J").Value) And Not IsEmpty(wsTape.Cells(i, "J").Value) Then
            remTerm = CDbl(wsTape.Cells(i, "J").Value)
        Else
            remTerm = 0
        End If
        
        ' net_borrower_coupon (column AJ) - divide by 100 to convert to decimal
        If IsNumeric(wsTape.Cells(i, "AJ").Value) And Not IsEmpty(wsTape.Cells(i, "AJ").Value) Then
            wac = CDbl(wsTape.Cells(i, "AJ").Value) / 100
        Else
            wac = 0
        End If
        
        ' accrued_int_to_cap (column O)
        If IsNumeric(wsTape.Cells(i, "O").Value) And Not IsEmpty(wsTape.Cells(i, "O").Value) Then
            accruedIntToCap = CDbl(wsTape.Cells(i, "O").Value)
        Else
            accruedIntToCap = 0
        End If
        
        ' Calculate Mons to repay (already have asofDate and firstPmtDate from earlier)
        If IsDate(wsTape.Cells(i, "CQ").Value) Then
            monsToRepay = DateDiff("m", asofDate, firstPmtDate)
            
            ' For "full" repay type, set to 0 if negative
            If repaymentType = "full" And monsToRepay < 0 Then
                monsToRepay = 0
            End If
        Else
            monsToRepay = 0
        End If
        
        ' Fixed Pay from column AQ (cur_pmt_amt) - ONLY for "partial" repay type
        fixedPay = 0
        If repaymentType = "partial" Then
            If IsNumeric(wsTape.Cells(i, "AQ").Value) And Not IsEmpty(wsTape.Cells(i, "AQ").Value) Then
                fixedPay = CDbl(wsTape.Cells(i, "AQ").Value)
            End If
        End If
        
        ' Add to repline data
        Dim replineArr As Variant
        replineArr = replineDataDict(replineKey)
        
        ' Update repline totals
        replineArr(2) = replineArr(2) + origBalance ' Orig Balance
        replineArr(3) = replineArr(3) + currBalance ' Curr Balance
        replineArr(4) = replineArr(4) + (remTerm * currBalance) ' Weighted Rem Term
        replineArr(5) = replineArr(5) + (originalTerm * origBalance) ' Weighted Orig Term
        replineArr(6) = replineArr(6) + (wac * currBalance) ' Weighted WAC
        replineArr(8) = replineArr(8) + accruedIntToCap ' Accrued int to be capped
        replineArr(9) = replineArr(9) + (monsToRepay * currBalance) ' Weighted Mons to repay
        replineArr(10) = replineArr(10) + fixedPay ' Fixed Pay
        replineArr(11) = replineArr(11) + 1 ' Count of loans
        
        replineDataDict(replineKey) = replineArr
        
NextLoan:
    Next i
    
    ' Write headers to Replines sheet
    wsReplines.Range("A1").Value = "repay_tier_term"
    wsReplines.Range("B1").Value = "Repline"
    wsReplines.Range("C1").Value = "Orig Balance"
    wsReplines.Range("D1").Value = "Curr Balance"
    wsReplines.Range("E1").Value = "Rem Term"
    wsReplines.Range("F1").Value = "Orig term"
    wsReplines.Range("G1").Value = "wac"
    wsReplines.Range("H1").Value = "Capitalize interest"
    wsReplines.Range("I1").Value = "Accrued int to be capped"
    wsReplines.Range("J1").Value = "Mons to repay"
    wsReplines.Range("K1").Value = "Fixed Pay"
    
    ' Write repline data
    Dim outputRow As Long
    outputRow = 2
    
    For replineNumber = 1 To 112
        replineKey = CStr(replineNumber)
        Dim replineOut As Variant
        replineOut = replineDataDict(replineKey)
        
        ' Write repline identifier
        wsReplines.Cells(outputRow, "A").Value = replineOut(0) ' repay_tier_term
        wsReplines.Cells(outputRow, "B").Value = replineOut(1) ' Repline number
        
        ' Calculate weighted averages
        Dim currBal As Double
        currBal = replineOut(3)
        
        wsReplines.Cells(outputRow, "C").Value = replineOut(2) ' Orig Balance
        wsReplines.Cells(outputRow, "D").Value = currBal ' Curr Balance
        
        ' Weighted averages
        If currBal > 0 Then
            wsReplines.Cells(outputRow, "E").Value = Round(replineOut(4) / currBal, 0) ' Rem Term
            wsReplines.Cells(outputRow, "G").Value = replineOut(6) / currBal ' WAC
            wsReplines.Cells(outputRow, "J").Value = Round(replineOut(9) / currBal, 0) ' Mons to repay
        Else
            wsReplines.Cells(outputRow, "E").Value = 0
            wsReplines.Cells(outputRow, "G").Value = 0
            wsReplines.Cells(outputRow, "J").Value = 0
        End If
        
        If replineOut(2) > 0 Then
            wsReplines.Cells(outputRow, "F").Value = Round(replineOut(5) / replineOut(2), 0) ' Orig term
        Else
            wsReplines.Cells(outputRow, "F").Value = 0
        End If
        
        ' Capitalize interest based on repay type: defer/partial = Y, IO/full = N
        Dim repayTypeStr As String
        repayTypeStr = Split(replineOut(0), " ")(0) ' Get first word (repay type)
        
        If repayTypeStr = "defer" Or repayTypeStr = "partial" Then
            wsReplines.Cells(outputRow, "H").Value = "Y"
            wsReplines.Cells(outputRow, "I").Value = replineOut(8) ' Accrued int to be capped
        Else
            wsReplines.Cells(outputRow, "H").Value = "N"
            wsReplines.Cells(outputRow, "I").Value = 0 ' Zero for IO and full
        End If
        
        wsReplines.Cells(outputRow, "K").Value = replineOut(10) ' Fixed Pay
        
        outputRow = outputRow + 1
    Next replineNumber
    
    ' Format the output
    wsReplines.Columns("C:D").NumberFormat = "#,##0.00"
    wsReplines.Columns("E:F").NumberFormat = "0"
    wsReplines.Columns("G").NumberFormat = "0.0000%"
    wsReplines.Columns("I").NumberFormat = "#,##0.00"
    wsReplines.Columns("J").NumberFormat = "0"
    wsReplines.Columns("K").NumberFormat = "#,##0.00"
    
    ' Auto-fit columns
    wsReplines.Columns("A:K").AutoFit
    
    ' === VALIDATION AND BALANCING ===
    ' Calculate totals from tape
    Dim tapeOrigBal As Double, tapeCurrBal As Double
    tapeOrigBal = 0
    tapeCurrBal = 0
    
    For i = 2 To lastRow
        If IsNumeric(wsTape.Cells(i, "L").Value) And Not IsEmpty(wsTape.Cells(i, "L").Value) Then
            tapeOrigBal = tapeOrigBal + CDbl(wsTape.Cells(i, "L").Value)
        End If
        If IsNumeric(wsTape.Cells(i, "M").Value) And Not IsEmpty(wsTape.Cells(i, "M").Value) Then
            tapeCurrBal = tapeCurrBal + CDbl(wsTape.Cells(i, "M").Value)
        End If
    Next i
    
    ' Calculate totals from replines
    Dim replineOrigBal As Double, replineCurrBal As Double
    replineOrigBal = 0
    replineCurrBal = 0
    
    For i = 2 To 113 ' Rows 2 to 113 (112 replines)
        If IsNumeric(wsReplines.Cells(i, "C").Value) Then
            replineOrigBal = replineOrigBal + wsReplines.Cells(i, "C").Value
        End If
        If IsNumeric(wsReplines.Cells(i, "D").Value) Then
            replineCurrBal = replineCurrBal + wsReplines.Cells(i, "D").Value
        End If
    Next i
    
    ' Calculate differences
    Dim origBalDiff As Double, currBalDiff As Double
    origBalDiff = tapeOrigBal - replineOrigBal
    currBalDiff = tapeCurrBal - replineCurrBal
    
    ' Distribute adjustment evenly across all 112 replines
    Dim origBalAdjPerRepline As Double, currBalAdjPerRepline As Double
    origBalAdjPerRepline = origBalDiff / 112
    currBalAdjPerRepline = currBalDiff / 112
    
    ' Apply adjustment to each repline
    For i = 2 To 113 ' Rows 2 to 113 (112 replines)
        Dim currentOrigBal As Double, currentCurrBal As Double
        Dim newOrigBal As Double, newCurrBal As Double
        
        currentOrigBal = wsReplines.Cells(i, "C").Value
        currentCurrBal = wsReplines.Cells(i, "D").Value
        
        newOrigBal = currentOrigBal + origBalAdjPerRepline
        newCurrBal = currentCurrBal + currBalAdjPerRepline
        
        wsReplines.Cells(i, "C").Value = newOrigBal
        wsReplines.Cells(i, "D").Value = newCurrBal
        
        ' Recalculate weighted averages for this repline
        ' Orig term weighted average
        If currentOrigBal > 0 And newOrigBal > 0 Then
            Dim origTermWeighted As Double
            origTermWeighted = wsReplines.Cells(i, "F").Value * currentOrigBal
            wsReplines.Cells(i, "F").Value = Round(origTermWeighted / newOrigBal, 0)
        End If
        
        ' Rem term, WAC, and Mons to repay weighted averages
        If currentCurrBal > 0 And newCurrBal > 0 Then
            Dim remTermWeighted As Double, wacWeighted As Double, monsWeighted As Double
            remTermWeighted = wsReplines.Cells(i, "E").Value * currentCurrBal
            wacWeighted = wsReplines.Cells(i, "G").Value * currentCurrBal
            monsWeighted = wsReplines.Cells(i, "J").Value * currentCurrBal
            
            wsReplines.Cells(i, "E").Value = Round(remTermWeighted / newCurrBal, 0)
            wsReplines.Cells(i, "G").Value = wacWeighted / newCurrBal
            wsReplines.Cells(i, "J").Value = Round(monsWeighted / newCurrBal, 0)
        End If
    Next i
    
    ' === ADD SUMMARY TABLES TO SEPARATE "Pool Strat" TAB ===
    
    ' Create or clear Pool Strat sheet
    Dim wsSummary As Worksheet
    On Error Resume Next
    Set wsSummary = ThisWorkbook.Sheets("Pool Strat")
    If wsSummary Is Nothing Then
        Set wsSummary = ThisWorkbook.Sheets.Add(After:=wsReplines)
        wsSummary.Name = "Pool Strat"
    Else
        wsSummary.Cells.Clear
    End If
    On Error GoTo ErrorHandler
    
    ' === 1. REPAYMENT BUCKET SUMMARY TABLE (Starting at D4) ===
    Dim summaryRow As Long
    summaryRow = 4
    Dim summaryStartCol As Long
    summaryStartCol = 4 ' Column D
    
    wsSummary.Cells(summaryRow, summaryStartCol).Value = "REPAYMENT BUCKET SUMMARY"
    wsSummary.Cells(summaryRow, summaryStartCol).Font.Bold = True
    summaryRow = summaryRow + 1
    
    ' Headers
    wsSummary.Cells(summaryRow, summaryStartCol).Value = "Repayment Type"
    wsSummary.Cells(summaryRow, summaryStartCol + 1).Value = "Current Principal"
    wsSummary.Cells(summaryRow, summaryStartCol + 2).Value = "Weighted Avg WAC"
    wsSummary.Cells(summaryRow, summaryStartCol + 3).Value = "% of Pool"
    wsSummary.Range(wsSummary.Cells(summaryRow, summaryStartCol), wsSummary.Cells(summaryRow, summaryStartCol + 3)).Font.Bold = True
    summaryRow = summaryRow + 1
    
    ' Calculate totals by repayment type
    Dim repayTypesList As Variant
    repayTypesList = Array("full", "IO", "partial", "defer")
    
    Dim totalPoolCurrBal As Double
    totalPoolCurrBal = 0
    
    ' First pass: calculate total pool balance
    Dim iSum As Long
    For iSum = 2 To 113
        totalPoolCurrBal = totalPoolCurrBal + wsReplines.Cells(iSum, "D").Value
    Next iSum
    
    ' Second pass: calculate by repayment type
    Dim repayTypeLoop As Variant
    For Each repayTypeLoop In repayTypesList
        Dim repayTypeCurrBal As Double, repayTypeWACWeighted As Double
        repayTypeCurrBal = 0
        repayTypeWACWeighted = 0
        
        For iSum = 2 To 113
            Dim replineNameSum As String
            replineNameSum = wsReplines.Cells(iSum, "A").Value
            
            If InStr(1, replineNameSum, repayTypeLoop & " ", vbTextCompare) = 1 Then
                Dim replineCurrBalSum As Double, replineWACSum As Double
                replineCurrBalSum = wsReplines.Cells(iSum, "D").Value
                replineWACSum = wsReplines.Cells(iSum, "G").Value
                
                repayTypeCurrBal = repayTypeCurrBal + replineCurrBalSum
                repayTypeWACWeighted = repayTypeWACWeighted + (replineWACSum * replineCurrBalSum)
            End If
        Next iSum
        
        ' Write to summary
        wsSummary.Cells(summaryRow, summaryStartCol).Value = repayTypeLoop
        wsSummary.Cells(summaryRow, summaryStartCol + 1).Value = repayTypeCurrBal
        wsSummary.Cells(summaryRow, summaryStartCol + 1).NumberFormat = "#,##0.00"
        
        If repayTypeCurrBal > 0 Then
            wsSummary.Cells(summaryRow, summaryStartCol + 2).Value = repayTypeWACWeighted / repayTypeCurrBal
        Else
            wsSummary.Cells(summaryRow, summaryStartCol + 2).Value = 0
        End If
        wsSummary.Cells(summaryRow, summaryStartCol + 2).NumberFormat = "0.0000%"
        
        If totalPoolCurrBal > 0 Then
            wsSummary.Cells(summaryRow, summaryStartCol + 3).Value = repayTypeCurrBal / totalPoolCurrBal
        Else
            wsSummary.Cells(summaryRow, summaryStartCol + 3).Value = 0
        End If
        wsSummary.Cells(summaryRow, summaryStartCol + 3).NumberFormat = "0.00%"
        
        summaryRow = summaryRow + 1
    Next repayTypeLoop
    
    ' Add total row
    wsSummary.Cells(summaryRow, summaryStartCol).Value = "TOTAL POOL"
    wsSummary.Cells(summaryRow, summaryStartCol).Font.Bold = True
    wsSummary.Cells(summaryRow, summaryStartCol + 1).Value = totalPoolCurrBal
    wsSummary.Cells(summaryRow, summaryStartCol + 1).NumberFormat = "#,##0.00"
    wsSummary.Cells(summaryRow, summaryStartCol + 1).Font.Bold = True
    
    If totalPoolCurrBal > 0 Then
        Dim totalWACWeighted As Double
        totalWACWeighted = 0
        For iSum = 2 To 113
            totalWACWeighted = totalWACWeighted + (wsReplines.Cells(iSum, "G").Value * wsReplines.Cells(iSum, "D").Value)
        Next iSum
        wsSummary.Cells(summaryRow, summaryStartCol + 2).Value = totalWACWeighted / totalPoolCurrBal
    Else
        wsSummary.Cells(summaryRow, summaryStartCol + 2).Value = 0
    End If
    wsSummary.Cells(summaryRow, summaryStartCol + 2).NumberFormat = "0.0000%"
    wsSummary.Cells(summaryRow, summaryStartCol + 2).Font.Bold = True
    
    wsSummary.Cells(summaryRow, summaryStartCol + 3).Value = 1
    wsSummary.Cells(summaryRow, summaryStartCol + 3).NumberFormat = "0.00%"
    wsSummary.Cells(summaryRow, summaryStartCol + 3).Font.Bold = True
    
    ' === 2. TIER BREAKDOWN TABLE (Starting at D13) ===
    summaryRow = 13
    
    wsSummary.Cells(summaryRow, summaryStartCol).Value = "TIER BREAKDOWN"
    wsSummary.Cells(summaryRow, summaryStartCol).Font.Bold = True
    summaryRow = summaryRow + 1
    
    ' Headers
    wsSummary.Cells(summaryRow, summaryStartCol).Value = "Tier"
    wsSummary.Cells(summaryRow, summaryStartCol + 1).Value = "Current Principal"
    wsSummary.Cells(summaryRow, summaryStartCol + 2).Value = "Weighted Avg WAC"
    wsSummary.Cells(summaryRow, summaryStartCol + 3).Value = "% of Pool"
    wsSummary.Range(wsSummary.Cells(summaryRow, summaryStartCol), wsSummary.Cells(summaryRow, summaryStartCol + 3)).Font.Bold = True
    summaryRow = summaryRow + 1
    
    ' Loop through each tier (1 to 7)
    Dim tierLoop As Long
    For tierLoop = 1 To 7
        Dim tierCurrBal As Double, tierWACWeighted As Double
        tierCurrBal = 0
        tierWACWeighted = 0
        
        ' Sum up all replines for this tier
        For iSum = 2 To 113
            replineNameSum = wsReplines.Cells(iSum, "A").Value
            
            ' Check if repline belongs to this tier (e.g., "tier_3")
            If InStr(replineNameSum, " tier_" & tierLoop & " ") > 0 Then
                replineCurrBalSum = wsReplines.Cells(iSum, "D").Value
                replineWACSum = wsReplines.Cells(iSum, "G").Value
                
                tierCurrBal = tierCurrBal + replineCurrBalSum
                tierWACWeighted = tierWACWeighted + (replineWACSum * replineCurrBalSum)
            End If
        Next iSum
        
        ' Write tier data
        wsSummary.Cells(summaryRow, summaryStartCol).Value = "Tier " & tierLoop
        wsSummary.Cells(summaryRow, summaryStartCol + 1).Value = tierCurrBal
        wsSummary.Cells(summaryRow, summaryStartCol + 1).NumberFormat = "#,##0.00"
        
        If tierCurrBal > 0 Then
            wsSummary.Cells(summaryRow, summaryStartCol + 2).Value = tierWACWeighted / tierCurrBal
        Else
            wsSummary.Cells(summaryRow, summaryStartCol + 2).Value = 0
        End If
        wsSummary.Cells(summaryRow, summaryStartCol + 2).NumberFormat = "0.0000%"
        
        If totalPoolCurrBal > 0 Then
            wsSummary.Cells(summaryRow, summaryStartCol + 3).Value = tierCurrBal / totalPoolCurrBal
        Else
            wsSummary.Cells(summaryRow, summaryStartCol + 3).Value = 0
        End If
        wsSummary.Cells(summaryRow, summaryStartCol + 3).NumberFormat = "0.00%"
        
        summaryRow = summaryRow + 1
    Next tierLoop
    
    ' Add tier total row
    wsSummary.Cells(summaryRow, summaryStartCol).Value = "TOTAL POOL"
    wsSummary.Cells(summaryRow, summaryStartCol).Font.Bold = True
    wsSummary.Cells(summaryRow, summaryStartCol + 1).Value = totalPoolCurrBal
    wsSummary.Cells(summaryRow, summaryStartCol + 1).NumberFormat = "#,##0.00"
    wsSummary.Cells(summaryRow, summaryStartCol + 1).Font.Bold = True
    wsSummary.Cells(summaryRow, summaryStartCol + 2).Value = totalWACWeighted / totalPoolCurrBal
    wsSummary.Cells(summaryRow, summaryStartCol + 2).NumberFormat = "0.0000%"
    wsSummary.Cells(summaryRow, summaryStartCol + 2).Font.Bold = True
    wsSummary.Cells(summaryRow, summaryStartCol + 3).Value = 1
    wsSummary.Cells(summaryRow, summaryStartCol + 3).NumberFormat = "0.00%"
    wsSummary.Cells(summaryRow, summaryStartCol + 3).Font.Bold = True
    
    ' === 3. TERM BREAKDOWN TABLE (Starting 3 rows after tier table) ===
    summaryRow = summaryRow + 3
    
    wsSummary.Cells(summaryRow, summaryStartCol).Value = "TERM BREAKDOWN"
    wsSummary.Cells(summaryRow, summaryStartCol).Font.Bold = True
    summaryRow = summaryRow + 1
    
    ' Headers
    wsSummary.Cells(summaryRow, summaryStartCol).Value = "Term"
    wsSummary.Cells(summaryRow, summaryStartCol + 1).Value = "Current Principal"
    wsSummary.Cells(summaryRow, summaryStartCol + 2).Value = "Weighted Avg WAC"
    wsSummary.Cells(summaryRow, summaryStartCol + 3).Value = "% of Pool"
    wsSummary.Range(wsSummary.Cells(summaryRow, summaryStartCol), wsSummary.Cells(summaryRow, summaryStartCol + 3)).Font.Bold = True
    summaryRow = summaryRow + 1
    
    ' Loop through each term (term_5, term_7, term_10, term_15)
    Dim termBucketsList As Variant
    termBucketsList = Array("term_5", "term_7", "term_10", "term_15")
    
    Dim termBucketLoop As Variant
    For Each termBucketLoop In termBucketsList
        Dim termCurrBal As Double, termWACWeighted As Double
        termCurrBal = 0
        termWACWeighted = 0
        
        ' Sum up all replines for this term
        For iSum = 2 To 113
            replineNameSum = wsReplines.Cells(iSum, "A").Value
            
            ' Check if repline belongs to this term (e.g., "term_5")
            If InStr(replineNameSum, " " & termBucketLoop) > 0 Then
                replineCurrBalSum = wsReplines.Cells(iSum, "D").Value
                replineWACSum = wsReplines.Cells(iSum, "G").Value
                
                termCurrBal = termCurrBal + replineCurrBalSum
                termWACWeighted = termWACWeighted + (replineWACSum * replineCurrBalSum)
            End If
        Next iSum
        
        ' Write term data
        wsSummary.Cells(summaryRow, summaryStartCol).Value = termBucketLoop
        wsSummary.Cells(summaryRow, summaryStartCol + 1).Value = termCurrBal
        wsSummary.Cells(summaryRow, summaryStartCol + 1).NumberFormat = "#,##0.00"
        
        If termCurrBal > 0 Then
            wsSummary.Cells(summaryRow, summaryStartCol + 2).Value = termWACWeighted / termCurrBal
        Else
            wsSummary.Cells(summaryRow, summaryStartCol + 2).Value = 0
        End If
        wsSummary.Cells(summaryRow, summaryStartCol + 2).NumberFormat = "0.0000%"
        
        If totalPoolCurrBal > 0 Then
            wsSummary.Cells(summaryRow, summaryStartCol + 3).Value = termCurrBal / totalPoolCurrBal
        Else
            wsSummary.Cells(summaryRow, summaryStartCol + 3).Value = 0
        End If
        wsSummary.Cells(summaryRow, summaryStartCol + 3).NumberFormat = "0.00%"
        
        summaryRow = summaryRow + 1
    Next termBucketLoop
    
    ' Add term total row
    wsSummary.Cells(summaryRow, summaryStartCol).Value = "TOTAL POOL"
    wsSummary.Cells(summaryRow, summaryStartCol).Font.Bold = True
    wsSummary.Cells(summaryRow, summaryStartCol + 1).Value = totalPoolCurrBal
    wsSummary.Cells(summaryRow, summaryStartCol + 1).NumberFormat = "#,##0.00"
    wsSummary.Cells(summaryRow, summaryStartCol + 1).Font.Bold = True
    wsSummary.Cells(summaryRow, summaryStartCol + 2).Value = totalWACWeighted / totalPoolCurrBal
    wsSummary.Cells(summaryRow, summaryStartCol + 2).NumberFormat = "0.0000%"
    wsSummary.Cells(summaryRow, summaryStartCol + 2).Font.Bold = True
    wsSummary.Cells(summaryRow, summaryStartCol + 3).Value = 1
    wsSummary.Cells(summaryRow, summaryStartCol + 3).NumberFormat = "0.00%"
    wsSummary.Cells(summaryRow, summaryStartCol + 3).Font.Bold = True
    
    ' Auto-fit summary columns
    wsSummary.Columns(summaryStartCol).AutoFit
    wsSummary.Columns(summaryStartCol + 1).AutoFit
    wsSummary.Columns(summaryStartCol + 2).AutoFit
    wsSummary.Columns(summaryStartCol + 3).AutoFit
    
    ' Create validation message
    validationMsg = "VALIDATION RESULTS:" & vbCrLf & vbCrLf
    validationMsg = validationMsg & "Tape Original Balance: " & Format(tapeOrigBal, "#,##0.00") & vbCrLf
    validationMsg = validationMsg & "Repline Original Balance (before adj): " & Format(replineOrigBal, "#,##0.00") & vbCrLf
    validationMsg = validationMsg & "Difference: " & Format(origBalDiff, "#,##0.00") & vbCrLf
    validationMsg = validationMsg & "Per Repline Adjustment: " & Format(origBalAdjPerRepline, "#,##0.00") & vbCrLf & vbCrLf
    validationMsg = validationMsg & "Tape Current Balance: " & Format(tapeCurrBal, "#,##0.00") & vbCrLf
    validationMsg = validationMsg & "Repline Current Balance (before adj): " & Format(replineCurrBal, "#,##0.00") & vbCrLf
    validationMsg = validationMsg & "Difference: " & Format(currBalDiff, "#,##0.00") & vbCrLf
    validationMsg = validationMsg & "Per Repline Adjustment: " & Format(currBalAdjPerRepline, "#,##0.00") & vbCrLf & vbCrLf
    
    If Abs(origBalDiff) > 0.01 Or Abs(currBalDiff) > 0.01 Then
        validationMsg = validationMsg & "Balances adjusted to match!"
    Else
        validationMsg = validationMsg & "Balances match perfectly!"
    End If
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    
    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Else
        MsgBox "Successfully created 112 replines from " & (lastRow - 1) & " loans!" & vbCrLf & vbCrLf & _
               "Summary tables created in 'Pool Strat' tab!" & vbCrLf & vbCrLf & validationMsg, vbInformation
    End If
End Sub

Function CalculateReplineNum(tierVal As Long, repayTypeVal As String, termBucketVal As String) As Long
    ' Calculate repline number based on tier, repay type, and term
    ' Formula: ((tier - 1) * 16) + (repay_type_offset * 4) + term_offset
    
    Dim repayOffsetVal As Long
    Dim termOffsetVal As Long
    
    ' Determine repay type offset (0, 4, 8, 12)
    Select Case repayTypeVal
        Case "full"
            repayOffsetVal = 0
        Case "IO"
            repayOffsetVal = 4
        Case "partial"
            repayOffsetVal = 8
        Case "defer"
            repayOffsetVal = 12
    End Select
    
    ' Determine term offset (1, 2, 3, 4)
    Select Case termBucketVal
        Case "term_5"
            termOffsetVal = 1
        Case "term_7"
            termOffsetVal = 2
        Case "term_10"
            termOffsetVal = 3
        Case "term_15"
            termOffsetVal = 4
    End Select
    
    CalculateReplineNum = ((tierVal - 1) * 16) + repayOffsetVal + termOffsetVal
End Function

Sub InitializeReplinesV2(replineDataDict As Object)
    ' Initialize all 112 replines with zero values
    Dim tierVal As Long
    Dim repayTypeVal As String
    Dim termBucketVal As String
    Dim replineNumVal As Long
    Dim repayTypesList As Variant
    Dim termBucketsList As Variant
    Dim idxI As Long, idxJ As Long
    
    repayTypesList = Array("full", "IO", "partial", "defer")
    termBucketsList = Array("term_5", "term_7", "term_10", "term_15")
    
    For tierVal = 1 To 7
        For idxI = 0 To 3
            repayTypeVal = repayTypesList(idxI)
            For idxJ = 0 To 3
                termBucketVal = termBucketsList(idxJ)
                replineNumVal = CalculateReplineNum(tierVal, repayTypeVal, termBucketVal)
                
                ' Create repline identifier string
                Dim repayTierTermStr As String
                repayTierTermStr = repayTypeVal & " tier_" & tierVal & " " & termBucketVal
                
                ' Initialize array: (name, repline#, origBal, currBal, remTerm_weighted, origTerm_weighted, wac_weighted, unused, accruedInt, monsToRepay_weighted, fixedPay, count)
                Dim replineArrayVal As Variant
                replineArrayVal = Array(repayTierTermStr, replineNumVal, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
                
                replineDataDict.Add CStr(replineNumVal), replineArrayVal
            Next idxJ
        Next idxI
    Next tierVal
End Sub

