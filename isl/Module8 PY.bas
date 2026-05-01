Attribute VB_Name = "Module8"
Sub Calculate_Pool_PV_From_Yield_v2()

    ' OUTPUT: E13 = Calculated PV, E14 = Pool Price (E13 / E1)
    ' This does NOT use Goal Seek - just calculates PV directly
    
    Dim wsAssump As Worksheet
    Dim wsPool As Worksheet
    Dim pool_yield_bey As Double
    Dim pool_yield_monthly As Double
    Dim lastRow As Long, i As Long
    Dim poolPV As Double
    
    ' Arrays for fast processing
    Dim monthArr As Variant, cfArr As Variant, pvArr() As Variant
    Dim arrSize As Long
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual  ' <-- CRITICAL: Turn off auto-calc
    
    On Error GoTo ErrorHandler
    
    Set wsAssump = ThisWorkbook.Sheets("Assumption")
    Set wsPool = ThisWorkbook.Sheets("Pool CF")
    
    If Not IsNumeric(wsAssump.Range("E12").Value) Then
        MsgBox "Please enter a valid BEY yield in cell E12.", vbExclamation
        GoTo ErrorHandler
    End If
    
    ' Convert BEY (E12) to monthly yield
    pool_yield_bey = wsAssump.Range("E12").Value
    pool_yield_monthly = 12 * ((1 + pool_yield_bey / 2) ^ (1 / 6) - 1)
    
    ' Store monthly yield in Pool CF C8
    wsPool.Range("C8").Value = pool_yield_monthly
    wsPool.Range("C8").NumberFormat = "0.0000%"
    
    ' Find last row with cash flows
    lastRow = wsPool.Cells(wsPool.Rows.Count, "O").End(xlUp).row
    
    ' Read columns into arrays (FAST - single read operation)
    monthArr = wsPool.Range("B12:B" & lastRow).Value
    cfArr = wsPool.Range("O12:O" & lastRow).Value
    arrSize = UBound(monthArr, 1)
    
    ' Size output array
    ReDim pvArr(1 To arrSize, 1 To 1)
    
    ' Calculate PV in memory (FAST - no Excel interaction)
    For i = 1 To arrSize
        If IsNumeric(monthArr(i, 1)) And IsNumeric(cfArr(i, 1)) Then
            If monthArr(i, 1) <> 0 Or cfArr(i, 1) <> 0 Then
                pvArr(i, 1) = cfArr(i, 1) / ((1 + pool_yield_monthly) ^ (monthArr(i, 1) / 12))
            Else
                pvArr(i, 1) = ""
            End If
        Else
            pvArr(i, 1) = ""
        End If
    Next i
    
    ' Write entire array to column P at once (FAST - single write operation)
    wsPool.Range("P12:P" & lastRow).Value = pvArr
    wsPool.Range("P12:P" & lastRow).NumberFormat = "#,##0.00"
    
    ' Calculate total PV
    poolPV = Application.WorksheetFunction.Sum(wsPool.Range("P12:P" & lastRow))
    
    ' Output PV to Assumption E13
    wsAssump.Range("E13").Value = poolPV
    wsAssump.Range("E13").NumberFormat = "#,##0"
    
 
    
    ' Update Pool CF sheet
    wsPool.Range("F3").Value = poolPV
    wsPool.Range("F3").NumberFormat = "#,##0"
    
    If wsPool.Range("C1").Value <> 0 Then
        wsPool.Range("F4").Value = poolPV / wsPool.Range("C1").Value
        wsPool.Range("F4").NumberFormat = "0.0000%"
    Else
        wsPool.Range("F4").Value = ""
    End If
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic  ' <-- Turn calc back on
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Else
        MsgBox "PV Calculation Complete!" & vbCrLf & vbCrLf & _
               "Input BEY Yield (E12): " & Format(pool_yield_bey, "0.00%") & vbCrLf & _
               "Monthly Yield: " & Format(pool_yield_monthly, "0.0000%") & vbCrLf & _
               "Calculated PV (E13): " & Format(poolPV, "#,##0") & vbCrLf & _
               ", vbInformation"
    End If
End Sub
Sub GoalSeek_Yield_From_Target_PV()
    ' INPUT: E2 = Target PV
    ' GOAL SEEK: Find monthly discount rate in C8 that makes PV of column O = E2
    ' OUTPUT: E4 = BEY yield
    ' Column P already exists with values - DO NOT touch it!
    
    Dim wsAssump As Worksheet
    Dim wsPool As Worksheet
    Dim targetPV As Double
    Dim monthly_rate As Double
    Dim bey_yield As Double
    Dim lastRow As Long, i As Long
    Dim period As Double, cashflow As Double, pv As Double
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    
    On Error GoTo ErrorHandler
    
    Set wsAssump = ThisWorkbook.Sheets("Assumption")
    Set wsPool = ThisWorkbook.Sheets("Pool CF")
    
    targetPV = wsAssump.Range("E2").Value
    
    ' Initial guess for monthly rate in C8
    wsPool.Range("C8").Value = 0.004
    
    lastRow = wsPool.Cells(wsPool.Rows.Count, "O").End(xlUp).row
    
    ' Put SUMPRODUCT formula in F3 that references column O and C8
    wsPool.Range("F3").Formula = "=SUMPRODUCT(O12:O" & lastRow & "/(1+C8)^(B12:B" & lastRow & "/12))"
    
    Application.Calculate
    
    ' GOAL SEEK: Change C8 until F3 = E2
    wsPool.Range("F3").GoalSeek Goal:=targetPV, ChangingCell:=wsPool.Range("C8")
    
    ' Read solved monthly rate
    monthly_rate = wsPool.Range("C8").Value
    
    ' Convert monthly to BEY
    bey_yield = 2 * ((1 + monthly_rate / 12) ^ 6 - 1)
    
    ' Output to E4
    wsAssump.Range("E4").Value = bey_yield
    wsAssump.Range("E4").NumberFormat = "0.00%"
    
    ' Clean up F3
    wsPool.Range("F3").ClearContents
    
ErrorHandler:
    Application.ScreenUpdating = True
    
    MsgBox "BEY Yield: " & Format(bey_yield, "0.00%"), vbInformation
End Sub


