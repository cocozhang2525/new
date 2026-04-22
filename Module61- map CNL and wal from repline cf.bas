Attribute VB_Name = "Module61"
Sub Export_CNL_And_Calculate_WAL()
    Dim wsAssump As Worksheet
    Dim ws As Worksheet
    Dim wsRepline As Worksheet
    Dim replineSheets As Collection
    Dim replineNum As Variant
    Dim replineRow As Long
    Dim totalWeightedPeriod As Double, totalPrincipal As Double
    Dim period As Double, princ As Double
    Dim wal As Double
    Dim lastRow As Long
    Dim i As Long
    Dim balPrev As Variant, balCur As Variant
    
    ' Arrays for batch output
    Dim cnlOutputArr() As Variant
    Dim walOutputArr() As Variant
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Set wsAssump = ThisWorkbook.Sheets("Assumption")
    Set replineSheets = New Collection
    
    ' Initialize output arrays
    ReDim cnlOutputArr(39 To 340, 1 To 1)
    ReDim walOutputArr(39 To 340, 1 To 1)
    
    ' Collect all Repline CF sheets
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "Repline * CF" Then
            replineSheets.Add ws
        End If
    Next ws
    
    ' Main loop: Export CNL and Calculate WAL for each repline
    For replineRow = 39 To 340
        replineNum = wsAssump.Cells(replineRow, "C").Value
        
        If IsNumeric(replineNum) And replineNum >= 1 And replineNum <= 299 Then
            Set wsRepline = Nothing
            On Error Resume Next
            Set wsRepline = ThisWorkbook.Sheets("Repline " & replineNum & " CF")
            On Error GoTo 0
            
            If Not wsRepline Is Nothing Then
                ' --- EXPORT CNL (I1) ---
                cnlOutputArr(replineRow, 1) = wsRepline.Range("I1").Value
                
                ' --- CALCULATE WAL ---
                lastRow = wsRepline.Cells(wsRepline.Rows.Count, "B").End(xlUp).Row
                totalWeightedPeriod = 0
                totalPrincipal = 0
                
                ' Loop from row 12, using D(i-1) - D(i) as the principal for row i
                For i = 12 To lastRow
                    princ = 0
                    balPrev = wsRepline.Cells(i - 1, "D").Value
                    balCur = wsRepline.Cells(i, "D").Value
                    
                    If IsNumeric(balPrev) And IsNumeric(balCur) Then
                        ' Decrease in balance from previous month
                        princ = balPrev - balCur
                    Else
                        princ = 0
                    End If
                    
                    ' Period from column B
                    If IsNumeric(wsRepline.Cells(i, "B").Value) Then
                        period = wsRepline.Cells(i, "B").Value
                    Else
                        period = 0
                    End If
                    
                    ' Use only rows with numeric period and positive principal
                    If period > 0 And princ <> 0 Then
                        totalWeightedPeriod = totalWeightedPeriod + (princ * period)
                        totalPrincipal = totalPrincipal + princ
                    End If
                Next i
                
                If totalPrincipal <> 0 Then
                    ' Period is in months, so divide by 12 to get WAL in years
                    wal = Round((totalWeightedPeriod / totalPrincipal) / 12, 3)
                    wsRepline.Range("F5").Value = wal
                    wsRepline.Range("F5").NumberFormat = "0.000"
                    wsRepline.Range("E5").Value = "WAL"
                    walOutputArr(replineRow, 1) = wal
                Else
                    wsRepline.Range("F5").Value = "N/A"
                    wsRepline.Range("E5").Value = "WAL"
                    walOutputArr(replineRow, 1) = "N/A"
                End If
            End If
        End If
    Next replineRow
    
    ' Batch write outputs to Assumption sheet
    wsAssump.Range("Q39:Q340").Value = cnlOutputArr
    wsAssump.Range("Q39:Q340").NumberFormat = "0.00%"
    
    wsAssump.Range("T39:T340").Value = walOutputArr
    ' Apply conditional formatting for WAL
    For replineRow = 39 To 340
        If walOutputArr(replineRow, 1) = "N/A" Then
            wsAssump.Cells(replineRow, "T").NumberFormat = "General"
        Else
            wsAssump.Cells(replineRow, "T").NumberFormat = "0.000"
        End If
    Next replineRow
    
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    MsgBox "Complete!" & vbCrLf & _
           "CNL values exported to Assumption column Q." & vbCrLf & _
           "WAL values exported to Assumption column T." & vbCrLf & _
           "Processed " & replineSheets.Count & " repline sheets.", vbInformation
End Sub

