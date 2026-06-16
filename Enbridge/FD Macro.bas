Attribute VB_Name = "Module1"
Sub DistributeDataToMonths()
    Dim wsSource As Worksheet, wsTarget As Worksheet, wsList As Worksheet
    Dim lastRow As Long, r As Long, nextRow As Long, targetLastRow As Long
    Dim colDate As String, trdType As String, bkStrat As String
    Dim dateParts() As String, startStr As String, endStr As String
    Dim m1_str As String, m2_str As String, y1_str As String, y2_str As String
    Dim m1_idx As Integer, m2_idx As Integer, StartYear As Integer, EndYear As Integer
    Dim StartAbs As Long, EndAbs As Long, i As Integer, CurAbs As Long
    Dim CurMonth As Integer, CurYear As Integer, SheetName As String
    Dim MonthNames As Variant, DataCopied As Boolean, tbl As ListObject
    Dim arrHeaders As Variant
    
    MonthNames = Array("", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    arrHeaders = Array("Trade Date", "Deal number", "Volume", "Differential", "Term", "Trade Type", "Book Strategy")
    
    ' --- 1. SYSTEM LISTS (Self-learning database for drop-downs) ---
    If Not SheetExists("System_Lists") Then
        Set wsList = ThisWorkbook.Sheets.Add
        wsList.Name = "System_Lists"
        wsList.Visible = xlSheetHidden ' Hides this sheet from normal view
        
        wsList.Cells(1, 1).Value = "Trade Type"
        wsList.Cells(2, 1).Value = "ARV"
        wsList.Cells(3, 1).Value = "HTT"
        wsList.Cells(4, 1).Value = "WTT"
        
        wsList.Cells(1, 2).Value = "Book Strategy"
        wsList.Cells(2, 2).Value = "FSP Initiative Sales"
        wsList.Cells(3, 2).Value = "Midcon - Sales"
        wsList.Cells(4, 2).Value = "TM - Sale"
        wsList.Cells(5, 2).Value = "TM - Time Trade"
    Else
        Set wsList = ThisWorkbook.Sheets("System_Lists")
    End If

    Set wsSource = ThisWorkbook.Sheets("Data Entry")
    
    ' --- 2. ENFORCE HEADERS ON DATA ENTRY SHEET ---
    wsSource.Range("A1:G1").Value = arrHeaders
    wsSource.Range("A1:G1").Font.Bold = True
    
    lastRow = wsSource.Cells(wsSource.Rows.Count, "E").End(xlUp).Row
    If lastRow < 2 Then
        UpdateDataValidation wsSource, wsList
        MsgBox "Headers and drop-downs updated, but no data found to distribute.", vbInformation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    DataCopied = False

    ' --- 3. PROCESS TRADES ---
    For r = 2 To lastRow
        colDate = Trim(wsSource.Cells(r, 5).Value) ' Term
        trdType = Trim(wsSource.Cells(r, 6).Value) ' Trade Type
        bkStrat = Trim(wsSource.Cells(r, 7).Value) ' Book Strategy
        
        ' Add new Trade Types/Strategies to the database quietly
        If Len(trdType) > 0 And wsList.Columns(1).Find(trdType, LookAt:=xlWhole) Is Nothing Then
            wsList.Cells(wsList.Cells(wsList.Rows.Count, 1).End(xlUp).Row + 1, 1).Value = trdType
        End If
        If Len(bkStrat) > 0 And wsList.Columns(2).Find(bkStrat, LookAt:=xlWhole) Is Nothing Then
            wsList.Cells(wsList.Cells(wsList.Rows.Count, 2).End(xlUp).Row + 1, 2).Value = bkStrat
        End If
        
        ' Process Dates
        If Len(colDate) >= 5 Then
            If InStr(colDate, "-") > 0 Then
                dateParts = Split(colDate, "-")
                startStr = Replace(Trim(dateParts(0)), " ", "")
                endStr = Replace(Trim(dateParts(1)), " ", "")
            Else
                startStr = Replace(colDate, " ", "")
                endStr = startStr
            End If
            
            If Len(startStr) >= 5 And Len(endStr) >= 5 Then
                m1_str = Left(startStr, 3): y1_str = Right(startStr, 2)
                m2_str = Left(endStr, 3): y2_str = Right(endStr, 2)
                m1_idx = GetMonthIndex(m1_str): m2_idx = GetMonthIndex(m2_str)
                
                If m1_idx > 0 And m2_idx > 0 And IsNumeric(y1_str) And IsNumeric(y2_str) Then
                    StartYear = CInt(y1_str): EndYear = CInt(y2_str)
                    StartAbs = (StartYear * 12) + m1_idx - 1
                    EndAbs = (EndYear * 12) + m2_idx - 1
                    
                    If EndAbs >= StartAbs Then
                        For i = 0 To (EndAbs - StartAbs)
                            CurAbs = StartAbs + i
                            CurMonth = (CurAbs Mod 12) + 1
                            CurYear = CurAbs \ 12
                            SheetName = MonthNames(CurMonth) & " " & Format(CurYear, "00")
                            
                            ' --- 4. AUTO-BUILD TARGET SHEET & ADD HEADERS ---
                            If Not SheetExists(SheetName) Then
                                Set wsTarget = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
                                wsTarget.Name = SheetName
                                wsTarget.Range("A1:G1").Value = arrHeaders
                                wsTarget.Range("A1:G1").Font.Bold = True
                            Else
                                Set wsTarget = ThisWorkbook.Sheets(SheetName)
                                ' Failsafe: Add headers if the existing sheet is totally blank
                                If wsTarget.Cells(1, 1).Value = "" Then
                                    wsTarget.Range("A1:G1").Value = arrHeaders
                                    wsTarget.Range("A1:G1").Font.Bold = True
                                End If
                            End If
                            
                            nextRow = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row + 1
                            wsSource.Range(wsSource.Cells(r, 1), wsSource.Cells(r, 7)).Copy Destination:=wsTarget.Cells(nextRow, 1)
                            
                            ' --- 5. FORMAT EXCEL TABLE, SIZES, AND FROZEN HEADERS ---
                            targetLastRow = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row
                            
                            ' Wipe the table variable from memory before checking the new sheet
                            Set tbl = Nothing
                            
                            On Error Resume Next
                            Set tbl = wsTarget.ListObjects(1)
                            On Error GoTo 0
                            
                            If tbl Is Nothing Then
                                Set tbl = wsTarget.ListObjects.Add(xlSrcRange, wsTarget.Range("A1:G" & targetLastRow), , xlYes)
                                tbl.TableStyle = "TableStyleMedium2"
                            Else
                                tbl.Resize wsTarget.Range("A1:G" & targetLastRow)
                            End If
                            
                            ' Enforce Data Formatting
                            wsTarget.Range("A2:A" & targetLastRow).NumberFormat = "d/m/yyyy"
                            wsTarget.Range("C2:C" & targetLastRow).NumberFormat = "#,##0"
                            wsTarget.Range("D2:D" & targetLastRow).NumberFormat = "$#,##0.00#;$(#,##0.00#)"
                            
                            ' Enforce Column Widths
                            wsTarget.Columns("A:F").ColumnWidth = 20
                            wsTarget.Columns("G:G").ColumnWidth = 25
                            
                            ' Enforce Frozen Header Row
                            wsTarget.Activate
                            With ActiveWindow
                                If .FreezePanes Then .FreezePanes = False
                                .SplitColumn = 0
                                .SplitRow = 1
                                .FreezePanes = True
                            End With
                            
                            DataCopied = True
                        Next i
                    End If
                End If
            End If
        End If
    Next r

    ' --- 6. APPLY DROP-DOWNS & CLEAN UP ---
    UpdateDataValidation wsSource, wsList

    If DataCopied Then
        wsSource.Range("A2:G" & lastRow).ClearContents
        
        ' ENFORCE FORMATTING ON THE DATA ENTRY SHEET FOR NEXT TIME
        wsSource.Range("A2:A" & lastRow).NumberFormat = "d/m/yyyy"
        wsSource.Range("C2:C" & lastRow).NumberFormat = "#,##0"
        wsSource.Range("D2:D" & lastRow).NumberFormat = "$#,##0.00#;$(#,##0.00#)"
        
        ' Return user to the Data Entry sheet
        wsSource.Activate
        
        Application.ScreenUpdating = True
        MsgBox "Success! Data distributed, columns sized, and headers pinned.", vbInformation
    Else
        wsSource.Activate
        Application.ScreenUpdating = True
        MsgBox "Macro finished, but no matching dates were formatted correctly. Data was not cleared.", vbExclamation
    End If
End Sub

' --- HELPER SUB: Updates Data Validation Drop-downs ---
Sub UpdateDataValidation(wsSource As Worksheet, wsList As Worksheet)
    Dim lastTrd As Long, lastStrat As Long
    lastTrd = wsList.Cells(wsList.Rows.Count, 1).End(xlUp).Row
    lastStrat = wsList.Cells(wsList.Rows.Count, 2).End(xlUp).Row
    
    ' Create dynamic lists based on the hidden database sheet
    ThisWorkbook.Names.Add Name:="TradeTypeList", RefersTo:="=System_Lists!$A$2:$A$" & lastTrd
    ThisWorkbook.Names.Add Name:="BookStrategyList", RefersTo:="=System_Lists!$B$2:$B$" & lastStrat
    
    ' Apply drop-down to Column F (Trade Type)
    With wsSource.Range("F2:F5000").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=TradeTypeList"
        .IgnoreBlank = True: .InCellDropdown = True: .ShowInput = True: .ShowError = False
    End With
    
    ' Apply drop-down to Column G (Book Strategy)
    With wsSource.Range("G2:G5000").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BookStrategyList"
        .IgnoreBlank = True: .InCellDropdown = True: .ShowInput = True: .ShowError = False
    End With
End Sub

' Helper Function: Converts 3-letter string to month number
Function GetMonthIndex(mStr As String) As Integer
    Dim MonthNames As Variant, k As Integer
    MonthNames = Array("", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    mStr = StrConv(mStr, vbProperCase)
    GetMonthIndex = 0
    For k = 1 To 12
        If MonthNames(k) = mStr Then
            GetMonthIndex = k: Exit For
        End If
    Next k
End Function

' Helper Function: Checks if a worksheet exists
Function SheetExists(shtName As String) As Boolean
    Dim sht As Worksheet
    On Error Resume Next
    Set sht = ThisWorkbook.Sheets(shtName)
    On Error GoTo 0
    SheetExists = Not sht Is Nothing
End Function
