Attribute VB_Name = "WarframeMarket"
Option Explicit

Public Sub SortTableByStock()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow > 3 Then
        If Trim(LCase(ws.Cells(lastRow, 1).Value)) = "total" Then lastRow = lastRow - 1
        If lastRow >= 3 Then
            Dim sortRange As Range
            Set sortRange = ws.Range("A3:H" & lastRow)
            ws.Sort.SortFields.Clear
            
            ' Primary: Column C (Stock) ASCENDING (lowest stock first)
            ws.Sort.SortFields.Add Key:=ws.Range("C3:C" & lastRow), _
                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                
            ' Secondary: Column B (Price) DESCENDING when stock is equal
            ws.Sort.SortFields.Add Key:=ws.Range("B3:B" & lastRow), _
                SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
                
            ' Tertiary: Column A (Item Name) ASCENDING
            ws.Sort.SortFields.Add Key:=ws.Range("A3:A" & lastRow), _
                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                
            With ws.Sort
                .SetRange sortRange
                .Header = xlNo
                .MatchCase = False
                .Orientation = xlTopToBottom
                .Apply
            End With
            
            ' Re-bind row formulas & styling cleanly after sort
            Dim r As Long, baseStock As Long, minusPos As Long, fStr As String
            Dim hStr As String, plusPos As Long, baseAllTime As Long
            For r = 3 To lastRow
                ' Col C: Stock formula (=base-E{r})
                fStr = Trim(CStr(ws.Cells(r, 3).Formula))
                baseStock = 1
                If Left(fStr, 1) = "=" Then
                    minusPos = InStr(fStr, "-")
                    If minusPos > 2 Then
                        baseStock = Val(Mid(fStr, 2, minusPos - 2))
                    Else
                        baseStock = Val(ws.Cells(r, 3).Value)
                    End If
                Else
                    baseStock = Val(ws.Cells(r, 3).Value)
                End If
                If baseStock < 0 Then baseStock = 0
                ws.Cells(r, 3).Formula = "=" & CStr(baseStock) & "-E" & CStr(r)
                
                ' Col G: Current Session Revenue (=B{r}*E{r})
                ws.Cells(r, 7).Formula = "=B" & CStr(r) & "*E" & CStr(r)
                
                ' Col H: All Time Revenue (=baseAllTime+G{r})
                hStr = Trim(CStr(ws.Cells(r, 8).Formula))
                baseAllTime = 0
                If Left(hStr, 1) = "=" Then
                    plusPos = InStr(hStr, "+")
                    If plusPos > 2 Then
                        baseAllTime = Val(Mid(hStr, 2, plusPos - 2))
                    Else
                        baseAllTime = Val(ws.Cells(r, 8).Value)
                    End If
                Else
                    baseAllTime = Val(ws.Cells(r, 8).Value)
                End If
                ws.Cells(r, 8).Formula = "=" & CStr(baseAllTime) & "+G" & CStr(r)
                
                ' Row zebra fills
                If r Mod 2 = 0 Then
                    ws.Range("A" & r & ":H" & r).Interior.Color = RGB(15, 23, 42)
                Else
                    ws.Range("A" & r & ":H" & r).Interior.Color = RGB(30, 41, 59)
                End If
                
                ' Checkbox font color
                If InStr(CStr(ws.Cells(r, 4).Value), ChrW(&H2611)) > 0 Or ws.Cells(r, 4).Value = True Then
                    ws.Cells(r, 4).Font.Color = RGB(52, 211, 153)
                Else
                    ws.Cells(r, 4).Font.Color = RGB(100, 116, 139)
                End If
            Next r
            
            ' Update Total row formulas
            ws.Cells(lastRow + 1, 7).Formula = "=SUM(G3:G" & CStr(lastRow) & ")"
            ws.Cells(lastRow + 1, 8).Formula = "=SUM(H3:H" & CStr(lastRow) & ")"
        End If
    End If
End Sub

Public Sub SyncFromMarket(Optional ByVal ExtraArgs As String = "")
    On Error Resume Next
    Range("A1").Select
    ThisWorkbook.Save
    
    Dim pyCmd As String
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    
    Dim pyScript As String
    pyScript = Chr(34) & ThisWorkbook.Path & "\warframe_market.py" & Chr(34)
    Dim xlFile As String
    xlFile = Chr(34) & ThisWorkbook.FullName & Chr(34)
    
    If ExtraArgs <> "" Then
        pyCmd = "cmd.exe /c cd /d " & Chr(34) & ThisWorkbook.Path & Chr(34) & " && python " & pyScript & " --sync --file " & xlFile & " " & ExtraArgs
        Call wsh.Run(pyCmd, 1, True)
    Else
        pyCmd = "cmd.exe /c cd /d " & Chr(34) & ThisWorkbook.Path & Chr(34) & " && python " & pyScript & " --sync --file " & xlFile
        Call wsh.Run(pyCmd, 1, False)
    End If
End Sub

Public Sub PushPricesToMarket(Optional ByVal ExtraArgs As String = "")
    On Error Resume Next
    Range("A1").Select
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)
    Dim rScan As Long, lastRowScan As Long
    Dim fStrScan As String, minusScan As Long, bScan As Long, sScan As Long
    lastRowScan = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For rScan = 3 To lastRowScan
        If Trim(LCase(ws.Cells(rScan, 1).Value)) = "total" Then Exit For
        If ws.Cells(rScan, 1).Value <> "" Then
            fStrScan = Trim(CStr(ws.Cells(rScan, 3).Formula))
            bScan = 1
            If Left(fStrScan, 1) = "=" Then
                minusScan = InStr(fStrScan, "-")
                If minusScan > 2 Then
                    bScan = Val(Mid(fStrScan, 2, minusScan - 2))
                Else
                    bScan = Val(ws.Cells(rScan, 3).Value)
                End If
            Else
                bScan = Val(ws.Cells(rScan, 3).Value)
            End If
            sScan = Val(ws.Cells(rScan, 5).Value)
            If bScan - sScan <= 0 Then
                ws.Cells(rScan, 4).Value = ChrW(&H2610) ' ☐
                ws.Cells(rScan, 4).Font.Color = RGB(100, 116, 139) ' Dim Slate
            End If
        End If
    Next rScan
    
    ThisWorkbook.Save
    
    Dim pyCmd As String
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    
    Dim pyScript As String
    pyScript = Chr(34) & ThisWorkbook.Path & "\warframe_market.py" & Chr(34)
    Dim xlFile As String
    xlFile = Chr(34) & ThisWorkbook.FullName & Chr(34)
    
    If ExtraArgs <> "" Then
        pyCmd = "cmd.exe /c cd /d " & Chr(34) & ThisWorkbook.Path & Chr(34) & " && python " & pyScript & " --push -y --file " & xlFile & " " & ExtraArgs
        Call wsh.Run(pyCmd, 1, True)
    Else
        pyCmd = "cmd.exe /c cd /d " & Chr(34) & ThisWorkbook.Path & Chr(34) & " && python " & pyScript & " --push -y --file " & xlFile
        Call wsh.Run(pyCmd, 1, False)
    End If
End Sub

Public Sub UpdateAllTimeRevenue()
    On Error Resume Next
    Range("A1").Select
    ThisWorkbook.Save
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)
    Dim r As Long, lastRow As Long
    Dim eVal As Double, fVal As Double, priceVal As Double
    Dim totalPlat As Double, sessionPlat As Double
    Dim stockFormula As String, baseStock As Double, newBase As Double
    Dim minusPos As Long, plusPos As Long
    Dim hStr As String, baseAllTime As Double, newAllTime As Double
    totalPlat = 0
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 3 To lastRow
        If Trim(LCase(ws.Cells(r, 1).Value)) = "total" Then Exit For
        If ws.Cells(r, 1).Value <> "" Then
            priceVal = Val(ws.Cells(r, 2).Value)
            eVal = Val(ws.Cells(r, 5).Value) ' Col E = Current Session Qty
            fVal = Val(ws.Cells(r, 6).Value) ' Col F = All Time Qty
            
            If eVal > 0 Then
                sessionPlat = eVal * priceVal
                totalPlat = totalPlat + sessionPlat
                ws.Cells(r, 6).Value = fVal + eVal
                
                ' Parse base All-Time Revenue from Col H formula (=base+G{r})
                hStr = Trim(CStr(ws.Cells(r, 8).Formula))
                baseAllTime = 0
                If Left(hStr, 1) = "=" Then
                    plusPos = InStr(hStr, "+")
                    If plusPos > 2 Then
                        baseAllTime = Val(Mid(hStr, 2, plusPos - 2))
                    Else
                        baseAllTime = Val(ws.Cells(r, 8).Value) - sessionPlat
                    End If
                Else
                    baseAllTime = Val(ws.Cells(r, 8).Value) - sessionPlat
                End If
                If baseAllTime < 0 Then baseAllTime = 0
                newAllTime = baseAllTime + sessionPlat
                ws.Cells(r, 8).Formula = "=" & CStr(newAllTime) & "+G" & CStr(r)
                
                ' Parse base stock in Col C formula (=base-E{r})
                stockFormula = Trim(CStr(ws.Cells(r, 3).Formula))
                baseStock = 1
                If Left(stockFormula, 1) = "=" Then
                    minusPos = InStr(stockFormula, "-")
                    If minusPos > 2 Then
                        baseStock = Val(Mid(stockFormula, 2, minusPos - 2))
                    Else
                        baseStock = Val(ws.Cells(r, 3).Value) + eVal
                    End If
                Else
                    baseStock = Val(ws.Cells(r, 3).Value) + eVal
                End If
                
                If baseStock - eVal <= 0 Then
                    ' When stock falls to 0: uncheck visibility checkbox and set base stock to 0
                    ws.Cells(r, 4).Value = ChrW(&H2610) ' ☐
                    ws.Cells(r, 4).Font.Color = RGB(100, 116, 139) ' Dim Slate
                    newBase = 0
                Else
                    newBase = baseStock - eVal
                End If
                
                ws.Cells(r, 3).Formula = "=" & CStr(newBase) & "-E" & CStr(r)
                ws.Cells(r, 5).Value = 0
            End If
        End If
    Next r
    
    Call SortTableByStock
    ThisWorkbook.Save
    If totalPlat > 0 Then
        Application.StatusBar = "All Time Revenue updated (" & Format(totalPlat, "#,##0") & " Plat added)."
    End If
End Sub

Public Sub RefreshColumnsAndFormulas(Optional ByVal ExtraArgs As String = "")
    On Error Resume Next
    Range("A1").Select
    ThisWorkbook.Save
    
    Dim pyCmd As String
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    
    Dim pyScript As String
    pyScript = Chr(34) & ThisWorkbook.Path & "\warframe_market.py" & Chr(34)
    Dim xlFile As String
    xlFile = Chr(34) & ThisWorkbook.FullName & Chr(34)
    
    If ExtraArgs <> "" Then
        pyCmd = "cmd.exe /c cd /d " & Chr(34) & ThisWorkbook.Path & Chr(34) & " && python " & pyScript & " --update-columns --file " & xlFile & " " & ExtraArgs
        Call wsh.Run(pyCmd, 1, True)
    Else
        pyCmd = "cmd.exe /c cd /d " & Chr(34) & ThisWorkbook.Path & Chr(34) & " && python " & pyScript & " --update-columns --file " & xlFile
        Call wsh.Run(pyCmd, 1, False)
    End If
End Sub

Public Sub StatusDropDownChange()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)
    Dim dd As Object
    Set dd = ws.DropDowns("Status_DropDown")
    If Not dd Is Nothing Then
        Dim selectedText As String
        selectedText = dd.List(dd.ListIndex)
        If selectedText <> "" Then
            Application.EnableEvents = False
            ws.Range("J6").Value = selectedText
            Select Case LCase(Trim(selectedText))
                Case "online in game"
                    ws.Range("J6").Font.Color = RGB(56, 189, 248) ' Cyan Accent
                Case "online"
                    ws.Range("J6").Font.Color = RGB(52, 211, 153) ' Emerald Green
                Case "invisible"
                    ws.Range("J6").Font.Color = RGB(148, 163, 184) ' Slate Gray
            End Select
            Application.EnableEvents = True
            ThisWorkbook.Save
            Call SetUserStatus(selectedText)
        End If
    End If
End Sub

Public Sub SetUserStatus(ByVal newStatus As String, Optional ByVal ExtraArgs As String = "")
    On Error Resume Next
    If Trim(newStatus) = "" Then Exit Sub
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)
    
    Select Case LCase(Trim(newStatus))
        Case "online in game"
            ws.Range("J6").Font.Color = RGB(56, 189, 248) ' Cyan Accent
        Case "online"
            ws.Range("J6").Font.Color = RGB(52, 211, 153) ' Emerald Green
        Case "invisible"
            ws.Range("J6").Font.Color = RGB(148, 163, 184) ' Slate Gray
    End Select
    
    Dim pyCmd As String
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    
    Dim pyScript As String
    pyScript = Chr(34) & ThisWorkbook.Path & "\warframe_market.py" & Chr(34)
    Dim xlFile As String
    xlFile = Chr(34) & ThisWorkbook.FullName & Chr(34)
    Dim stArg As String
    stArg = Chr(34) & newStatus & Chr(34)
    
    If ExtraArgs <> "" Then
        pyCmd = "cmd.exe /c cd /d " & Chr(34) & ThisWorkbook.Path & Chr(34) & " && python " & pyScript & " --status " & stArg & " --file " & xlFile & " " & ExtraArgs
        Call wsh.Run(pyCmd, 1, True)
    Else
        pyCmd = "cmd.exe /c cd /d " & Chr(34) & ThisWorkbook.Path & Chr(34) & " && python " & pyScript & " --status " & stArg & " --file " & xlFile
        Call wsh.Run(pyCmd, 0, False)
    End If
End Sub
