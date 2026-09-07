"""
Warframe Market Sell Sync - VBA Macro & Excel COM Automation
============================================================
Injects VBA modules, shape macro buttons, selection event listeners,
and manages clean .xlsx exports.
"""

import os
import openpyxl
from typing import Optional

from .config import FONT_NAME
from .sheet_layout import get_open_excel_workbook

VBA_MODULE_CODE = '''Option Explicit

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
    pyScript = Chr(34) & ThisWorkbook.Path & "\\warframe_market.py" & Chr(34)
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
    pyScript = Chr(34) & ThisWorkbook.Path & "\\warframe_market.py" & Chr(34)
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
    pyScript = Chr(34) & ThisWorkbook.Path & "\\warframe_market.py" & Chr(34)
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
    pyScript = Chr(34) & ThisWorkbook.Path & "\\warframe_market.py" & Chr(34)
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
'''

SHEET_EVENT_CODE = """Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    If Target.Cells.CountLarge > 20 Then Exit Sub
    
    ' Auto-sync warframe.market online presence when J6 status dropdown changes
    If Target.Column = 10 And Target.Row = 6 Then
        Dim newSt As String
        newSt = Trim(CStr(Target.Value))
        If newSt <> "" Then
            Dim ddSync As Object
            Set ddSync = Me.DropDowns("Status_DropDown")
            If Not ddSync Is Nothing Then
                Dim iSt As Long
                For iSt = 1 To ddSync.ListCount
                    If LCase(Trim(ddSync.List(iSt))) = LCase(newSt) Then
                        ddSync.ListIndex = iSt
                        Exit For
                    End If
                Next iSt
            End If
            Call WarframeMarket.SetUserStatus(newSt)
        End If
        Exit Sub
    End If
    
    Dim cell As Range
    For Each cell In Target
        If (cell.Column = 5 Or cell.Column = 3) And cell.Row >= 3 Then
            Dim r As Long
            r = cell.Row
            If Trim(LCase(Cells(r, 1).Value)) <> "total" And Cells(r, 1).Value <> "" Then
                Dim fStr As String, minusPos As Long, bStock As Long, sSold As Long, remStk As Long
                fStr = Trim(CStr(Cells(r, 3).Formula))
                bStock = 1
                If Left(fStr, 1) = "=" Then
                    minusPos = InStr(fStr, "-")
                    If minusPos > 2 Then
                        bStock = Val(Mid(fStr, 2, minusPos - 2))
                    Else
                        bStock = Val(Cells(r, 3).Value)
                    End If
                Else
                    bStock = Val(Cells(r, 3).Value)
                End If
                sSold = Val(Cells(r, 5).Value)
                remStk = bStock - sSold
                
                Application.EnableEvents = False
                If remStk <= 0 Then
                    Cells(r, 4).Value = ChrW(&H2610) ' ☐
                    Cells(r, 4).Font.Color = RGB(100, 116, 139) ' Dim Slate
                Else
                    If InStr(CStr(Cells(r, 4).Value), ChrW(&H2610)) > 0 Or InStr(CStr(Cells(r, 4).Value), "☐") > 0 Or Cells(r, 4).Value = False Then
                        Cells(r, 4).Value = ChrW(&H2611) ' ☑
                        Cells(r, 4).Font.Color = RGB(52, 211, 153) ' Emerald Green
                    End If
                End If
                Application.EnableEvents = True
            End If
        End If
    Next cell
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error Resume Next
    If Target.Cells.CountLarge > 1 Then Exit Sub
    
    ' 1. Interactive Checkbox Toggle in Column D (Visible)
    If Target.Column = 4 And Target.Row >= 3 Then
        If Trim(LCase(Cells(Target.Row, 1).Value)) <> "total" And Cells(Target.Row, 1).Value <> "" Then
            Application.EnableEvents = False
            If InStr(CStr(Target.Value), ChrW(&H2611)) > 0 Or InStr(CStr(Target.Value), "☑") > 0 Or Target.Value = True Then
                Target.Value = ChrW(&H2610) ' ☐
                Target.Font.Color = RGB(100, 116, 139) ' Dim Slate
            Else
                Target.Value = ChrW(&H2611) ' ☑
                Target.Font.Color = RGB(52, 211, 153) ' Emerald Green
            End If
            Range("A1").Select
            Application.EnableEvents = True
            Exit Sub
        End If
    End If
    
    ' 2. Trigger actions on cell clicks (with automatic save first)
    If Target.Row = 2 And (Target.Column = 8 Or Target.Column = 9) Then ' I2: Update All Time Revenue
        Application.EnableEvents = False
        Range("A1").Select
        ThisWorkbook.Save
        Call WarframeMarket.UpdateAllTimeRevenue
        Application.EnableEvents = True
    ElseIf Target.Column = 9 Or Target.Column = 10 Then
        Select Case Target.Row
            Case 7
                Application.EnableEvents = False
                Range("A1").Select
                ThisWorkbook.Save
                Call WarframeMarket.SyncFromMarket
                Application.EnableEvents = True
            Case 9
                Application.EnableEvents = False
                Range("A1").Select
                ThisWorkbook.Save
                Call WarframeMarket.PushPricesToMarket
                Application.EnableEvents = True
            Case 11
                Application.EnableEvents = False
                Range("A1").Select
                ThisWorkbook.Save
                Call WarframeMarket.UpdateAllTimeRevenue
                Application.EnableEvents = True
            Case 13
                Application.EnableEvents = False
                Range("A1").Select
                ThisWorkbook.Save
                Call WarframeMarket.RefreshColumnsAndFormulas
                Application.EnableEvents = True
        End Select
    End If
End Sub
"""


def inject_vba_and_shapes(input_path: str, output_xlsm: Optional[str] = None) -> bool:
    """Uses win32com to inject VBA module, sheet event handler, and macro shapes."""
    abs_input = os.path.abspath(input_path)
    abs_output = os.path.abspath(output_xlsm) if output_xlsm else abs_input

    try:
        import win32com.client
    except ImportError:
        print("[!] pywin32 not installed. Skipping VBA injection.")
        return False

    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(abs_input)
        ws = wb.Sheets(1)

        # 1. Inject or update WarframeMarket standard module
        vb_proj = wb.VBProject
        mod_found = False
        for comp in vb_proj.VBComponents:
            if comp.Name == "WarframeMarket":
                mod_found = True
                comp.CodeModule.DeleteLines(1, comp.CodeModule.CountOfLines)
                comp.CodeModule.AddFromString(VBA_MODULE_CODE)
                break
        if not mod_found:
            new_mod = vb_proj.VBComponents.Add(1)  # vbext_ct_StdModule
            new_mod.Name = "WarframeMarket"
            new_mod.CodeModule.AddFromString(VBA_MODULE_CODE)

        # 2. Inject Worksheet_SelectionChange event into Sheet1
        sheet_comp = None
        sheet_code_name = ws.CodeName or ws.Name
        for comp in vb_proj.VBComponents:
            if comp.Type == 100 and (comp.Name == sheet_code_name or comp.Name == ws.Name):
                sheet_comp = comp
                break
        if sheet_comp:
            code_mod = sheet_comp.CodeModule
            if code_mod.CountOfLines > 0:
                code_mod.DeleteLines(1, code_mod.CountOfLines)
            code_mod.AddFromString(SHEET_EVENT_CODE)

        # 3. Add or update OnAction Shape Buttons in Column J (Col 10)
        buttons_info = [
            ("Btn_Sync", 7, "▶  Sync from Market", "WarframeMarket.SyncFromMarket", (2, 132, 199)),
            ("Btn_Push", 9, "⬆  Push Prices & Stock", "WarframeMarket.PushPricesToMarket", (5, 150, 105)),
            ("Btn_End", 11, "🔄  Update All Time Revenue", "WarframeMarket.UpdateAllTimeRevenue", (217, 119, 6)),
            ("Btn_Refresh", 13, "⚡  Refresh Formulas", "WarframeMarket.RefreshColumnsAndFormulas", (71, 85, 105))
        ]

        # Remove existing buttons and status dropdown shape if already present
        for shape in list(ws.Shapes):
            if shape.Name.startswith("Btn_") or shape.Name == "Status_DropDown":
                shape.Delete()

        for btn_name, row_idx, text, macro_name, (r, g, b) in buttons_info:
            target_cell = ws.Cells(row_idx, 10)
            left = target_cell.Left + 4
            top = target_cell.Top + 2
            width = min(260, target_cell.Width - 8)
            height = target_cell.Height - 4

            shp = ws.Shapes.AddShape(5, left, top, width, height)  # 5 = msoShapeRoundedRectangle
            shp.Name = btn_name
            shp.TextFrame2.TextRange.Characters.Text = text
            shp.TextFrame2.TextRange.Font.Name = FONT_NAME
            shp.TextFrame2.TextRange.Font.Size = 9.5
            shp.TextFrame2.TextRange.Font.Bold = True
            shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 16777215  # White
            shp.TextFrame2.VerticalAnchor = 3  # msoAnchorMiddle
            shp.TextFrame2.TextRange.ParagraphFormat.Alignment = 2  # msoAlignCenter
            shp.Fill.ForeColor.RGB = r + (g * 256) + (b * 65536)
            shp.Line.Visible = False
            shp.OnAction = macro_name

        # 4. Add Status Dropdown Form Control on J6 (Col 10, Row 6)
        try:
            target_j6 = ws.Range("J6")
            dd = ws.DropDowns().Add(target_j6.Left + 2, target_j6.Top + 2, target_j6.Width - 4, target_j6.Height - 4)
            dd.Name = "Status_DropDown"
            dd.AddItem("Online in Game")
            dd.AddItem("Online")
            dd.AddItem("Invisible")
            curr_j6 = str(target_j6.Value or "").strip().lower()
            if curr_j6 == "online":
                dd.ListIndex = 2
            elif curr_j6 == "invisible":
                dd.ListIndex = 3
            else:
                dd.ListIndex = 1
            dd.OnAction = "WarframeMarket.StatusDropDownChange"
        except Exception as e:
            print(f"[!] Note on Status DropDown shape: {e}")

        # 5. Ensure Data Validation on J6 (Status Dropdown)
        try:
            val_rng = ws.Range("J6")
            val_rng.Validation.Delete()
            # 3 = xlValidateList, 1 = xlValidAlertStop, 1 = xlBetween
            val_rng.Validation.Add(3, 1, 1, "Online,Online in Game,Invisible")
            val_rng.Validation.IgnoreBlank = False
            val_rng.Validation.InCellDropdown = True
        except Exception:
            pass

        if abs_output == abs_input:
            wb.Save()
        else:
            wb.SaveAs(abs_output, 52)  # 52 = xlOpenXMLWorkbookMacroEnabled
        wb.Close(False)
        wb = None
        return True
    except Exception as e:
        print(f"[!] Note on VBA injection: {e}")
        return False
    finally:
        try:
            if 'wb' in locals() and wb:
                wb.Close(False)
        except Exception:
            pass
        try:
            if excel:
                excel.Quit()
        except Exception:
            pass
