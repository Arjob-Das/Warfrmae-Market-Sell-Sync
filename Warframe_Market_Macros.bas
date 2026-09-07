Attribute VB_Name = "WarframeMarket"
Option Explicit

Public Sub ExportNormalXLSX()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)
    Dim exportDir As String
    exportDir = Trim(CStr(ws.Range("H16").Value))
    If exportDir = "" Or InStr(exportDir, "<") > 0 Or LCase(exportDir) = "none" Then Exit Sub
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(exportDir) Then
        fso.CreateFolder(exportDir)
    End If
    
    Dim baseName As String
    baseName = fso.GetBaseName(ThisWorkbook.Name)
    Dim destPath As String
    destPath = exportDir
    If Right(destPath, 1) <> "\" Then destPath = destPath & "\"
    destPath = destPath & baseName & ".xlsx"
    
    Dim prevAlerts As Boolean
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    
    Dim newWb As Workbook
    ws.Copy
    Set newWb = ActiveWorkbook
    
    ' Remove macro shapes / buttons from clean exported copy
    Dim shp As Shape
    For Each shp In newWb.Sheets(1).Shapes
        shp.Delete
    Next shp
    
    newWb.SaveAs Filename:=destPath, FileFormat:=51 ' 51 = xlOpenXMLWorkbook (.xlsx)
    newWb.Close SaveChanges:=False
    
    Application.DisplayAlerts = prevAlerts
End Sub

Public Sub SyncFromMarket(Optional ByVal ExtraArgs As String = "")
    Dim pyCmd As String
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    
    ThisWorkbook.Save
    Call ExportNormalXLSX
    
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
    Dim pyCmd As String
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    
    ThisWorkbook.Save
    Call ExportNormalXLSX
    
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
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)
    Dim r As Long, lastRow As Long
    Dim cVal As Double, dVal As Double, priceVal As Double
    Dim totalPlat As Double
    totalPlat = 0
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 3 To lastRow
        If Trim(LCase(ws.Cells(r, 1).Value)) = "total" Then Exit For
        If ws.Cells(r, 1).Value <> "" Then
            cVal = Val(ws.Cells(r, 3).Value)
            dVal = Val(ws.Cells(r, 4).Value)
            priceVal = Val(ws.Cells(r, 2).Value)
            If cVal > 0 Then
                totalPlat = totalPlat + (cVal * priceVal)
                ws.Cells(r, 4).Value = dVal + cVal
                ws.Cells(r, 3).Value = 0
            End If
        End If
    Next r
    
    ThisWorkbook.Save
    Call ExportNormalXLSX
    If totalPlat > 0 Then
        Application.StatusBar = "All Time Revenue updated (" & Format(totalPlat, "#,##0") & " Plat added). Clean .xlsx copy exported."
    End If
End Sub

Public Sub RefreshColumnsAndFormulas(Optional ByVal ExtraArgs As String = "")
    Dim pyCmd As String
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    
    ThisWorkbook.Save
    Call ExportNormalXLSX
    
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

