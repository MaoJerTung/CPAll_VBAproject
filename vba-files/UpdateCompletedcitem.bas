Attribute VB_Name = "UpdateCompletedcitem"
Sub Update()
dim Dcitem_name, Locationfile, Datedata As string
Application.displayalerts = False
Application.ScreenUpdating = False

Datedata = [Text(U4,"DDMMYY")]
Locationfile = range("U5").Value & "\"
Dcitem_name = "CompleteDcitem" & Datedata & ".xlsx"
    'Cleardata'
    ActiveSheet.PivotTables("Datadcitem").ClearAllFilters
    Sheets("dcitem").Select
    On error Resume Next
    ActiveSheet.ShowAlldata
    'Copy Pasee Data to Filelocation'
    Workbooks.Open Filename:= Locationfile & Dcitem_name
        Cells.Select
    Selection.Copy
    Windows("CompleteDcitemUpdate.xlsb").Activate
    Sheets("dcitem").Select
        Cells.Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Pivot").Select
    Application.CutCopyMode = False
    'Refreshdata'
    ActiveWorkbook.RefreshAll
    'CloseRawdata'
    Windows(Dcitem_name).Activate
    Activewindow.Close
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Sheets("Pivot").Select
    ActiveWorkbook.Save
    MsgBox "Complete!", , "CompleteDcitemUpdate"
End Sub

Private Sub HideTooltips ()
    Application.ExecuteExcel4Macro "Show.ToolBar(""Ribbon"", False)"
End Sub

Private Sub UnhideTooltips ()
    Application.ExecuteExcel4Macro "Show.ToolBar(""Ribbon"", True)"
End Sub

