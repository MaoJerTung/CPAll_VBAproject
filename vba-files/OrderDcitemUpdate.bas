Attribute VB_Name = "OrderDcitemUpdate"
Public  Sub OrderDcitemandXponsUpdate()
    Dim Dcitemfile_lct, Branchfile_name, Orderfile_lct, Orderfile_name, Dcitem_name, Branch_name, _
    Xponsfile_lct, Xponsfile_name, Xp_name, Xp_Sheet_name As String
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Orderfile_lct = range("C7").Value
    Dcitemfile_lct = range("C8").Value
    Branchfile_name = range("C9").Value
    Xponsfile_lct = range("C11").Value
    Xp_Sheet_name = range("C12").Value
    Datedata = [Text(C3,"DDMMBB")]
    Xponsfile_name = "XPONS SUMMARY ALLDC" & " " & Datedata & range("E5").Value

    Windows("OrderingUpdate Automation.xlsm").Activate
    Sheets("HomePage").Select
    'Open Branch & Xpons'
    Workbooks.Open Filename:= Dcitemfile_lct & Branchfile_name
    Workbooks.Open Filename:= Xponsfile_lct & Xponsfile_name
    'Loop Section'
    Windows("OrderingUpdate Automation.xlsm").Activate
    Sheets("HomePage").Select
    Range("DataPageAI3").Value = 0
    Count = 0
    Do While Range("DataPageAI3").Value < Range("Maximum").Value
        Count = Count + 1
        Windows("OrderingUpdate Automation.xlsm").Activate
        Sheets("HomePage").Select
        Range("DataPageAI3").Value = Range("DataPageAI3").Value + 1
        Orderfile_name = range("AC6").Value
        Dcitem_name = range("AD6").Value
        Branch_name = range("AF6").Value
        Xp_name = range("AE6").Value
        Workbooks.Open Filename:= Orderfile_lct & Orderfile_name 'Open Order file'
        Workbooks.Open Filename:= Dcitemfile_lct & Dcitem_name 'Open Dcitem file'
        'Part Dcitem'
        Cells.Select
        Selection.Copy
        Windows(Orderfile_name).Activate
        Sheets("Database").Select
        Cells.Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        'Part Branch'
        Windows(Branchfile_name).Activate
        Sheets(Branch_name).Select
        Cells.Select
        Selection.Copy
        Windows(Orderfile_name).Activate
        Sheets("Order Branch").Select
        Cells.Select
        ActiveSheet.Paste
        'Part Xpons'
        Windows(Xponsfile_name).Activate
        Sheets(Xp_name).Select
        Cells.Select
        Selection.Copy
        Windows(Orderfile_name).Activate
        On error Resume Next
        Sheets(Xp_Sheet_name).Select
        ActiveSheet.ShowAlldata
        Cells.Select
        ActiveSheet.Paste
        'Part Close Order File'
        Sheets("Database").Select
        ActiveWorkbook.Save
        ActiveWindow.Close
        Windows(Dcitem_name).Activate
        ActiveWindow.Close
        Windows("OrderingUpdate Automation.xlsm").Activate
        Sheets("HomePage").Select
        pctCompl = Count * 9
        Application.StatusBar = "Create Ordering & Xpons Data... " & pctCompl & "% Completed"
    Loop
    Windows(Branchfile_name).Activate
    ActiveWindow.Close
    Windows(Xponsfile_name).Activate
    ActiveWindow.Close
    Application.StatusBar = "Create Ordering & Xpons Data... " & "100% Completed"
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Complete!!",,"Dcitem & XponsUpdate"

End Sub

Public  Sub OrderDcitemUpdate()
    Dim Dcitemfile_lct, Branchfile_name, Orderfile_lct, Orderfile_name, Dcitem_name, Branch_name, Xp_Sheet_name As String
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Orderfile_lct = range("C7").Value
    Dcitemfile_lct = range("C8").Value
    Branchfile_name = range("C9").Value
    Xp_Sheet_name = range("C12").Value
    Datedata = [Text(C3,"DDMMBB")]

    Windows("OrderingUpdate Automation.xlsm").Activate
    Sheets("HomePage").Select
    'Open Branch'
    Workbooks.Open Filename:= Dcitemfile_lct & Branchfile_name
    'Loop Section'
    Windows("OrderingUpdate Automation.xlsm").Activate
    Sheets("HomePage").Select
    Range("DataPageAI3").Value = 0
    Count = 0
    Do While Range("DataPageAI3").Value < Range("Maximum").Value
        Count = Count + 1
        Windows("OrderingUpdate Automation.xlsm").Activate
        Sheets("HomePage").Select
        Range("DataPageAI3").Value = Range("DataPageAI3").Value + 1
        Orderfile_name = range("AC6").Value
        Dcitem_name = range("AD6").Value
        Branch_name = range("AF6").Value
        Workbooks.Open Filename:= Orderfile_lct & Orderfile_name 'Open Order file'
        Workbooks.Open Filename:= Dcitemfile_lct & Dcitem_name 'Open Dcitem file'
        'Part Dcitem'
        Cells.Select
        Selection.Copy
        Windows(Orderfile_name).Activate
        Sheets("Database").Select
        Cells.Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        'Part Branch'
        Windows(Branchfile_name).Activate
        Sheets(Branch_name).Select
        Cells.Select
        Selection.Copy
        Windows(Orderfile_name).Activate
        Sheets("Order Branch").Select
        Cells.Select
        ActiveSheet.Paste
        Sheets(Xp_Sheet_name).Select
        Cells.ClearContents
        Sheets("Database").Select
        ActiveWorkbook.Save
        ActiveWindow.Close
        Windows(Dcitem_name).Activate
        ActiveWindow.Close
        Windows("OrderingUpdate Automation.xlsm").Activate
        Sheets("HomePage").Select
        pctCompl = Count * 9
        Application.StatusBar = "Create Ordering Data... " & pctCompl & "% Completed"
    Loop
    Windows(Branchfile_name).Activate
    ActiveWindow.Close
    Application.StatusBar = "Create Ordering Data... " & "100% Completed"
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Complete!!",,"DcitemUpdate"
End Sub

Public  Sub XponsUpdate()
    Dim Xponsfile_lct, Xponsfile_name, Orderfile_lct, Orderfile_name, Xp_name, Xp_Sheet_name As String
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Orderfile_lct = range("C7").Value
    Xponsfile_lct = range("C11").Value
    Xp_Sheet_name = range("C12").Value
    Datedata = [Text(C3,"DDMMBB")]
    Xponsfile_name = "XPONS SUMMARY ALLDC" & " " & Datedata & range("E5").Value

    Windows("OrderingUpdate Automation.xlsm").Activate
    Sheets("HomePage").Select
    'Open Xpons SUMMARY ALLDC'
    Workbooks.Open Filename:= Xponsfile_lct & Xponsfile_name
    'Loop Section'
    Windows("OrderingUpdate Automation.xlsm").Activate
    Sheets("HomePage").Select
    Range("DataPageAI3").Value = 0
    Count = 0
    Do While Range("DataPageAI3").Value < Range("Maximum").Value
        Count = Count + 1
        Windows("OrderingUpdate Automation.xlsm").Activate
        Sheets("HomePage").Select
        Range("DataPageAI3").Value = Range("DataPageAI3").Value + 1
        Orderfile_name = range("AC6").Value
        Xp_name = range("AE6").Value
        Workbooks.Open Filename:= Orderfile_lct & Orderfile_name 'Open Order file'
        'Part Xpons'
        Windows(Xponsfile_name).Activate
        Sheets(Xp_name).Select
        Cells.Select
        Selection.Copy
        Windows(Orderfile_name).Activate
        On error Resume Next
        Sheets(Xp_Sheet_name).Select
        ActiveSheet.ShowAlldata
        Cells.Select
        ActiveSheet.Paste
        'Part Close Order File'
        ActiveWorkbook.Save
        ActiveWindow.Close
        Windows("OrderingUpdate Automation.xlsm").Activate
        Sheets("HomePage").Select
        pctCompl = Count * 9
        Application.StatusBar = "Create XponsData... " & pctCompl & "% Completed"
    Loop
    Windows(Xponsfile_name).Activate
    ActiveWindow.Close
    Application.StatusBar = "Create XponsData... " & "100% Completed"
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Complete!!",,"XponsUpdate"
End Sub

Public  Sub UpdatepromotionOffline()
    Dim Orderfile_lct, Orderfile_name, Promotionfile_name, Pro_new As String
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Orderfile_lct = range("C7").Value
    Promotionfile_name = range("C21").Value
    Pro_new = range("C22").Value

    Windows("OrderingUpdate Automation.xlsm").Activate
    Sheets("HomePage").Select
    'Open Promotion and Create format'
    Workbooks.Open Filename:= Orderfile_lct & Promotionfile_name
    Windows(Promotionfile_name).Activate
    Sheets("Pro Offline").Select
    LastRow = [COUNTA(A:A)]
    Columns("J").Insert Shift:=xlToRight
    Windows("OrderingUpdate Automation.xlsm").Activate
    Sheets("Form_ProOnline").Select
    Range("A2").Copy
    Windows(Promotionfile_name).Activate
    Range("J3").pastespecial xlpasteformulas
    Range("J3:J" & LastRow).pastespecial xlpasteformulas
    'Loop Section'
    Windows("OrderingUpdate Automation.xlsm").Activate
    Range("DataPageAI3").Value = 0
    Count = 0
    Do While Range("DataPageAI3").Value < Range("Maximum").Value
        Count = Count + 1
        Windows("OrderingUpdate Automation.xlsm").Activate
        Sheets("HomePage").Select
        Range("DataPageAI3").Value = Range("DataPageAI3").Value + 1
        Orderfile_name = range("AC6").Value
        Workbooks.Open Filename:= Orderfile_lct & Orderfile_name 'Open Order file'
        'Part Clean Promotion in Orderfile'
        On error Resume Next
        Sheets(Pro_new).Select
        ActiveSheet.ShowAlldata
        range("A:J").ClearContents
        'Part Paste New Promotion Offline'
        Windows(Promotionfile_name).Activate
        Sheets("Pro Offline").Select
        LastRow = [COUNTA(A:A)]
        range("A1:J" & LastRow).Copy
        Windows(Orderfile_name).Activate
        Sheets(Pro_new).Select
        range("A1").pastespecial xlpastevalues
        LastRow = [COUNTA(A:A)]
        Range("A3:J3").Copy
        Range("A3:J" & LastRow).pastespecial xlpasteformats
        Range("A" & LastRow + 1, "J100000").Clear
        'Part Close Order File'
        ActiveWorkbook.Save
        ActiveWindow.Close
        Windows("OrderingUpdate Automation.xlsm").Activate
        Sheets("HomePage").Select
        pctCompl = Count * 9
        Application.StatusBar = "Update PromotionOffline... " & pctCompl & "% Completed"
    Loop
    Windows(Promotionfile_name).Activate
    ActiveWindow.Close
    Application.StatusBar = "Update PromotionOffline... " & "100% Completed"
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Complete!!",,"PromotionOfflineUpdate"
End Sub

Public  Sub UpdatepromotionOld()
    Dim Orderfile_lct, Orderfile_name, Promotionfile_name, Pro_new, Pro_old As String
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Orderfile_lct = range("C7").Value
    Promotionfile_name = range("C21").Value
    Pro_new = range("C22").Value
    Pro_old = range("D22").Value
    Datedata = [Text(C3,"DDMMYY")]
    Dateset = [Text(C20,"DDMMYY")]
    If (Datedata = Dateset) Then
        Windows("OrderingUpdate Automation.xlsm").Activate
        Sheets("HomePage").Select
        'Loop Section'
        Windows("OrderingUpdate Automation.xlsm").Activate
        Range("DataPageAI3").Value = 0
        Count = 0
        Do While Range("DataPageAI3").Value < Range("Maximum").Value
            Count = Count + 1
            Windows("OrderingUpdate Automation.xlsm").Activate
            Sheets("HomePage").Select
            Range("DataPageAI3").Value = Range("DataPageAI3").Value + 1
            Orderfile_name = range("AC6").Value
            Workbooks.Open Filename:= Orderfile_lct & Orderfile_name 'Open Order file'
            'Part Clean Promotion in Orderfile'
            On error Resume Next
            Sheets(Pro_old).Select
            ActiveSheet.ShowAlldata
            LastRow = [COUNTA(A:A)]
            range("A3:J" & LastRow).ClearContents
            Sheets(Pro_new).Select
            ActiveSheet.ShowAlldata
            LastRow = [COUNTA(A:A)]
            range("A3:J" & LastRow).Copy
            Sheets(Pro_old).Select
            range("A3").pastespecial xlpastevalues
            LastRow = [COUNTA(A:A)]
            Range("A" & LastRow + 1, "J100000").Clear
            'Part Close Order File'
            ActiveWorkbook.Save
            ActiveWindow.Close
            Windows("OrderingUpdate Automation.xlsm").Activate
            Sheets("HomePage").Select
            pctCompl = Count * 9
            Application.StatusBar = "Move PromotionOld... " & pctCompl & "% Completed"
        Loop
    Else
        MsgBox "โปรดตรวจสอบวันที่จะอัพโปรเก่า (ScheduleUpdatePromotion)",,"Alert!!"
        Exit Sub
    End If
    Application.StatusBar = "Move PromotionOld... " & "100% Completed"
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Complete!!",,"PromotionOldUpdate"
End Sub

Public  Sub UpdatepromotionOnline()
    Dim Orderfile_lct, Orderfile_name, Promotionfile_name, Pro_online As String
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Orderfile_lct = range("C7").Value
    Promotionfile_name = range("C21").Value
    Pro_online = range("E22").Value

    Windows("OrderingUpdate Automation.xlsm").Activate
    Sheets("HomePage").Select
    'Open Promotion and Create format'
    Workbooks.Open Filename:= Orderfile_lct & Promotionfile_name
    Windows(Promotionfile_name).Activate
    Sheets("Pro Online").Select
    LastRow = [COUNTA(A:A)]
    ActiveWorkbook.Worksheets("Pro Online").AutoFilter.Sort.SortFields.ADD Key:= _
    Range("L2:L" & LastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
    :=xlSortNormal
    With ActiveWorkbook.Worksheets("Pro Online").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ' Range("L2:L" & LastRow).Sort Key1:=Range("L2"), _
    '                  Order1:=xlAscending, _
    '                  Header:=xlYes
    Columns("A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Windows("OrderingUpdate Automation.xlsm").Activate
    Sheets("Form_ProOnline").Select
    Range("A1").Copy
    Windows(Promotionfile_name).Activate
    Range("A3").pastespecial xlpasteformulas
    Range("A3:A" & LastRow).pastespecial xlpasteformulas
    Range("A2:AB" & LastRow).Autofilter
    Range("A2:AB" & LastRow).Autofilter
    ' Range("A2:A" & LastRow).Sort Key1:=Range("A2"), _
    '                  Order1:=xlDescending, _
    '                  Header:=xlYes
    ActiveWorkbook.Worksheets("Pro Online").AutoFilter.Sort.SortFields.ADD Key:= _
    Range("A2:A" & LastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
    :=xlSortNormal
    With ActiveWorkbook.Worksheets("Pro Online").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Check = WorksheetFunction.CountIf(Range("A:A"), "FALSE")
    If Check + 2 = LastRow Then
        Columns("A").Delete
    Else: ActiveSheet.Range("A2:A" & LastRow).AutoFilter Field:=1, Criteria1:="TRUE"
        Range("A3:AB" & LastRow).ClearContents
        ActiveSheet.ShowAllData
    End If
    On error Resume Next
    ' ActiveSheet.ShowAlldata
    ' Columns("A").Delete
    'Loop Section'
    Windows("OrderingUpdate Automation.xlsm").Activate
    Range("DataPageAI3").Value = 0
    Count = 0
    Do While Range("DataPageAI3").Value < Range("Maximum").Value
        Count = Count + 1
        Windows("OrderingUpdate Automation.xlsm").Activate
        Sheets("HomePage").Select
        Range("DataPageAI3").Value = Range("DataPageAI3").Value + 1
        Orderfile_name = range("AC6").Value
        Workbooks.Open Filename:= Orderfile_lct & Orderfile_name 'Open Order file'
        'Part Clean Promotion in Orderfile'
        Sheets(Pro_online).Select
        ActiveSheet.ShowAlldata
        Cells.ClearContents
        'Part Paste New Promotion Online'
        Windows(Promotionfile_name).Activate
        Sheets("Pro Online").Select
        Cells.Copy
        Windows(Orderfile_name).Activate
        Sheets(Pro_online).Select
        Cells.pastespecial xlpastevalues
        LastRow = [COUNTA(A:A)]
        Range("A3:AB3").Copy
        Range("A3:AB" & LastRow).pastespecial xlpasteformats
        Range("A" & LastRow + 1, "AB100000").Clear
        'Part Close Order File'
        ActiveWorkbook.Save
        ActiveWindow.Close
        Windows("OrderingUpdate Automation.xlsm").Activate
        Sheets("HomePage").Select
        pctCompl = Count * 9
        Application.StatusBar = "Update PromotionfileOnline... " & pctCompl & "% Completed"
    Loop
    Windows(Promotionfile_name).Activate
    ActiveWindow.Close
    Application.StatusBar = "Update PromotionOnline... " & "100% Completed"
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Complete!!",,"PromotionOnlineUpdate"
End Sub

Public  Sub UpdateTop1000()
    Dim Top750file_name, Top750_name, Orderfile_lct, Orderfile_name, Top_Sheet_name As String
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Orderfile_lct = range("C7").Value
    Top750file_name = range("C14").Value
    Top_Sheet_name = range("C15").Value

    Windows("OrderingUpdate Automation.xlsm").Activate
    Sheets("HomePage").Select
    'open Top750'
    Workbooks.Open Filename: = Orderfile_lct & Top750file_name
    'Loop Section'
    Windows("OrderingUpdate Automation.xlsm").Activate
    Sheets("HomePage").Select
    Range("DataPageAI3").Value = 0
    Count = 0
    Do While Range("DataPageAI3").Value < Range("Maximum").Value
        Count = Count + 1
        Windows("OrderingUpdate Automation.xlsm").Activate
        Sheets("HomePage").Select
        Range("DataPageAI3").Value = Range("DataPageAI3").Value + 1
        Orderfile_name = range("AC6").Value
        Top750_name = range("AG6").Value
        Workbooks.Open Filename:= Orderfile_lct & Orderfile_name 'Open Order file'
        'Part Top750file'
        Windows(Top750file_name).Activate
        Sheets(Top750_name).Select
        Columns("E:I").Select
        Selection.EntireColumn.Hidden = True
        Range("A5:J5").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.SpecialCells(xlCellTypeVisible).Select
        Selection.Copy
        Windows(Orderfile_name).Activate
        Sheets(Top_Sheet_name).Select
        Range("B4").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        'Part Close Order File'
        ActiveWorkbook.Save
        ActiveWindow.Close
        Windows("OrderingUpdate Automation.xlsm").Activate
        Sheets("HomePage").Select
        pctCompl = Count * 9
        Application.StatusBar = "Update Top1000... " & pctCompl & "% Completed"
    Loop
    Windows(Top750file_name).Activate
    ActiveWindow.Close
    Application.StatusBar = "Update Top1000... " & "100% Completed"
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Complete!!",,"Top1000Update"

End Sub

Public Sub UpdateRank()
    Dim Rankfile_name, RankWH_name, Orderfile_lct, Orderfile_name, Rank_Sheet_name As String
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Orderfile_lct = range("C7").Value
    Rankfile_name = range("C17").Value
    Rank_Sheet_name = range("C18").Value

    Windows("OrderingUpdate Automation.xlsm").Activate
    Sheets("HomePage").Select
    'Open Rank'
    Workbooks.Open Filename: = Orderfile_lct & Rankfile_name
    Windows(Rankfile_name).Activate
    Columns("C:C").Select
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    'Loop Section'
    Windows("OrderingUpdate Automation.xlsm").Activate
    Sheets("HomePage").Select
    Range("DataPageAI3").Value = 0
    Count = 0
    Do While Range("DataPageAI3").Value < Range("Maximum").Value
        Count = Count + 1
        Windows("OrderingUpdate Automation.xlsm").Activate
        Sheets("HomePage").Select
        Range("DataPageAI3").Value = Range("DataPageAI3").Value + 1
        Orderfile_name = range("AC6").Value
        RankWH_name = range("AH6").Value
        Workbooks.Open Filename:= Orderfile_lct & Orderfile_name 'Open Order file'
        'Part Rankfile'
        Windows(Rankfile_name).Activate
        ActiveSheet.Range("A:A").AutoFilter Field:=1, Criteria1:=RankWH_name
        Cells.Select
        Selection.Copy
        Windows(Orderfile_name).Activate
        Sheets(Rank_Sheet_name).Select
        Cells.Select
        ActiveSheet.Paste
        ActiveWorkbook.Save
        ActiveWindow.Close
        Windows("OrderingUpdate Automation.xlsm").Activate
        Sheets("HomePage").Select
        pctCompl = Count * 9
        Application.StatusBar = "UpdateRank... " & pctCompl & "% Completed"
    Loop
    Windows(Rankfile_name).Activate
    ActiveWindow.Close
    Application.StatusBar = "UpdateRank... " & "100% Completed"
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Complete!!",,"RankUpdate"
End Sub

Public  Sub Rename()
    Dim ADD, FilesDir, Newname As String
    Dim Files As Variant
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    If Range("Maximum").Value = 11 Then
        Files = Range("P6:P" & [Counta(P6:P30)] + 5).Value
    Else
        Files = Range("P7:P" & [Counta(P7:P30)] + 6).Value
    End If
    Files = range("Q6:Q" & [Counta(Q6:Q30)] + 5).Value
    ADD = ThisWorkbook.Path & "\"
    For i = 1 To Range("Maximum").Value
        FilesDir = Dir(ADD & Files(i,1) & "*" & range("E5").Value)
        If Range("Maximum").Value = 11 Then
            Newname = Cells(i + 5, 13)
            Name ADD & FilesDir As ADD & Newname
        Else
            Newname = Cells(i + 6, 13)
            Name ADD & FilesDir As ADD & Newname
        End If
        Next
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        MsgBox "Complete!!",,"Rename"
End Sub

Public  Sub RenameNew()
    Dim ADD, FilesDir, Newname, Maxy As String
    Dim Files As Variant
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Maxy = Range("AO5").Value
    Count = Maxy + 5
    Files = Range("AO6:AO" & Count)
    Newname = Range("AP6:AP" & Count)
    ADD = ThisWorkbook.Path & "\"

    For i = 1 To Maxy
        FilesDir = Dir(ADD & Files(i, 1) & "*" & Range("E5").Value)
        Name ADD & FilesDir As ADD & Newname(i, 1)
        Next
End Sub

Public  Sub Save_close()
    Windows("OrderingUpdate Automation.xlsm").Activate
    Sheets("HomePage").Select
    ActiveWorkbook.Save
    ActiveWindow.Close
End Sub

Private Sub HideTooltips ()
    Application.ExecuteExcel4Macro "Show.ToolBar(""Ribbon"", False)"
End Sub

Private Sub UnhideTooltips ()
    Application.ExecuteExcel4Macro "Show.ToolBar(""Ribbon"", True)"
End Sub

Public  Sub LoopTest()
    Dim Dcitemfile_lct, Branchfile_name, Orderfile_lct, Orderfile_name, Dcitem_name, Branch_name As String
    Orderfile_lct = range("C7").Value
    Dcitemfile_lct = range("J24").Value
    Branchfile_name = range("J28").Value

    Windows("OrderingUpdate Automation.xlsm").Activate
    Sheets("HomePage").Select
    Range("DataPageAI3").Value = 0
    Do While Range("DataPageAI3").Value < Range("Maximum").Value
        Windows("OrderingUpdate Automation.xlsm").Activate
        Sheets("HomePage").Select
        Range("DataPageAI3").Value = Range("DataPageAI3").Value + 1
        Orderfile_name = range("AC6").Value
        Dcitem_name = range("AD6").Value
        Branch_name = range("AF6").Value
        Workbooks.Open Filename:= Orderfile_lct & Orderfile_name 'Open Order file'
        Workbooks.Open Filename:= Dcitemfile_lct & Dcitem_name 'Open Dcitem file'
        'Part Dcitem'
        Cells.Select
        Selection.Copy
        Windows(Orderfile_name).Activate
        Sheets("Database").Select
        Cells.Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        'Part Branch'
        Windows(Branchfile_name).Activate
        Sheets(Branch_name).Select
        Cells.Select
        Selection.Copy
        Windows(Orderfile_name).Activate
        Sheets("Order Branch").Select
        Cells.Select
        ActiveSheet.Paste
        'Part Close Order File'
        ActiveWorkbook.Save
        ActiveWindow.Close
        Windows("OrderingUpdate Automation.xlsm").Activate
        Sheets("HomePage").Select
    Loop
    MsgBox "Complete!!",,"OrderDcitemUpdate"

End Sub

Public  Sub Backupfile()
    Dim Run, Path, FilesDir, Folder, Orderfile_lct As String
    Dim Files As Variant
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Orderfile_lct = range("C7").Value
    Datedata = [Text(C3,"DDMMBB")]
    Foldername = "Orderfile " & Datedata
    Files = Array(Range("M6:M20").Value)
    ADD = Orderfile_lct & "BACKUP"
    Count_Orderfile = [Counta(M6:M20)]

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' MkDir ADD & "\" & Foldername
    On Error Resume next
    fso.DeleteFolder ADD & "\" & Foldername

    MkDir ADD & "\" & Foldername
    On Error Resume next
    For i = 1 To Count_Orderfile
        FilesDir = Dir(ThisWorkbook.Path & "\" & Files(0)(i,1))
        fso.CopyFile Orderfile_lct & FilesDir, ADD & "\" & Foldername & "\" & FilesDir
    Next i
    MsgBox "Complete","BACKUP Orderfile"
End Sub