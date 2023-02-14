Attribute VB_Name = "test"

Public  Sub Test()
    MsgBox "Hello World",,"Test" 
End Sub

Sub Combine()
    Dim Branchfile_lct, Branchfile_name As String
    Branchfile_lct = range("B3").Value
    Branchfile_name = range("Branch").Value
        Workbooks.Open Filename:= Branchfile_lct & Branchfile_name
        Windows("BranchCombinationAllOnline.xlsm").Activate
            Sheets("Combine").Select
            Sheets("Combine").Copy After:=Workbooks(Range("Branch").Value).Sheets(10)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Sheets("BB 39001").Select
            Range("A2:EM2").Select
            Selection.Copy
            Sheets("Combine").Select
            Range("B2").Select
            ActiveSheet.Paste
'''''W101'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Sheets("Combine").Select
            Range("F1").Select
            Range("F1").Value = "OP""A""&A1+2"
            Range("F1").Select
            ActiveCell.Replace What:="OP", Replacement:="=", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
                
            Range("F1").Select
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Application.CutCopyMode = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Sheets("BB 39001").Select
            Range("A3:EM50000").Select
            Application.CutCopyMode = False
            Selection.Copy
            Sheets("Combine").Select
            Range("INDIRECT(B1)").Select
            ActiveSheet.Paste
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Range("INDIRECT(F1)").Value = "BB 39001"
            Range("INDIRECT(F1)").Select
            Selection.Copy
            Range("INDIRECT(F1):INDIRECT(G1)").Select
            ActiveSheet.Paste
            Application.CutCopyMode = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''W121'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Sheets("Combine").Select
            Range("F1").Select
            Range("F1").Value = "OP""A""&A1+2"
            Range("F1").Select
            ActiveCell.Replace What:="OP", Replacement:="=", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
                
            Range("F1").Select
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Application.CutCopyMode = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Sheets("CB 39009").Select
            Range("A3:EM50000").Select
            Application.CutCopyMode = False
            Selection.Copy
            Sheets("Combine").Select
            Range("INDIRECT(B1)").Select
            ActiveSheet.Paste
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Range("INDIRECT(F1)").Value = "CB 39009"
            Range("INDIRECT(F1)").Select
            Selection.Copy
            Range("INDIRECT(F1):INDIRECT(G1)").Select
            ActiveSheet.Paste
            Application.CutCopyMode = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''W131'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Sheets("Combine").Select
            Range("F1").Select
            Range("F1").Value = "OP""A""&A1+2"
            Range("F1").Select
            ActiveCell.Replace What:="OP", Replacement:="=", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
                
            Range("F1").Select
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Application.CutCopyMode = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Sheets("BR 39012").Select
            Range("A3:EM50000").Select
            Application.CutCopyMode = False
            Selection.Copy
            Sheets("Combine").Select
            Range("INDIRECT(B1)").Select
            ActiveSheet.Paste
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Range("INDIRECT(F1)").Value = "BR 39012"
            Range("INDIRECT(F1)").Select
            Selection.Copy
            Range("INDIRECT(F1):INDIRECT(G1)").Select
            ActiveSheet.Paste
            Application.CutCopyMode = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''W141'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Sheets("Combine").Select
            Range("F1").Select
            Range("F1").Value = "OP""A""&A1+2"
            Range("F1").Select
            ActiveCell.Replace What:="OP", Replacement:="=", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
                
            Range("F1").Select
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Application.CutCopyMode = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Sheets("HY 39010").Select
            Range("A3:EM50000").Select
            Application.CutCopyMode = False
            Selection.Copy
            Sheets("Combine").Select
            Range("INDIRECT(B1)").Select
            ActiveSheet.Paste
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Range("INDIRECT(F1)").Value = "HY 39010"
            Range("INDIRECT(F1)").Select
            Selection.Copy
            Range("INDIRECT(F1):INDIRECT(G1)").Select
            ActiveSheet.Paste
            Application.CutCopyMode = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''W151'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Sheets("Combine").Select
            Range("F1").Select
            Range("F1").Value = "OP""A""&A1+2"
            Range("F1").Select
            ActiveCell.Replace What:="OP", Replacement:="=", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
                
            Range("F1").Select
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Application.CutCopyMode = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Sheets("NS 39013").Select
            Range("A3:EM50000").Select
            Application.CutCopyMode = False
            Selection.Copy
            Sheets("Combine").Select
            Range("INDIRECT(B1)").Select
            ActiveSheet.Paste
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Range("INDIRECT(F1)").Value = "NS 39013"
            Range("INDIRECT(F1)").Select
            Selection.Copy
            Range("INDIRECT(F1):INDIRECT(G1)").Select
            ActiveSheet.Paste
            Application.CutCopyMode = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''W301'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Sheets("Combine").Select
            Range("F1").Select
            Range("F1").Value = "OP""A""&A1+2"
            Range("F1").Select
            ActiveCell.Replace What:="OP", Replacement:="=", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
                
            Range("F1").Select
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Application.CutCopyMode = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Sheets("LP 39007").Select
            Range("A3:EM50000").Select
            Application.CutCopyMode = False
            Selection.Copy
            Sheets("Combine").Select
            Range("INDIRECT(B1)").Select
            ActiveSheet.Paste
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Range("INDIRECT(F1)").Value = "LP 39007"
            Range("INDIRECT(F1)").Select
            Selection.Copy
            Range("INDIRECT(F1):INDIRECT(G1)").Select
            ActiveSheet.Paste
            Application.CutCopyMode = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''W401'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Sheets("Combine").Select
            Range("F1").Select
            Range("F1").Value = "OP""A""&A1+2"
            Range("F1").Select
            ActiveCell.Replace What:="OP", Replacement:="=", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
                
            Range("F1").Select
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Application.CutCopyMode = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Sheets("ST 39011").Select
            Range("A3:EM50000").Select
            Application.CutCopyMode = False
            Selection.Copy
            Sheets("Combine").Select
            Range("INDIRECT(B1)").Select
            ActiveSheet.Paste
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Range("INDIRECT(F1)").Value = "ST 39011"
            Range("INDIRECT(F1)").Select
            Selection.Copy
            Range("INDIRECT(F1):INDIRECT(G1)").Select
            ActiveSheet.Paste
            Application.CutCopyMode = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''W501'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Sheets("Combine").Select
            Range("F1").Select
            Range("F1").Value = "OP""A""&A1+2"
            Range("F1").Select
            ActiveCell.Replace What:="OP", Replacement:="=", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
                
            Range("F1").Select
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Application.CutCopyMode = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Sheets("SB 39014").Select
            Range("A3:EM50000").Select
            Application.CutCopyMode = False
            Selection.Copy
            Sheets("Combine").Select
            Range("INDIRECT(B1)").Select
            ActiveSheet.Paste
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Range("INDIRECT(F1)").Value = "SB 39014"
            Range("INDIRECT(F1)").Select
            Selection.Copy
            Range("INDIRECT(F1):INDIRECT(G1)").Select
            ActiveSheet.Paste
            Application.CutCopyMode = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''W801'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Sheets("Combine").Select
            Range("F1").Select
            Range("F1").Value = "OP""A""&A1+2"
            Range("F1").Select
            ActiveCell.Replace What:="OP", Replacement:="=", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
                
            Range("F1").Select
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Application.CutCopyMode = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Sheets("KK 39008").Select
            Range("A3:EM50000").Select
            Application.CutCopyMode = False
            Selection.Copy
            Sheets("Combine").Select
            Range("INDIRECT(B1)").Select
            ActiveSheet.Paste
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Range("INDIRECT(F1)").Value = "KK 39008"
            Range("INDIRECT(F1)").Select
            Selection.Copy
            Range("INDIRECT(F1):INDIRECT(G1)").Select
            ActiveSheet.Paste
            Application.CutCopyMode = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''W901'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Sheets("Combine").Select
            Range("F1").Select
            Range("F1").Value = "OP""A""&A1+2"
            Range("F1").Select
            ActiveCell.Replace What:="OP", Replacement:="=", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
                
            Range("F1").Select
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Application.CutCopyMode = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Sheets("MC 39015").Select
            Range("A3:EM50000").Select
            Application.CutCopyMode = False
            Selection.Copy
            Sheets("Combine").Select
            Range("INDIRECT(B1)").Select
            ActiveSheet.Paste
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Range("INDIRECT(F1)").Value = "MC 39015"
            Range("INDIRECT(F1)").Select
            Selection.Copy
            Range("INDIRECT(F1):INDIRECT(G1)").Select
            ActiveSheet.Paste
            Application.CutCopyMode = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Range("A2").Value = "WH"
        Range("A2").Select
        
        Sheets("Combine").Select
        Sheets("Combine").Move
        
        Windows("BranchCombinationAllOnline.xlsm").Activate
        Sheets("Home").Select
        
        MsgBox "Complete!!",,"CombinationAllOnline" 
        
End Sub

Sub Clear()
    Range("A2:BZ200000").Clear
End Sub
        