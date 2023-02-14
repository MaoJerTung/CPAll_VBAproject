Public  Sub Listdirectoryfile()
    Dim FileName, PathName As String
    Dim FSO, Folder, File As Object
    PathName = "D:\RPA_Generatedata\Picture\"
    'Change the PathName to path to your folder
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set Folder = FSO.GetFolder(PathName)
    For Each File In Folder.Files
        FileName = File.name
        ' MsgBox FileName,,"title"
        i = 1
        Cells(i,1) = FileName
        i+=1
    Next File
End Sub
Public Sub Listdirectoryfiletest()
    Dim FileName, PathName, Filepath, Msg As String
    Dim FSO, Folder, File As Object
        PathName = "D:\RPA_Generatedata\Picture\"
        'Change the PathName to path to your folder
        Set FSO = CreateObject("Scripting.FileSystemObject")
        Set Folder = FSO.GetFolder(PathName)
        'i = 1
        For Each File In Folder.Files
            FileName = File.Name
            ' MsgBox FileName,,"title"
            'Cells(i, 1) = FileName
            'i = i + 1
            Filepath = "D:\RPA_Generatedata\Picture\" & FileName
            Msg = "Test RPA"
            Call LineNotifyUploadPic(Msg, Filepath)
        Next File
End Sub
Public  Sub sendpic()
    Dim FileName, PathName,FullFileName, ADD As String 'แก้เป็น path (ที่อยู่ไฟล์) ของรูปในเครื่อง
    Dim Msg As String

    ADD = ThisWorkbook.Path & "\"
    filepath = Dir(ADD & "Picture\" & "*" & ".xls"
    Msg = "Test sendpic"
    Call LineNotifyUploadPic(Msg, filepath)
End Sub

Public  Sub LineNotifyUploadPic(Msg As String, Optional UploadFilepath As String = "none")
    Dim sFilepath As String
    Dim nFile As Integer
    Dim baBuffer() As Byte
    Dim ssPostData1 As String
    Dim ssPostData2 As String
    Dim ssPostData3 As String
    Dim Messagelength As Integer
    Dim objectXML As Object

    Dim arr1() As Byte
    Dim arr2() As Byte
    Dim arr3() As Byte
    Dim arr4() As Byte

    Const STR_BOUNDARY As String = "xxx---RPA---RPA---RPA---xxx"
    LineToken = "8CDnbg8qlWBdBhw0UQcBlWnt5cNL7vaqkZfTNo4nNJ3"

    sFilepath = UploadFilePath
    sFileName = GetFilenameFromPath(sFilepath)

    ssPostData1 = vbCrLf & _
    "--" & STR_BOUNDARY & vbCrLf & _
    "Content-Disposition: form-data; name=" & """message""" & vbCrLf & vbCrLf
    arr1 = StrConv(ssPostData1, vbFromUnicode)
    arr2 = (Msg)

    ssPostData2 = vbCrLf & _
    "--" & STR_BOUNDARY & vbCrLf & _
    "Content-Disposition: form-data; name=" & """imageFile""" & "; filename=" & sFileName & vbCrLf & _
    "Content-Type: image/jpng" & vbCrLf & vbCrLf

    arr3 = StrConv(ssPostData2, vbFromUnicode)
    nFile = FreeFile
    Open sFilepath For Binary Access Read As nFile
    If LOF(nFile) > 0 Then
    ReDim baBuffer(0 To LOF(nFile) - 1) As Byte
    Get nFile, , baBuffer
    imagear = baBuffer
    End If
    Close nFile

    arr4 = StrConv(vbCrLf & "--" & STR_BOUNDARY & "--" & vbCrLf, vbFromUnicode)

    Dim arraytotal As Long
    Dim sendarray() As Byte
    arraytotal = UBound(arr1) + UBound(arr2) + UBound(arr3) + UBound(imagear) + UBound(arr4) + 4
    ReDim sendarray(arraytotal)

    For i = 0 To UBound(arr1)
    sendarray(i) = arr1(i)
    Next

    For i = 0 To UBound(arr2)
    sendarray(UBound(arr1) + i + 1) = arr2(i)
    Next

    For i = 0 To UBound(arr3)
    sendarray(UBound(arr1) + UBound(arr2) + i + 2) = arr3(i)
    Next

    For i = 0 To UBound(imagear)
    sendarray(UBound(arr1) + UBound(arr2) + UBound(arr3) + i + 3) = imagear(i)
    Next

    For i = 0 To UBound(arr4)
    sendarray(UBound(arr1) + UBound(arr2) + UBound(arr3) + UBound(imagear) + i + 4) = arr4(i)
    Next

    Set objectXML = CreateObject("Microsoft.XMLHTTP")
    URL = "https://notify-api.line.me/api/notify"
    With objectXML
        .Open "POST", URL, 0
        .setRequestHeader "Content-Type", "multipart/form-data; boundary=" & STR_BOUNDARY
        .setRequestHeader "Authorization", "Bearer " & LineToken
        .send sndar(sendarray)
        Debug.Print .responseText
        End With

    Set objectXML = Nothing
End Sub
Public Function sndar(sendarray As Variant) As Byte()
    sndar = sendarray
End Function
Public Function GetFilenameFromPath(ByVal strPath As String) As String
    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
        
    GetFilenameFromPath = GetFilenameFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
    End If
End Function