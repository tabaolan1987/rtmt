'@author Hai Lu
' Work with File & Dir
Option Compare Database

Function GetCSVFile() As String
    Dim fDialog As Object
    Set fDialog = Application.FileDialog(3)
    With fDialog
        .Title = "Select the CSV file to import"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv", 1
        .Filters.Add "All Files", "*.*", 2
        If .Show = True Then
            GetCSVFile = .SelectedItems(1)
        End If
    End With
    Set fDialog = Nothing
End Function

Public Sub CheckDir(strDirPath As String)
    strDirPath = Replace(strDirPath, "/", "\")
    Dim listStr() As String, strTemp As String, Str As String
    listStr = Split(strDirPath, "\")
    strTemp = ""
    Dim i As Integer
    For i = LBound(listStr) To UBound(listStr)
       strTemp = strTemp & listStr(i) & "\"
        If Dir(strTemp, vbDirectory) = "" Then
            MkDir strTemp
        End If
    Next i
End Sub

Public Function ReadFile(strFilePath As String) As String
   Dim nSourceFile As Integer, sText As String
   Close
   ''Get the number of the next free text file
   nSourceFile = FreeFile
   ''Write the entire file to sText
   Open CurrentDbPath & strFilePath For Input As #nSourceFile
        sText = Input$(LOF(1), 1)
   Close
   ReadFile = sText
End Function

Function CurrentDbPath() As String
    Dim cRes As String
    Dim nPos As Long
    cRes = CurrentDb.Name
    nPos = Len(cRes)
    Do Until Right(cRes, 1) = "\"
        nPos = nPos - 1
        cRes = Left(cRes, nPos)
    Loop
    CurrentDbPath = cRes
End Function

Function IsExist(path As String) As Boolean
    IsExist = Dir(path) <> ""
End Function

Function Delete(path As String) As Boolean
    If IsExist(path) Then
        SetAttr path, vbNormal
        Kill path
        Delete = True
    Else
        Delete = False
    End If
End Function

Public Function ReadSSFile(path As String) As String()
    Dim tmpList() As String
    Dim arraySize As Integer
    Dim sInput As String
    Dim check As Boolean
    Dim i As Long
    If IsExist(path) Then
        Open path For Input As #1
        Do While Not EOF(1)
            Input #1, sInput
            If StringHelper.StartsWith(sInput, "#", True) = False And Len(sInput) <> 0 Then
                ReDim Preserve tmpList(arraySize)
                tmpList(arraySize) = sInput
                arraySize = arraySize + 1
                'tmpList(UBound(tmpList) - 1) = sInput
            End If
        Loop
    End If
    ReadSSFile = tmpList
End Function