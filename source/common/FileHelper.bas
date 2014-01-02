'@author Hai Lu
' Work with File & Dir
Option Explicit
Const ForReading = 1

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

Public Function readFile(strFilePath As String) As String
   Dim nSourceFile As Integer, sText As String
   Close
   ''Get the number of the next free text file
   nSourceFile = FreeFile
   ''Write the entire file to sText
   Open CurrentDbPath & strFilePath For Input As #nSourceFile
        sText = Input$(LOF(1), 1)
   Close
   readFile = sText
End Function

Function CurrentDbPath() As String
    Dim cRes As String
    Dim nPos As Long
    cRes = CurrentDb.name
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
    Dim arraySize As Integer
    Dim sInput As String
    Dim check As Boolean
    Dim i As Long
    Dim tmpList() As String
    Dim ln As String
    If IsExist(path) Then
        Dim FSO As Object
        Dim readFile As Object
        Set FSO = CreateObject("Scripting.FileSystemObject")
        Set readFile = FSO.OpenTextFile(path, ForReading, False)
        Do Until readFile.AtEndOfStream = True
            ln = Trim(readFile.ReadLine)
            If StringHelper.StartsWith(ln, "#", True) = False And Len(ln) <> 0 Then
                ReDim Preserve tmpList(arraySize)
                tmpList(arraySize) = ln
                arraySize = arraySize + 1
            End If
        Loop
        readFile.Close
        Set FSO = Nothing
        Set readFile = Nothing
    End If
    
    ReadSSFile = tmpList
End Function

Public Function TrimSourceFile(fileToRead As String, fileToWrite As String, LineToRemove() As Integer)
    Dim FSO As Object
    Dim readFile As Object
    Dim writeFile As Object
    Dim repLine As Variant
    Dim ln As String
    Dim l As Long
    Dim tmpList() As String
    Dim arraySize As Integer
    Dim check As Boolean
    Dim ltm As Variant
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set readFile = FSO.OpenTextFile(fileToRead, ForReading, False)
    Set writeFile = FSO.CreateTextFile(fileToWrite, True, False)
    
    '# iterate the array and do the replacement line by line
    Do Until readFile.AtEndOfStream = True
        ln = readFile.ReadLine
        check = False
        For Each ltm In LineToRemove
            If (ltm = l + 1) Then
            '    Logger.LogDebug "FileHelper.TrimSourceFile", "Line to remove " & CStr(ltm)
                check = True
            End If
        Next
        If check = False Then
           ' Logger.LogDebug "FileHelper.TrimSourceFile", "Readline " & CStr(l) & " . Text: " & ln
            ReDim Preserve tmpList(arraySize)
            tmpList(arraySize) = ln
            arraySize = arraySize + 1
        End If
        l = l + 1
    Loop
    readFile.Close
    '# Write to the array items to the file
    writeFile.Write Join(tmpList, vbNewLine)
    writeFile.Close
    
    '# clean up
    Set readFile = Nothing
    Set writeFile = Nothing
    Set FSO = Nothing
End Function