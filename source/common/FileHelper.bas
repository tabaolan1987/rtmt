'@author Hai Lu
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
    If Dir(strDirPath, vbDirectory) = "" Then
        MkDir strDirPath
    End If
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