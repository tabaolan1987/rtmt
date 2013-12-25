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