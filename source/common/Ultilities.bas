'@author Hai Lu
' General utilities function
Option Compare Database

Public Function ifTableExists(tblName As String) As Boolean
    ifTableExists = False
    If DCount("[Name]", "MSysObjects", "[Name] = '" & tblName & "'") = 1 Then
    ifTableExists = True
    End If
End Function

Function IsVarArrayEmpty(anArray As Variant)

    Dim i As Integer
    
    On Error Resume Next
        i = UBound(anArray, 1)
    If Err.Number = 0 Then
        IsVarArrayEmpty = False
    Else
        IsVarArrayEmpty = True
    End If

End Function