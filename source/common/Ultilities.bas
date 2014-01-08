'@author Hai Lu
' General utilities function
Option Compare Database

Public Sub MakeAccde()
    Dim sourcedb As String, targetdb As String
    sourcedb = FileHelper.CurrentDbPath & "target\rolemapping.accdb"
    targetdb = FileHelper.CurrentDbPath & "target\rolemapping.accde"
    Logger.LogDebug "Ultilities.MakeAccde", "source db:" & sourcedb
    Logger.LogDebug "Ultilities.MakeAccde", "target db:" & targetdb
    
    Dim AccessApplication As New Access.Application
    
    With AccessApplication
        .AutomationSecurity = 1 'MsoAutomationSecurityLow
        .UserControl = True
        .SysCmd 603, sourcedb, targetdb 'this makes the ACCDE file
        .Quit
    End With
End Sub

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