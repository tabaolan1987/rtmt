' @author Hai Lu
' System logger
Option Compare Database
Private Const L_INFO = 1
Private Const L_DEBUG = 2
Private Const L_ERROR = 3

Public Sub LogInfo(caller As String, log As String)
    CallLog caller, L_INFO, log
End Sub

Public Sub LogDebug(caller As String, log As String)
    CallLog caller, L_DEBUG, log
End Sub

Public Sub LogError(caller As String, log As String, errX As ErrObject)
    If errX Is Nothing Then
        CallLog caller, L_ERROR, log
    Else
        
        CallLog caller, L_ERROR, "Number: " _
                                & CStr(errX.Number) _
                                & ", Description: " _
                                & errX.Description _
                                & ", Help Context: " _
                                & errX.HelpContext _
                                & ". Message: " & log
        errX.Clear
    End If
End Sub

Private Sub CallLog(caller As String, lvl As Integer, log As String)
    Dim fn As Integer, logPath As String, line As String
    fn = FreeFile
    logPath = FileHelper.CurrentDbPath & "logs"
    FileHelper.CheckDir logPath
    logPath = logPath & "\" & Format(Now, "yyyy-MM-dd") & ".log"
    Open logPath For Append As #fn
    line = Format(Now, "MM-dd-yyyy hh:mm:ss")
    Select Case lvl
        Case L_INFO:  line = line & " | INFO  > "
        Case L_DEBUG: line = line & " | DEBUG > "
        Case L_ERROR: line = line & " | ERROR > "
        Case Else:    line = line & " | OTHER > "
    End Select
    line = line & caller & ": " & log
    Debug.Print line
    Print #fn, line
    Close #fn
End Sub