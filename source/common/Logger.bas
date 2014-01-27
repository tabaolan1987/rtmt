' @author Hai Lu
' System logger
Option Compare Database
Private Const L_INFO = 1
Private Const L_DEBUG = 2
Private Const L_ERROR = 3
Private logLevelCfg As String

Private Sub GetLogLevelCfg()
    If Len(logLevelCfg) = 0 Then
        Dim s As SystemSetting
        Set s = Session.Settings()
        logLevelCfg = s.LogLevel
        Debug.Print "Log level config: " & logLevelCfg
    End If
End Sub


Public Sub LogInfo(caller As String, log As String)
    CallLog caller, L_INFO, log
End Sub

Public Sub LogDebug(caller As String, log As String)
    CallLog caller, L_DEBUG, log
End Sub

Public Function GetErrorMessage(message As String, errX As ErrObject) As String
    GetErrorMessage = "Number: " _
                                & CStr(errX.Number) _
                                & ", Description: " _
                                & errX.Description _
                                & ", Help Context: " _
                                & errX.HelpContext _
                                & ". Message: " & message
End Function

Public Sub LogError(caller As String, log As String, errX As ErrObject)
    If errX Is Nothing Then
        CallLog caller, L_ERROR, log
    Else
        CallLog caller, L_ERROR, GetErrorMessage(log, errX)
        errX.Clear
    End If
End Sub

Private Sub CallLog(caller As String, lvl As Integer, log As String)
    GetLogLevelCfg
    If StringHelper.IsEqual(logLevelCfg, "INFO", True) And lvl = L_DEBUG Then GoTo QuitLog
    If StringHelper.IsEqual(logLevelCfg, "ERROR", True) And lvl < L_ERROR Then GoTo QuitLog
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
QuitLog:
    Exit Sub
End Sub