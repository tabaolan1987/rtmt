' Author Hai Lu
' To control Date & Time
Option Explicit
Private Type SystemTime
    Year As Integer
    Month As Integer
    DayOfWeek As Integer
    Day As Integer
    Hour As Integer
    Minute As Integer
    Second As Integer
    Milliseconds As Integer
End Type
#If VBA7 Then
    Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
    Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (st As SystemTime)
#Else
    Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
    Private Declare Sub GetSystemTime Lib "kernel32" (st As SystemTime)
#End If



Function GetCurrentMillisecond() As Long
    Dim CurrentTime As SystemTime
    GetSystemTime CurrentTime
    GetCurrentMillisecond = Hour(Now) * 3600000 + Minute(Now) * 60000 + Second(Now) * 1000 + CurrentTime.Milliseconds
End Function

Function MsToString(ms As Long) As String
    Dim sec As Integer, last As Integer
    If ms > 999 Then
        sec = Int(ms / 1000)
        last = ms - 1000 * sec
        MsToString = CStr(sec) & "." & Format(last, "000")
    Else
        MsToString = "0." & Format(ms, "000")
    End If
End Function