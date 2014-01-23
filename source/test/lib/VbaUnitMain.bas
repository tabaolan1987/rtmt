Option Explicit

Public Sub OnTest()
    'Logger.LogDebug "VbaUnitMain.OnTest", StringHelper.GetGUID
    'DoCmd.SetWarnings False
    'Run "ReportingTester"
    Reporting.GenerateReport "End_User_ to_BB_ Job_Role_report"
End Sub

Public Sub Run(Optional TestClassName As String)
    Dim r As TestRunner
    Set r = New TestRunner
    r.Run TestClassName
End Sub

Public Sub Prep(Optional className As String)
    Dim AG As AutoGen
    Set AG = New AutoGen
    AG.Prep className
End Sub

Public Function QW(s As String) As String
    QW = Chr(34) & s & Chr(34)
End Function