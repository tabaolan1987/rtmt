Option Explicit

Public Sub OnTest()
    'Run '"FileHelperTester"
    Dim rmd As New ReportMetaData
    rmd.Init Constants.RP_END_USER_TO_BB_JOB_ROLE
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