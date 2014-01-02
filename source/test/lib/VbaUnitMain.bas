Option Explicit

Public Sub OnTest()
    Run '"AutoGenTester"
End Sub

Public Sub Run(Optional TestClassName As String)
    Dim R As TestRunner
    Set R = New TestRunner
    R.Run TestClassName
End Sub

Public Sub Prep(Optional className As String)
    Dim AG As AutoGen
    Set AG = New AutoGen
    AG.Prep className
End Sub

Public Function QW(s As String) As String
    QW = Chr(34) & s & Chr(34)
End Function