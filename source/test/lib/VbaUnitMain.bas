Option Explicit

Public Sub OnTest()
   ' Run
   ' Dim sh As New SyncHelper
   ' sh.Init "BlueprintRole_mapping_BpRole"
   ' sh.sync
   ' sh.Recycle
   Dim test As New Scripting.Dictionary
   test.Add "llag", "t1"
   test.Add "llag2", "t3"
   MsgBox StringHelper.GenerateFilterDict(test, True)
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