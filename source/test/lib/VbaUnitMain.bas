Option Explicit

Public Sub OnTest()
    'Run
   ' Dim sh As New SyncHelper
   ' sh.init "privileges"
   ' sh.sync
   ' sh.Recycle
    Dim dbm As New DbManager
    dbm.init
    dbm.ExecuteQuery "update activity set ActivityGroup='testing' where id='98748e5e-3eda-452d-9c74-65bc1d1582dc'"
    dbm.Recycle
    
    
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