Option Explicit

Public Sub OnTest()
   ' Run
   ' Dim sh As New SyncHelper
   ' sh.Init "user_data"
   ' sh.sync
   ' sh.Recycle
<<<<<<< HEAD
   Dim dh As New DofaHelper
   dh.Init
   dh.ImportDofa
=======
   Dim rpm As New ReportMetaData
   rpm.Init "end_user_to_bb_job_role_report_qualification"
   Reporting.GenerateReport rpm
   rpm.OpenReport
>>>>>>> 8d0059633564530c5abf9599127ee9de1f99a7f6
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