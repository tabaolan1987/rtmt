Option Explicit

Public Sub OnTest()
   ' Run
   ' Dim sh As New SyncHelper
   ' sh.Init "dofa"
   ' sh.sync
   ' sh.Recycle
   'Dim rpm As New ReportMetaData
   'rpm.Init "end_user_to_dofa"
   
   'Reporting.GenerateReport rpm
   'rpm.OpenReport
   'Dim user As New CurrentUser
   'user.AddReportCache Constants.RP_AD_HOC_REPORTING, "C:/test/test.xml"
   'user.AddReportCache Constants.RP_COURSE_ANALYTICS, "C:/test/test1.xml"
   'Dim v As String
   'v = user.GetReportCache(Constants.RP_COURSE_ANALYTICS)
   'user.RemoveReportCache (Constants.RP_COURSE_ANALYTICS)
    Dim rpm As ReportMetaData
    Set rpm = Session.ReportMetaData(Constants.RP_END_USER_TO_BB_JOB_ROLE)
    Reporting.GenerateReport (rpm)
    rpm.OpenReport
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