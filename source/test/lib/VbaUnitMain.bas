Option Explicit

Public Sub OnTest()
    'Logger.LogDebug "VbaUnitMain.OnTest", StringHelper.GetGUID
    'DoCmd.SetWarnings False
    'Run "MappingHelperTester"
   ' Run "UserManagementTester"
   ' Run "DbManagerTester"
    'Dim um As New UserManagement
    'um.CheckConflict
    'Reporting.GenerateReport Constants.RP_ROLE_MAPPING_OUTPUT_OF_TOOL_FOR_SECURITY
    'Session.Init
    
   ' Logger.LogDebug "test", Session.currentUser.Valid
    'Dim rpmd As New ReportMetaData
    'rpmd.Init Constants.RP_END_USER_TO_BB_JOB_ROLE
   ' Reporting.GenerateReport Constants.RP_ROLE_MAPPING_OUTPUT_OF_TOOL_FOR_SECURITY
   
    'FileHelper.Zip "D:\test.zip", "D:\GIT\rolemapping-repo\data\images"
    Dim s As String
    s = FileHelper.FileSaveDialog("test.xlsx")
    Logger.LogDebug "test", s
    
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