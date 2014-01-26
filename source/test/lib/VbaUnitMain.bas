Option Explicit

Public Sub OnTest()
    'Logger.LogDebug "VbaUnitMain.OnTest", StringHelper.GetGUID
    'DoCmd.SetWarnings False
    'Run "MappingHelperTester"
   ' Run "UserManagementTester"
    'Run "DbManagerTester"
    'Dim um As New UserManagement
    'um.CheckConflict
    Reporting.GenerateReport "ROLE_MAPPING_OUTPUT_OF_TOOL_FOR_SECURITY"
    
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