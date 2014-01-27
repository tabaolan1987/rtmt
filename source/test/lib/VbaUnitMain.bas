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
    Dim dbm As DbManager, _
        SyncTables() As String, _
        prop As SystemSetting, _
        isEmpty As Boolean, _
        stTable As String
    Set dbm = New DbManager
    Set prop = Session.Settings()
    SyncTables = prop.SyncTables
    isEmpty = Ultilities.IsVarArrayEmpty(SyncTables)
    If isEmpty = False Then
        Dim i As Integer
        dbm.RecycleTableName Constants.TABLE_SYNC_CONFLICT
        For i = LBound(SyncTables) To UBound(SyncTables)
            stTable = Trim(SyncTables(i))
            Logger.LogDebug "DbManagerTester.TestSyncTable", "Start sync table: " & stTable
            dbm.ExecuteQuery "update [" & stTable & "] set timestamp=null"
        Next i
    End If
    
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