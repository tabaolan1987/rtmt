Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITest
Implements ITestCase

Private mManager As TestCaseManager
Private mAssert As IAssert

Private Sub Class_Initialize()
    Set mManager = New TestCaseManager
End Sub

Private Property Get ITestCase_Manager() As TestCaseManager
    Set ITestCase_Manager = mManager
End Property

Private Property Get ITest_Manager() As ITestManager
    Set ITest_Manager = mManager
End Property

Private Sub ITestCase_SetUp(Assert As IAssert)
    Set mAssert = Assert
End Sub

Private Sub ITestCase_TearDown()

End Sub

Public Sub TestImportSettings()
    On Error GoTo OnError
    Dim s As SystemSetting
    Set s = Session.Settings()
    'mAssert.Equals s.ServerName, "CMGSRV2\SQLEXPRESS"
    'mAssert.Equals s.Port, "1433"
    'mAssert.Equals s.DatabaseName, "upstream_role_mapping"
    'mAssert.Equals s.userNAme, "sa"
    'mAssert.Equals s.Password, "admincmg@3f"
    'mAssert.Equals s.RegionName, "Indo"
    'mAssert.Equals s.LogLevel, "DEBUG"
    'mAssert.Equals UBound(s.LineToRemove), 2
   ' mAssert.Equals UBound(s.SyncTables), 9
   ' mAssert.Equals s.SyncUsers.Count > 0, True
    Dim dic As Scripting.Dictionary, i As Integer
    Set dic = s.SyncUsers
    For i = 0 To dic.count - 1
        Logger.LogDebug "SystemSettingsTester.TestImportSettings", "key: " & dic.keys(i) & " | value: " & dic.Items(i)
    Next i
OnExit:
    ' finally
    Exit Sub
OnError:
    'mAssert.Should False, Logger.GetErrorMessage("", Err)
    Logger.LogError "SystemSettingsTester.TestImportSettings", "", Err
    Resume OnExit
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.className, "TestImportSettings"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.MethodName
        Case "TestImportSettings": TestImportSettings
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.MethodName
    End Select
End Sub