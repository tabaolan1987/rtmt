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

Public Sub TestImportData()
    Dim csvPath As String
    csvPath = FileHelper.CurrentDbPath & Constants.END_USER_DATA_CSV_FILE_PATH
    Dim im As DbManager: Set im = New DbManager
    im.Init
    im.ImportData csvPath
    im.Recycle
End Sub

Public Sub TestImportSqlTable()
    Dim dbm As DbManager, _
        SyncTables() As String, _
        prop As SystemSettings, _
        isEmpty As Boolean, _
        stTable As String
    Set dbm = New DbManager
    Set prop = New SystemSettings
    prop.Init
    SyncTables = prop.SyncTables
    isEmpty = Ultilities.IsVarArrayEmpty(SyncTables)
    If isEmpty = False Then
        Dim i As Integer
        For i = LBound(SyncTables) To UBound(SyncTables)
            stTable = Trim(SyncTables(i))
            Logger.LogDebug "Form_Main.btnSync_Click", "Start sync table: " & stTable
            dbm.ImportSqlTable prop.ServerName & "," & prop.Port, prop.DatabaseName, stTable, stTable, prop.Username, prop.Password
        Next i
    End If
End Sub

Public Sub TestExecuteQuery()
    Dim im As DbManager: Set im = New DbManager
    im.Init
    If Ultilities.ifTableExists(Constants.END_USER_DATA_TABLE_NAME) = False Then
        im.ExecuteQuery FileHelper.ReadQuery(Constants.END_USER_DATA_TABLE_NAME, Constants.Q_CREATE)
    End If
End Sub

Public Sub TestOpenRecordSet()
    Dim params As New Scripting.Dictionary
    Dim dm As DbManager: Set dm = New DbManager
    Dim rInfo As ReportMetaData: Set rInfo = New ReportMetaData
    dm.Init
    rInfo.Init (name)
    params.Add "SYSTEM_ROLE_NAME", "Procurement Catalogue Approver"
    params.Add "BP_ROLE_STANDARD_NAME", "POQR Approver"
    If rInfo.Valid = True Then
        dm.OpenRecordSet rInfo.query, params
    End If
    dm.Recycle
End Sub

Public Sub TestSyncUserData()
    On Error GoTo OnError
    Dim dm As DbManager: Set dm = New DbManager
    dm.Init
    dm.SyncUserData
    dm.Recycle
OnExit:
    ' finally
    Exit Sub
OnError:
    Logger.LogError "DbManagerTester.TestSyncUserData", "Error TestSyncUserData", Err
    Resume OnExit
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.className, "TestImportData"
    ITest_Suite.AddTest ITest_Manager.className, "TestImportSqlTable"
    ITest_Suite.AddTest ITest_Manager.className, "TestExecuteQuery"
    ITest_Suite.AddTest ITest_Manager.className, "TestSyncUserData"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.MethodName
        Case "TestImportData": TestImportData
        Case "TestImportSqlTable": TestImportSqlTable
        Case "TestExecuteQuery": TestExecuteQuery
        Case "TestSyncUserData": TestSyncUserData
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.MethodName
    End Select
End Sub