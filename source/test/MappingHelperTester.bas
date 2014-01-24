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

Public Sub TestMappingMetaData()
    On Error GoTo OnError
    Dim mmd As New MappingMetadata
    mmd.Init Constants.MAPPING_ACTIVITIES_SPECIALISM
    Dim data As New Scripting.Dictionary
    data.Add Constants.Q_KEY_ID, "test id 1"
    data.Add Constants.Q_KEY_ID_LEFT, "test id left 1"
    data.Add Constants.Q_KEY_ID_TOP, "test id top 1"
    data.Add Constants.Q_KEY_FUNCTION_REGION_ID, "test funct"
    data.Add Constants.Q_KEY_REGION_NAME, "GoM"
    data.Add Constants.Q_KEY_CHECK, "true"
    mmd.Query Constants.Q_CHECK, data
    mmd.Query Constants.Q_CREATE, data
    mmd.Query Constants.Q_UPDATE, data
    mmd.Query Constants.Q_TOP, data
    mmd.Query Constants.Q_LEFT, data
OnExit:
    ' finally
    Exit Sub
OnError:
    mAssert.Should False, Logger.GetErrorMessage("", Err)
    Logger.LogError "UserManagementTeser.TestCheckDuplicate", "", Err
    Resume OnExit
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.className, "TestMappingMetaData"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.MethodName
        Case "TestMappingMetaData": TestMappingMetaData
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.MethodName
    End Select
End Sub