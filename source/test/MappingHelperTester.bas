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
    Dim mmd As New MappingMetaData
    mmd.Init Constants.MAPPING_ACTIVITIES_SPECIALISM
    Dim data As New Scripting.Dictionary
    data.Add Constants.Q_KEY_ID, "test id 1"
    data.Add Constants.Q_KEY_ID_LEFT, "test id left 1"
    data.Add Constants.Q_KEY_ID_TOP, "test id top 1"
    data.Add Constants.Q_KEY_FUNCTION_REGION_ID, "test funct"
    data.Add Constants.Q_KEY_REGION_NAME, "GoM"
    data.Add Constants.Q_KEY_CHECK, "true"
    mmd.query Constants.Q_CHECK, data
    mmd.query Constants.Q_CREATE, data
    mmd.query Constants.Q_UPDATE, data
    mmd.query Constants.Q_TOP, data
    mmd.query Constants.Q_LEFT, data
    mmd.Recyle
OnExit:
    ' finally
    Exit Sub
OnError:
    mAssert.Should False, Logger.GetErrorMessage("", Err)
    Logger.LogError "MappingHelperTester.TestMappingMetaData", "", Err
    Resume OnExit
End Sub

Public Sub TestInitMapping()
    On Error GoTo OnError
    Dim mmd As New MappingMetaData
    Dim ss As SystemSetting
    Set ss = Session.Settings()
    mmd.Init Constants.MAPPING_ACTIVITIES_SPECIALISM
    
    Dim mh As New MappingHelper
    mh.Init mmd, ss
    mh.GenerateMapping
    'mh.OpenMapping
    mh.ParseMapping
    mmd.Recyle
OnExit:
    ' finally
    Exit Sub
OnError:
    mAssert.Should False, Logger.GetErrorMessage("", Err)
    Logger.LogError "UserManagementTeser.TestInitMapping", "", Err
    Resume OnExit
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.className, "TestMappingMetaData"
    ITest_Suite.AddTest ITest_Manager.className, "TestInitMapping"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.MethodName
        Case "TestMappingMetaData": TestMappingMetaData
        Case "TestInitMapping": TestInitMapping
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.MethodName
    End Select
End Sub