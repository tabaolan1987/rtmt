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

Public Sub TestReadFile()
    Dim source As String
    source = FileHelper.ReadFile(Constants.CREATE_TABLE_END_USER_QUERY)
    'Logger.LogDebug "FileHelperTester.TestReadFile", source
    mAssert.Equals Len(source) > 0, True, "Source file length > 0"
End Sub

Public Sub TestGetCurrentDbPath()
    Dim path As String
    path = FileHelper.CurrentDbPath
    Logger.LogDebug "FileHelperTester.TestGetCurrentDbPath", path
    mAssert.Equals Len(path) > 0, True
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.className, "TestReadFile"
    ITest_Suite.AddTest ITest_Manager.className, "TestGetCurrentDbPath"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.MethodName
        Case "TestReadFile": TestReadFile
        Case "TestGetCurrentDbPath": TestGetCurrentDbPath
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.MethodName
    End Select
End Sub