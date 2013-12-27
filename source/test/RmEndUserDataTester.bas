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
    
End Sub

Public Sub TestCreateTable()
    Dim im As RmEndUserData: Set im = New RmEndUserData
    im.CreateTable
End Sub

Public Sub TestDropTable()
    Dim im As RmEndUserData: Set im = New RmEndUserData
    im.DropTable
End Sub

Public Sub TestDeleteAllData()
    Dim im As RmEndUserData: Set im = New RmEndUserData
    im.DeleteAllData
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.className, "TestCreateTable"
    ITest_Suite.AddTest ITest_Manager.className, "TestImportData"
    ITest_Suite.AddTest ITest_Manager.className, "TestDeleteAllData"
    ITest_Suite.AddTest ITest_Manager.className, "TestDropTable"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.MethodName
        Case "TestImportData": TestImportData
        Case "TestCreateTable": TestCreateTable
        Case "TestDropTable": TestDropTable
        Case "TestDeleteAllData": TestDeleteAllData
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.MethodName
    End Select
End Sub