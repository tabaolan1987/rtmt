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

Public Sub TestSomethingThatFails()
    TimerHelper.Sleep 767
    mAssert.Equals 2, 1
End Sub

Public Sub TestSomethingThatPasses()
    TimerHelper.Sleep 431
    mAssert.Equals 2, 1 + 1
    
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.className, "TestSomethingThatFails"
    ITest_Suite.AddTest ITest_Manager.className, "TestSomethingThatPasses"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.MethodName
        Case "TestSomethingThatFails": TestSomethingThatFails
        Case "TestSomethingThatPasses": TestSomethingThatPasses
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.MethodName
    End Select
End Sub