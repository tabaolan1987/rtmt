Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ITestManager
Private mClassName

Private Function ITestManager_CountTestCases(test As ITest) As Long
    Dim count As Long
    Dim TestSuite As TestSuite, Tests As Collection
    Dim SubTest As ITest
    Set TestSuite = test
    Set Tests = TestSuite.Tests()
    For Each SubTest In Tests
        count = count + SubTest.Manager.CountTestCases(SubTest)
    Next
    ITestManager_CountTestCases = count
End Function

Private Property Let ITestManager_ClassName(RHS As String)
    mClassName = RHS
End Property

Private Property Get ITestManager_ClassName() As String
    ITestManager_ClassName = mClassName
End Property

Private Function ITestManager_Run(test As ITest, Optional RunManager As IRunManager) As IRunManager
    Dim TestSuite As TestSuite, Tests As Collection
    Dim SubTest As ITest
    Set TestSuite = test
    Set Tests = TestSuite.Tests()
    If RunManager Is Nothing Then Set RunManager = New RunManager
    For Each SubTest In TestSuite.Tests()
        SubTest.Manager.Run SubTest, RunManager
    Next
    Set ITestManager_Run = RunManager
End Function