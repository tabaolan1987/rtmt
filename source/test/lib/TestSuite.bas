Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITest

Private mTests As Collection
Private mManager As TestSuiteManager

Public Function Tests() As Collection
    Set Tests = mTests
End Function

Public Sub AddTest(TestClassName As String, Optional MethodName As String)
    Dim test As ITest
    Dim TestCase As ITestCase
    Dim tl As TestClassLister
    Set tl = New TestClassLister
    Set test = tl.NewTestClass(TestClassName)
    If MethodName <> "" Then
        Set TestCase = test
        TestCase.Manager.Init MethodName
        mTests.Add test
    Else
        mTests.Add test.Suite
    End If
End Sub

Private Sub Class_Initialize()
    Set mTests = New Collection
    Set mManager = New TestSuiteManager
End Sub

Private Property Get ITest_Manager() As ITestManager
    Set ITest_Manager = mManager
End Property

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = Me
End Function