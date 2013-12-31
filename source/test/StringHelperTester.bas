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

Public Sub TestIsEqual()
    On Error GoTo OnError
        Dim checker As Boolean
        Dim test As Integer
        checker = StringHelper.IsEqual("a", "A", True)
        mAssert.Equals checker, True, "ignoreCase = True"
        checker = StringHelper.IsEqual("a", "A", False)
        mAssert.Equals checker, False, "ignoreCase = False"
OnExit:
    Exit Sub
OnError:
    Logger.LogError "StringHelperTester.TestIsEqual", "error when test is equal", Err
    Resume OnExit
End Sub

Public Sub TestEncodeXml()
    Dim src As String
    src = "&""'<>"
    mAssert.Equals "&amp;&quot;&apos;&lt;&gt;", StringHelper.EncodeXml(src)
End Sub

Public Sub TestIsContain()
    Dim checker As Boolean
    On Error GoTo OnError
        checker = StringHelper.IsContain("one is nothing ONl true", "onl", True)
        mAssert.Equals checker, True, "ignoreCase = True"
        checker = StringHelper.IsContain("one is nothing ONl true", "onl", False)
        mAssert.Equals checker, False, "ignoreCase = False"
OnExit:
    ' finally
    Exit Sub
OnError:
    Logger.LogError "StringHelperTester.TestIsContain", "error when test is contain", Err
    Resume OnExit
End Sub

Public Sub TestEndsWith()
    Dim checker As Boolean
    checker = StringHelper.EndsWith("test", "ST", True)
    mAssert.Equals checker, True, "ignoreCase = True"
    'checker = StringHelper.EndsWith("test", "ST", False)
    'mAssert.Equals checker, True, "ignoreCase = False"
End Sub

Public Sub TestStartsWith()
    Dim checker As Boolean
    checker = StringHelper.StartsWith("test", "Te", True)
    mAssert.Equals checker, True, "ignoreCase = True"
    'checker = StringHelper.StartsWith("test", "Te", False)
    'mAssert.Equals checker, True, "ignoreCase = False"
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.className, "TestEncodeXml"
    ITest_Suite.AddTest ITest_Manager.className, "TestIsEqual"
    ITest_Suite.AddTest ITest_Manager.className, "TestIsContain"
    ITest_Suite.AddTest ITest_Manager.className, "TestEndsWith"
    ITest_Suite.AddTest ITest_Manager.className, "TestStartsWith"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.MethodName
        Case "TestEncodeXml": TestEncodeXml
        Case "TestIsEqual": TestIsEqual
        Case "TestIsContain": TestIsContain
        Case "TestEndsWith": TestEndsWith
        Case "TestStartsWith": TestStartsWith
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.MethodName
    End Select
End Sub