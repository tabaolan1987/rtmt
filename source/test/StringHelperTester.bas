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
        Dim Test As Integer
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

Public Sub TestEncodeURL()
    Dim src As String
    src = "It's me & nothing"
    mAssert.Equals "It%27s%20me%20%26%20nothing", StringHelper.EncodeURL(src)
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

Public Sub TestCompareStringDate()
    mAssert.Equals TimerHelper.Compare("1/20/2014 5:12:34 AM", "1/20/2014 11:12:34 PM"), -1, "1/20/2014 5:12:34 AM VS 1/20/2014 11:12:34 PM"
    mAssert.Equals TimerHelper.Compare("10/20/2014 5:12:34 AM", "12/20/2014 11:12:34 AM"), -1, "10/20/2014 5:12:34 AM VS 12/20/2014 11:12:34 AM"
    mAssert.Equals TimerHelper.Compare("1/20/2016 5:12:34 AM", "1/20/2015 11:12:34 PM"), 1, "1/20/2016 5:12:34 AM VS 1/20/2015 11:12:34 PM"
    mAssert.Equals TimerHelper.Compare("11/20/2014 11:12:34 PM", "11/20/2014 11:12:34 PM"), 0, "11/20/2014 11:12:34 PM VS 11/20/2014 11:12:34 PM"
    mAssert.Equals TimerHelper.Compare("1/20/2014 5:12:34 AM", "1/20/2014 5:12:34 AM"), 0, "1/20/2014 5:12:34 AM VS 1/20/2014 5:12:34 AM"
    mAssert.Equals TimerHelper.Compare("2/17/2014 5:12:34 AM", "1/20/2014 11:12:34 PM"), 1, "2/17/2014 5:12:34 AM VS 1/20/2014 11:12:34 PM"
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.className, "TestEncodeXml"
    ITest_Suite.AddTest ITest_Manager.className, "TestEncodeURL"
    ITest_Suite.AddTest ITest_Manager.className, "TestIsEqual"
    ITest_Suite.AddTest ITest_Manager.className, "TestIsContain"
    ITest_Suite.AddTest ITest_Manager.className, "TestEndsWith"
    ITest_Suite.AddTest ITest_Manager.className, "TestStartsWith"
    ITest_Suite.AddTest ITest_Manager.className, "TestCompareStringDate"
    
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.MethodName
        Case "TestEncodeXml": TestEncodeXml
        Case "TestEncodeURL": TestEncodeURL
        Case "TestIsEqual": TestIsEqual
        Case "TestIsContain": TestIsContain
        Case "TestEndsWith": TestEndsWith
        Case "TestStartsWith": TestStartsWith
        Case "TestCompareStringDate": TestCompareStringDate
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.MethodName
    End Select
End Sub