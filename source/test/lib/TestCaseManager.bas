Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITestManager

Private mMethodName As String, mClassName As String

Public Sub Init(MethodName As String)
    mMethodName = MethodName
End Sub

Public Property Get MethodName() As String
    MethodName = mMethodName
End Property

Private Property Let ITestManager_ClassName(RHS As String)
    mClassName = RHS
End Property

Private Property Get ITestManager_ClassName() As String
    ITestManager_ClassName = mClassName
End Property

Private Function ITestManager_CountTestCases(Test As ITest) As Long
    ITestManager_CountTestCases = 1
End Function

Private Function ITestManager_Run(Test As ITest, Optional RunManager As IRunManager) As IRunManager
    ' Changed by Hai to control timestamp
    FileHelper.CheckDir FileHelper.CurrentDbPath & "target"
    Dim TestCase As ITestCase, sngStart As Long, sngEnd As Long, sngElapsed As Long
    If RunManager Is Nothing Then Set RunManager = New RunManager
    
    Set TestCase = Test
    RunManager.result.StartTest TestCase
    sngStart = TimerHelper.GetCurrentMillisecond
    
    TestCase.SetUp RunManager.Assert
    TestCase.RunTest
    'debug.print "Run Runtest"
    TestCase.TearDown
    'debug.print "Run teardown"
    sngEnd = TimerHelper.GetCurrentMillisecond
    sngElapsed = sngEnd - sngStart
    RunManager.result.EndTest TestCase, sngElapsed
    'debug.print "Run endtest"
End Function