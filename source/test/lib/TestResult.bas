Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Private mAssert As Assert
Private mCurrentTestCase As ITestCase
Private mTestCasesRun As Long
Private mFailures As Collection
Private mTestCaseRows As Collection
Public Property Get CurrentTestCase() As ITestCase
    Set CurrentTestCase = mCurrentTestCase
End Property

Public Sub StartTest(TestCase As ITestCase)
    Set mCurrentTestCase = TestCase
    mTestCasesRun = mTestCasesRun + 1
End Sub

Public Property Get TestCasesRun() As Long
    TestCasesRun = mTestCasesRun
End Property

Public Sub EndTest(TestCase As ITestCase, time As Long)
    Dim TCR As TestCaseRow
    Set TCR = New TestCaseRow
    TCR.Init TestCase, time
    TestCaseRows.Add TCR
    Set mCurrentTestCase = Nothing
End Sub

Public Sub AddFailure(TestCase As ITestCase, Comment As String)
    Dim Failure As TestFailure
    Set Failure = New TestFailure
    Failure.Init TestCase, Comment
    Failures.Add Failure
End Sub

Public Property Get Failures() As Collection
    Set Failures = mFailures
End Property

Public Property Get TestCaseRows() As Collection
    If mTestCaseRows Is Nothing Then
       Set mTestCaseRows = New Collection
    End If
    Set TestCaseRows = mTestCaseRows
End Property

Public Property Get TotalTime(className As String) As Long
    Dim total As Long, TC As TestCaseRow, test As ITest
    total = 0
    For Each TC In mTestCaseRows
        Set test = TC.TestCase
        If StrComp(className, test.Manager.className, vbTextCompare) = 0 Then
            total = total + TC.time
        End If
    Next
    TotalTime = total
End Property

Public Property Get TestCaseCount(className As String) As Integer
    Dim Count As Integer, TC As TestCaseRow, test As ITest
    Count = 0
    For Each TC In mTestCaseRows
        Set test = TC.TestCase
        If StrComp(className, test.Manager.className, vbTextCompare) = 0 Then
            Count = Count + 1
        End If
    Next
    TestCaseCount = Count
End Property

Public Property Get FailureCount(className As String) As Integer
    Dim Count As Integer, FL As TestFailure, test As ITest, TC As ITestCase
    Count = 0
    For Each FL In mFailures
        Set TC = FL.TestCase
       Set test = TC
        If StrComp(className, test.Manager.className, vbTextCompare) = 0 Then
            Count = Count + 1
        End If
   Next
   FailureCount = Count
End Property

Public Property Get isFailures(TestCase As ITestCase) As TestFailure
    Dim FL As TestFailure, TC As ITestCase, check As Boolean, returnVal As TestFailure, test As ITest, TE As ITest
    Dim s1 As String, s2 As String
    Set TE = TestCase
    For Each FL In mFailures
        Set TC = FL.TestCase
        Set test = TC
        s1 = TC.Manager.MethodName & "." & test.Manager.className
        s2 = TestCase.Manager.MethodName & "." & TE.Manager.className
       ' Debug.Print s1 & " | "; s2
        If StrComp(s1, s2, vbTextCompare) = 0 Then
            Set returnVal = FL
            check = True
        End If
    Next
    If check = False Then
        Set returnVal = Nothing
    End If
    Set isFailures = returnVal
End Property

Private Sub Class_Initialize()
    Set mFailures = New Collection
End Sub

Public Property Get WasSuccessful() As Boolean
    WasSuccessful = Failures.Count = 0
End Property