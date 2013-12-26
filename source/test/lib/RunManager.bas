Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IRunManager

Private mAssert As Assert
Private mAssertAsResultUser As IResultUser

Private Sub Class_Initialize()
    Set mAssert = New Assert
    Set mAssertAsResultUser = mAssert
    Set mAssertAsResultUser.Result = New TestResult
End Sub

Private Property Get IRunManager_Assert() As IAssert
    Set IRunManager_Assert = mAssert
End Property

Private Sub IRunManager_Report()
    Dim Failure As TestFailure, RM As IRunManager, Test As ITest, TestCase As ITestCase, TCR As TestCaseRow
    
    If Result.WasSuccessful Then
        Debug.Print "OK (" & Result.TestCasesRun & ")"
    Else
        ' Added by Hai to create Unit test report
        Dim fso As Object, oFile As Object, reportPath As String, lastClass As String
        lastClass = ""
        reportPath = Application.CurrentProject.Path & "\target"
        FileHelper.CheckDir (reportPath)
        reportPath = reportPath & "\reports"
        FileHelper.CheckDir (reportPath)
        Debug.Print "Tests run: " & Result.TestCasesRun & " Failures: " & Result.Failures.count
        For Each TCR In Result.TestCaseRows
            Set Failure = Result.isFailures(TCR.TestCase)
            Set TestCase = TCR.TestCase
            Set Test = TestCase
            If Not StrComp(lastClass, Test.Manager.className, vbTextCompare) = 0 Then
                If Not fso Is Nothing Then
                    oFile.WriteLine "</testsuite>"
                    oFile.Close
                    Set fso = Nothing
                    Set oFile = Nothing
                End If
                lastClass = Test.Manager.className
                Set fso = CreateObject("Scripting.FileSystemObject")
                Set oFile = fso.CreateTextFile(reportPath & "\" & lastClass & ".xml")
                oFile.WriteLine "<?xml version=""1.0"" encoding=""UTF-8""?>"
                oFile.WriteLine "<testsuite name=""" & lastClass & """ time=""" & TimerHelper.MsToString(Result.TotalTime(lastClass)) & """ errors=""0"" tests=""" & CStr(Result.TestCaseCount(lastClass)) & """ failures=""" & CStr(Result.FailureCount(lastClass)) & """>"
            End If
            
            Debug.Print Test.Manager.className & "." & TestCase.Manager.MethodName & ": " & TCR.Time
            
            oFile.WriteLine "   <testcase time=""" & TimerHelper.MsToString(TCR.Time) & """ name=""" & TestCase.Manager.MethodName & """ >"
            If Not Failure Is Nothing Then
                oFile.WriteLine "       <failure type=""runtime"" message=""" & StringHelper.encodeXml(Failure.Comment) & """>"
                oFile.WriteLine "           " & Failure.Comment
                oFile.WriteLine "       </failure>"
                oFile.WriteLine "       <system-out>"
                oFile.WriteLine "           " & Failure.Comment
                oFile.WriteLine "       </system-out>"
            End If
            oFile.WriteLine "   </testcase>"
        Next
        If Not oFile Is Nothing Then
            oFile.WriteLine "</testsuite>"
            oFile.Close
            Set fso = Nothing
            Set oFile = Nothing
        End If
    End If
End Sub

Private Property Get IRunManager_Result() As TestResult
    Set IRunManager_Result = Result
End Property

Public Property Get Result() As TestResult
    Set Result = mAssertAsResultUser.Result
End Property