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
    Set mAssertAsResultUser.result = New TestResult
End Sub

Private Property Get IRunManager_Assert() As IAssert
    Set IRunManager_Assert = mAssert
End Property

Private Sub IRunManager_Report()
    Dim Failure As TestFailure, RM As IRunManager, test As ITest, TestCase As ITestCase, TCR As TestCaseRow
        ' Added by Hai to create Unit test report
        Dim fso As Object, oFile As Object, reportPath As String, lastClass As String
        lastClass = ""
        reportPath = Application.CurrentProject.path & "\target"
        FileHelper.CheckDir (reportPath)
        reportPath = reportPath & "\reports"
        FileHelper.CheckDir (reportPath)
        Logger.LogInfo "RunManager.IRunManager_Report", "Tests run: " & result.TestCasesRun & " Failures: " & result.Failures.Count
        For Each TCR In result.TestCaseRows
            Set Failure = result.isFailures(TCR.TestCase)
            Set TestCase = TCR.TestCase
            Set test = TestCase
            If Not StrComp(lastClass, test.Manager.className, vbTextCompare) = 0 Then
                If Not fso Is Nothing Then
                    oFile.WriteLine "</testsuite>"
                    oFile.Close
                    Set fso = Nothing
                    Set oFile = Nothing
                End If
                lastClass = test.Manager.className
                Set fso = CreateObject("Scripting.FileSystemObject")
                Set oFile = fso.CreateTextFile(reportPath & "\" & lastClass & ".xml")
                oFile.WriteLine "<?xml version=""1.0"" encoding=""UTF-8""?>"
                oFile.WriteLine "<testsuite name=""" & lastClass & """ time=""" & TimerHelper.MsToString(result.TotalTime(lastClass)) & """ errors=""0"" tests=""" & CStr(result.TestCaseCount(lastClass)) & """ failures=""" & CStr(result.FailureCount(lastClass)) & """>"
            End If
            
           Logger.LogInfo "RunManager.IRunManager_Report", test.Manager.className & "." & TestCase.Manager.MethodName & ": " & TCR.Time
            
            oFile.WriteLine "   <testcase time=""" & TimerHelper.MsToString(TCR.Time) & """ name=""" & TestCase.Manager.MethodName & """ >"
            If Not Failure Is Nothing Then
                Logger.LogError "RunManager.IRunManager_Report", " #Failure -> " & Failure.Comment, Nothing
                oFile.WriteLine "       <failure type=""runtime"" message=""" & StringHelper.EncodeXml(Failure.Comment) & """>"
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

End Sub

Private Property Get IRunManager_Result() As TestResult
    Set IRunManager_Result = result
End Property

Public Property Get result() As TestResult
    Set result = mAssertAsResultUser.result
End Property