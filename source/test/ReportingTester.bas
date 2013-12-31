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

Public Sub TestExportReport()
    Dim check As Boolean: check = False
    On Error GoTo OnError
        Dim testTempXlsx As String: testTempXlsx = FileHelper.CurrentDbPath & Constants.END_USER_DATA_REPORTING_TEMPLATE
        Dim output As String: output = FileHelper.CurrentDbPath & Constants.END_USER_DATA_REPORTING_OUTPUT_DIR
        FileHelper.CheckDir output
        output = output & "/" & Constants.END_USER_DATA_REPORTING_OUTPUT_FILE
        FileHelper.Delete (output)
        Reporting.ExportExcelReport "select * from tblImport", testTempXlsx, output, "Role Mapping Template", "A5"
        check = FileHelper.IsExist(output)
        
OnExit:
    mAssert.Equals check, True
    Exit Sub
OnError:
    Logger.LogError "Reporting.TestExportReport", "Error when export report", Err
    Resume OnExit
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.className, "TestExportReport"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.MethodName
        Case "TestExportReport": TestExportReport
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.MethodName
    End Select
End Sub