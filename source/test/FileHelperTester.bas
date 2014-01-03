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

Public Sub TestReadQuery()
    Dim source As String
    source = FileHelper.ReadQuery(Constants.END_USER_DATA_TABLE_NAME, Constants.Q_CREATE)
    Logger.LogDebug "FileHelperTester.TestReadQuery", source
    mAssert.Equals Len(source) > 0, True, "Source file length > 0"
End Sub

Public Sub TestGetCurrentDbPath()
    Dim path As String
    path = FileHelper.CurrentDbPath
    Logger.LogDebug "FileHelperTester.TestGetCurrentDbPath", path
    mAssert.Equals Len(path) > 0, True
End Sub

Public Sub TestIsExist()
    Dim path As String
    path = FileHelper.CurrentDbPath & Constants.END_USER_DATA_CSV_FILE_PATH
    mAssert.Equals FileHelper.IsExist(path), True
    path = FileHelper.CurrentDbPath & "nothing.test"
    mAssert.Equals FileHelper.IsExist(path), False
End Sub

Public Sub TestReadSSFile()
    Dim source() As String
    source() = FileHelper.ReadSSFile(Constants.SS_SYNC_TABLES)
    
    mAssert.Equals UBound(source) > 0, True
    Dim i As Integer
    For i = LBound(source) To UBound(source)
        Logger.LogDebug "FileHelperTester.TestReadSSFile", source(i)
    Next
End Sub

Public Sub TestTrimSourceFile()
    Dim LineToRemove(2) As Integer
    LineToRemove(0) = 1
    LineToRemove(1) = 2
    LineToRemove(2) = 4
    FileHelper.TrimSourceFile FileHelper.CurrentDbPath & Constants.END_USER_DATA_CSV_TEMPLATE_FILE_PATH, _
                    FileHelper.CurrentDbPath & Constants.END_USER_DATA_CSV_TEMPLATE_TRIM_FILE_PATH, _
                    LineToRemove
End Sub

Public Sub TestSaveAsCSV()
    Dim inFile As String
    Dim outFile As String
    inFile = FileHelper.CurrentDbPath & Constants.END_USER_DATA_FILE_XLSX
    outFile = FileHelper.CurrentDbPath & Constants.END_USER_DATA_FILE_CSV
    FileHelper.SaveAsCSV inFile, outFile
    mAssert.Equals FileHelper.IsExist(inFile), True
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.className, "TestReadQuery"
    ITest_Suite.AddTest ITest_Manager.className, "TestGetCurrentDbPath"
    ITest_Suite.AddTest ITest_Manager.className, "TestIsExist"
    ITest_Suite.AddTest ITest_Manager.className, "TestReadSSFile"
    ITest_Suite.AddTest ITest_Manager.className, "TestTrimSourceFile"
    ITest_Suite.AddTest ITest_Manager.className, "TestSaveAsCSV"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.MethodName
        Case "TestReadQuery": TestReadQuery
        Case "TestGetCurrentDbPath": TestGetCurrentDbPath
        Case "TestIsExist": TestIsExist
        Case "TestReadSSFile": TestReadSSFile
        Case "TestTrimSourceFile": TestTrimSourceFile
        Case "TestSaveAsCSV": TestSaveAsCSV
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.MethodName
    End Select
End Sub