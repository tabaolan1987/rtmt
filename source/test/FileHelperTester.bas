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
    On Error GoTo OnError
    Dim source As String
    source = FileHelper.ReadQuery(Constants.END_USER_DATA_CACHE_TABLE_NAME, Constants.Q_CREATE)
    Logger.LogDebug "FileHelperTester.TestReadQuery", source
    mAssert.Equals Len(source) > 0, True, "Source file length > 0"
OnExit:
    ' finally
    Exit Sub
OnError:
    mAssert.Should False, Logger.GetErrorMessage("", Err)
    Logger.LogError "FileHelperTester.TestReadQuery", "", Err
    Resume OnExit
End Sub

Public Sub TestGetCurrentDbPath()
    On Error GoTo OnError
    Dim path As String
    path = FileHelper.CurrentDbPath
    Logger.LogDebug "FileHelperTester.TestGetCurrentDbPath", path
    mAssert.Equals Len(path) > 0, True
OnExit:
    ' finally
    Exit Sub
OnError:
    mAssert.Should False, Logger.GetErrorMessage("", Err)
    Logger.LogError "FileHelperTester.TestGetCurrentDbPath", "", Err
    Resume OnExit
End Sub

Public Sub TestIsExist()
    On Error GoTo OnError
    Dim path As String
    path = FileHelper.CurrentDbPath & Constants.END_USER_DATA_CSV_FILE_PATH
    mAssert.Equals FileHelper.IsExistFile(path), True
    path = FileHelper.CurrentDbPath & "nothing.test"
    mAssert.Equals FileHelper.IsExistFile(path), False
OnExit:
    ' finally
    Exit Sub
OnError:
    mAssert.Should False, Logger.GetErrorMessage("", Err)
    Logger.LogError "FileHelperTester.TestIsExitst", "", Err
    Resume OnExit
End Sub

Public Sub TestReadSSFile()
    On Error GoTo OnError
    Dim source() As String
    source() = FileHelper.ReadSSFile(Constants.SS_SYNC_TABLES)
    
    mAssert.Equals UBound(source) > 0, True
    Dim i As Integer
    For i = LBound(source) To UBound(source)
        Logger.LogDebug "FileHelperTester.TestReadSSFile", source(i)
    Next
OnExit:
    ' finally
    Exit Sub
OnError:
    mAssert.Should False, Logger.GetErrorMessage("", Err)
    Logger.LogError "FileHelperTester.TestReadSSFile", "", Err
    Resume OnExit
End Sub

Public Sub TestTrimSourceFile()
    On Error GoTo OnError
    Dim LineToRemove(2) As Integer
    LineToRemove(0) = 1
    LineToRemove(1) = 2
    LineToRemove(2) = 4
    FileHelper.TrimSourceFile FileHelper.CurrentDbPath & Constants.END_USER_DATA_CSV_TEMPLATE_FILE_PATH, _
                    FileHelper.CurrentDbPath & Constants.END_USER_DATA_CSV_TEMPLATE_TRIM_FILE_PATH, _
                    LineToRemove
OnExit:
    ' finally
    Exit Sub
OnError:
    mAssert.Should False, Logger.GetErrorMessage("", Err)
    Logger.LogError "FileHelperTester.TestTrimSourceFile", "", Err
    Resume OnExit
End Sub

Public Sub TestSaveAsCSV()
    On Error GoTo OnError
    Dim ss As SystemSetting
    ss = Session.Settings()
    Dim inFile As String
    Dim outFile As String
    inFile = FileHelper.CurrentDbPath & Constants.END_USER_DATA_FILE_XLSX
    outFile = FileHelper.CurrentDbPath & Constants.END_USER_DATA_FILE_CSV
    FileHelper.SaveAsCSV inFile, outFile, ss.worksheet
'    mAssert.Equals FileHelper.IsExistFile(outFile), True
OnExit:
    ' finally
    Exit Sub
OnError:
 '   mAssert.Should False, Logger.GetErrorMessage("", Err)
    Logger.LogError "FileHelperTester.TestSaveAsCSV", "", Err
    Resume OnExit
End Sub

Public Sub TestPrepareUserData()
    On Error GoTo OnError
    Dim ss As SystemSetting
    Set ss = Session.Settings()
    Dim inFile As String
    Dim outFile As String
    inFile = FileHelper.CurrentDbPath & Constants.END_USER_DATA_FILE_XLSX
    outFile = FileHelper.PrepareUserData(inFile, ss)
    mAssert.Equals FileHelper.IsExistFile(outFile), True
OnExit:
    ' finally
    Exit Sub
OnError:
    mAssert.Should False, Logger.GetErrorMessage("", Err)
    Logger.LogError "FileHelperTester.TestPrepareUserData", "", Err
    Resume OnExit
End Sub

Public Sub TestGetFileLastModified()
        On Error GoTo OnError
    Dim s As String
    s = FileHelper.FileLastModified(FileHelper.CurrentDbPath & Constants.END_USER_DATA_FILE_XLSX)
    Logger.LogDebug "FileHelperTester.TestGetFileLastModified", "Last modified: " & s
    mAssert.Equals Len(s) <> 0, True
OnExit:
    ' finally
    Exit Sub
OnError:
    mAssert.Should False, Logger.GetErrorMessage("", Err)
    Logger.LogError "FileHelperTester.TestPrepareUserData", "", Err
    Resume OnExit
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.className, "TestReadQuery"
    ITest_Suite.AddTest ITest_Manager.className, "TestGetCurrentDbPath"
    ITest_Suite.AddTest ITest_Manager.className, "TestIsExist"
    ITest_Suite.AddTest ITest_Manager.className, "TestReadSSFile"
    ITest_Suite.AddTest ITest_Manager.className, "TestTrimSourceFile"
    ITest_Suite.AddTest ITest_Manager.className, "TestSaveAsCSV"
    ITest_Suite.AddTest ITest_Manager.className, "TestPrepareUserData"
    ITest_Suite.AddTest ITest_Manager.className, "TestGetFileLastModified"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.MethodName
        Case "TestReadQuery": TestReadQuery
        Case "TestGetCurrentDbPath": TestGetCurrentDbPath
        Case "TestIsExist": TestIsExist
        Case "TestReadSSFile": TestReadSSFile
        Case "TestTrimSourceFile": TestTrimSourceFile
        Case "TestSaveAsCSV": TestSaveAsCSV
        Case "TestPrepareUserData": TestPrepareUserData
        Case "TestGetFileLastModified": TestGetFileLastModified
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.MethodName
    End Select
End Sub