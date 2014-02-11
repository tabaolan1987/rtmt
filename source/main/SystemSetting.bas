Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @author Hai Lu
' To read & write Settings.ini
Option Explicit

Private mValidatorURL As String
Private mToken As String
Private mBulkSize As Long
Private mNtidField As String
Private mServerName As String
Private mDatabaseName As String
Private mPort As String
Private mUsername As String
Private mPassword As String
Private mSyncTables() As String
Private mSyncUsers As Scripting.Dictionary
Private mValidatorMapping As Scripting.Dictionary
Private mWorkSheet As String
Private mLineToRemove() As Integer
Private mTableNames() As String
Private mRegionName As String
Private mRegionFunctionId As String
Private mLogLevel As String
Private mEnableTesting As Boolean
Private mEnableValidation As Boolean
Private mTestNtid As String
Private mEnv As String

Public Function Init()
    Dim ir As IniReader: Set ir = Ultilities.SystemIniReader
    mServerName = ir.ReadKey(Constants.SECTION_REMOTE_DATABASE, Constants.KEY_SERVER_NAME)
    mDatabaseName = ir.ReadKey(Constants.SECTION_REMOTE_DATABASE, Constants.KEY_DATABASE_NAME)
    mPort = ir.ReadKey(Constants.SECTION_REMOTE_DATABASE, Constants.KEY_PORT)
    mUsername = ir.ReadKey(Constants.SECTION_REMOTE_DATABASE, Constants.KEY_USERNAME)
    mPassword = ir.ReadKey(Constants.SECTION_REMOTE_DATABASE, Constants.KEY_PASSWORD)
    
    mSyncTables = FileHelper.ReadSSFile(Constants.SS_SYNC_TABLES)
    
    Dim source As String, tmpList() As String, ln As String, arraySize As Integer, i As Integer
    source = ir.ReadKey(Constants.SECTION_USER_DATA, Constants.KEY_LINE_TO_REMOVE)
    tmpList = Split(source, ",")
    For i = LBound(tmpList) To UBound(tmpList)
        ln = Trim(tmpList(i))
        If Len(ln) <> 0 Then
            ReDim Preserve mLineToRemove(arraySize)
            mLineToRemove(arraySize) = CInt(ln)
            arraySize = arraySize + 1
        End If
    Next
    Dim tl() As String
    Set mSyncUsers = New Scripting.Dictionary
    tmpList = FileHelper.ReadSSFile(Constants.SS_SYNC_USERS)
    For i = LBound(tmpList) To UBound(tmpList)
        ln = Trim(tmpList(i))
        If Len(ln) <> 0 Then
            tl = Split(ln, ":")
            mSyncUsers.Add Trim(tl(0)), Trim(tl(1))
        End If
    Next
    mRegionName = ir.ReadKey(Constants.SECTION_USER_DATA, Constants.KEY_REGION_NAME)
    mRegionFunctionId = ir.ReadKey(Constants.SECTION_USER_DATA, Constants.KEY_REGION_FUNCTION_ID)
    source = ir.ReadKey(Constants.SECTION_USER_DATA, Constants.KEY_TABLE_NAME)
    mTableNames = Split(source, ",")
    mWorkSheet = ir.ReadKey(Constants.SECTION_USER_DATA, Constants.KEY_WORK_SHEET)
    mLogLevel = ir.ReadKey(Constants.SECTION_APPLICATION, Constants.KEY_LOG_LEVEL)
    mTestNtid = ir.ReadKey(Constants.SECTION_APPLICATION, Constants.KEY_TEST_NTID)
    mEnv = ir.ReadKey(Constants.SECTION_APPLICATION, Constants.KEY_ENV)
    
    mEnableTesting = ir.ReadBooleanKey(Constants.SECTION_APPLICATION, Constants.KEY_ENABLE_TESTING)
    mEnableValidation = ir.ReadBooleanKey(Constants.SECTION_USER_DATA, Constants.KEY_ENABLE_VALIDATION)
    mValidatorURL = ir.ReadKey(Constants.SECTION_USER_DATA, Constants.KEY_VALIDATOR_URL)
    mToken = ir.ReadKey(Constants.SECTION_USER_DATA, Constants.KEY_TOKEN)
    mBulkSize = ir.ReadLongKey(Constants.SECTION_USER_DATA, Constants.KEY_BULK_SIZE)
    mNtidField = ir.ReadKey(Constants.SECTION_USER_DATA, Constants.KEY_NTID_FIELD)
    Set mValidatorMapping = New Scripting.Dictionary
    tmpList = FileHelper.ReadSSFile(Constants.SS_VALIDATOR_MAPPING)
    For i = LBound(tmpList) To UBound(tmpList)
        ln = Trim(tmpList(i))
        If Len(ln) <> 0 Then
            tl = Split(ln, ":")
            mValidatorMapping.Add Trim(tl(0)), Trim(tl(1))
        End If
    Next
End Function

Public Property Get ServerName() As String
    ServerName = mServerName
End Property

Public Property Get DatabaseName() As String
    DatabaseName = mDatabaseName
End Property

Public Property Get Port() As String
    Port = mPort
End Property

Public Property Get userNAme() As String
    userNAme = mUsername
End Property

Public Property Get Password() As String
    Password = mPassword
End Property

Public Property Get SyncTables() As String()
    SyncTables = mSyncTables
End Property

Public Property Get SyncUsers() As Scripting.Dictionary
    Set SyncUsers = mSyncUsers
End Property

Public Property Get LineToRemove() As Integer()
    LineToRemove = mLineToRemove
End Property

Public Property Get TableNames() As String()
    TableNames = mTableNames
End Property

Public Property Get RegionName() As String
    If Len(mRegionName) <> 0 Then
        RegionName = mRegionName
    Else
        RegionName = Session.CurrentUser.FuncRegion.Region
    End If
End Property

Public Property Get LogLevel() As String
    LogLevel = mLogLevel
End Property

Public Property Get ValidatorURL() As String
    ValidatorURL = mValidatorURL
End Property

Public Property Get Token() As String
    Token = mToken
End Property

Public Property Get BulkSize() As Long
    BulkSize = mBulkSize
End Property

Public Property Get validatorMapping() As Scripting.Dictionary
    Set validatorMapping = mValidatorMapping
End Property

Public Property Get NtidField() As String
    NtidField = mNtidField
End Property

Public Property Get WorkSheet() As String
    WorkSheet = mWorkSheet
End Property

Public Property Get RegionFunctionId() As String
    If Len(mRegionFunctionId) <> 0 Then
        RegionFunctionId = mRegionFunctionId
    Else
        RegionFunctionId = Session.CurrentUser.FuncRegion.FuncRgID
    End If
End Property

Public Property Get EnableTesting() As Boolean
    EnableTesting = mEnableTesting
End Property

Public Property Get EnableValidation() As Boolean
    EnableValidation = mEnableValidation
End Property

Public Property Get TestNtid() As String
    TestNtid = mTestNtid
End Property

Public Property Get Env() As String
    Env = mEnv
End Property