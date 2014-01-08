Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @author Hai Lu
' Report meta data object
Option Explicit
Private mName As String
Private mQueryType As String
Private mQuery As String
Private mStartCol As Long
Private mStartRow As Long
Private mValid As Boolean
Private mConfigFilePath As String
Private mQueryFilePath As String
Private mTemplateFilePath As String

Private mReportSections As Collection

Public Function Init(Name As String)
    Logger.LogDebug "ReportMetaData.Init", "Start init report meta name: " & Name
    Dim tmpRawSection() As String, tmpStr As String, i As Integer
    Dim rpSection As ReportSection
    Dim ir As New IniReader
    
    mQueryFilePath = FileHelper.CurrentDbPath & Constants.RP_ROOT_FOLDER & Name & Constants.RP_QUERY_FILE_EXTENSION
    mTemplateFilePath = FileHelper.CurrentDbPath & Constants.RP_ROOT_FOLDER & Name & Constants.RP_TEMPLATE_FILE_EXTENSION
    mConfigFilePath = FileHelper.CurrentDbPath & Constants.RP_ROOT_FOLDER & Name & Constants.RP_CONFIG_FILE_EXTENSION
    Logger.LogDebug "ReportMetaData.Init", "Read configuration path: " & mConfigFilePath
    ir.Init mConfigFilePath
    
    mName = ir.ReadKey(Constants.SECTION_GENERAL, Constants.KEY_NAME)
    Logger.LogDebug "ReportMetaData.Init", "Report name: " & mName
    mQueryType = ir.ReadKey(Constants.SECTION_GENERAL, Constants.KEY_QUERY_TYPE)
    Logger.LogDebug "ReportMetaData.Init", "Query type: " & mQueryType
    
    mStartCol = CLng(ir.ReadKey(Constants.SECTION_FORMAT, Constants.KEY_START_COL))
    Logger.LogDebug "ReportMetaData.Init", "Start column: " & mStartCol
    mStartRow = CLng(ir.ReadKey(Constants.SECTION_FORMAT, Constants.KEY_START_ROW))
    Logger.LogDebug "ReportMetaData.Init", "Start row: " & mStartRow
    
    Logger.LogDebug "ReportMetaData.Init", "Read query path: " & mQueryFilePath
    mValid = False
    mQuery = FileHelper.ReadFileFullPath(mQueryFilePath)
    
    If Len(mQuery) <> 0 And StringHelper.IsEqual(mQueryType, Constants.RP_QUERY_TYPE_SECTION, True) Then
        Set mReportSections = New Collection
        mValid = True
        tmpRawSection = Split(mQuery, Constants.RP_SPLIT_LEVEL_1)
        For i = LBound(tmpRawSection) To UBound(tmpRawSection)
            Logger.LogDebug "ReportMetaData.Init", "Found section " & CStr(i + 1)
            Set rpSection = New ReportSection
            tmpStr = Trim(tmpRawSection(i))
            rpSection.Init tmpStr
            If Not rpSection.Valid Then
                mValid = False
            End If
            mReportSections.Add rpSection
        Next i
    End If
End Function

Public Property Get ReportSections() As Collection
    If mReportSections Is Nothing Then
        Set mReportSections = New Collection
    End If
    Set ReportSections = mReportSections
End Property

Public Property Get Query() As String
    Query = mQuery
End Property

Public Property Get Valid() As Boolean
    Valid = mValid
End Property

Public Property Get Name() As String
    Name = mName
End Property

Public Property Get QueryType() As String
    QueryType = mQueryType
End Property

Public Property Get StartRow() As Long
    StartRow = mStartRow
End Property

Public Property Get StartCol() As Long
    StartCol = mStartCol
End Property