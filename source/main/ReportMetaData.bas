Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @author Hai Lu
' Report meta data object
Option Explicit
Private rawName As String
Private mName As String
Private mWorkSheet As String

Private mQuery As String

Private mFillHeader As Boolean
Private mStartHeaderRow As Long
Private mStartHeaderCol As Long
Private mStartCol As Long
Private mStartRow As Long
Private mValid As Boolean
Private mConfigFilePath As String
Private mQueryFilePath As String
Private mTemplateFilePath As String

Private mOutputPath As String
Private mReportSections As Collection

Public Property Get OutputPath() As String
    If Len(mOutputPath) = 0 Then
        Dim tmpDir As String
        tmpDir = FileHelper.CurrentDbPath & Constants.RP_DEFAULT_OUTPUT_FOLDER
        FileHelper.CheckDir tmpDir
        mOutputPath = tmpDir & "\" & rawName & Constants.FILE_EXTENSION_REPORT
    End If
    OutputPath = mOutputPath
End Property

Public Function Recyle()
    FileHelper.Delete mQueryFilePath
    FileHelper.Delete mTemplateFilePath
    FileHelper.Delete mConfigFilePath
End Function

Public Function Init(Name As String, Optional ss As SystemSettings)
    rawName = Name
    Logger.LogDebug "ReportMetaData.Init", "Start init report meta name: " & rawName
    Dim tmpRawSection() As String, tmpStr As String, i As Integer
    Dim rpSection As ReportSection
    Dim ir As New IniReader
    
    mQueryFilePath = FileHelper.DuplicateAsTemporary(FileHelper.CurrentDbPath & Constants.RP_ROOT_FOLDER & Name & Constants.FILE_EXTENSION_QUERY)
    mTemplateFilePath = FileHelper.DuplicateAsTemporary(FileHelper.CurrentDbPath & Constants.RP_ROOT_FOLDER & Name & Constants.FILE_EXTENSION_TEMPLATE)
    mConfigFilePath = FileHelper.DuplicateAsTemporary(FileHelper.CurrentDbPath & Constants.RP_ROOT_FOLDER & Name & Constants.FILE_EXTENSION_CONFIG)
    Logger.LogDebug "ReportMetaData.Init", "Read configuration path: " & mConfigFilePath
    ir.Init mConfigFilePath
    
    mName = ir.ReadKey(Constants.SECTION_GENERAL, Constants.KEY_NAME)
    Logger.LogDebug "ReportMetaData.Init", "Report name: " & mName
    
    mWorkSheet = ir.ReadKey(Constants.SECTION_GENERAL, Constants.KEY_WORK_SHEET)
    Logger.LogDebug "ReportMetaData.Init", "Work sheet: " & mWorkSheet
    
    
    mFillHeader = ir.ReadBooleanKey(Constants.SECTION_FORMAT, Constants.KEY_FILL_HEADER)
    Logger.LogDebug "ReportMetaData.Init", "Fill header: " & CStr(mFillHeader)
    
    mStartHeaderCol = ir.ReadLongKey(Constants.SECTION_FORMAT, Constants.KEY_START_HEADER_COL)
    Logger.LogDebug "ReportMetaData.Init", "Start header column: " & CStr(mStartHeaderCol)
    mStartHeaderRow = ir.ReadLongKey(Constants.SECTION_FORMAT, Constants.KEY_START_HEADER_ROW)
    Logger.LogDebug "ReportMetaData.Init", "Start header row: " & CStr(mStartHeaderRow)
    mStartCol = ir.ReadLongKey(Constants.SECTION_FORMAT, Constants.KEY_START_COL)
    Logger.LogDebug "ReportMetaData.Init", "Start column: " & CStr(mStartCol)
    mStartRow = ir.ReadLongKey(Constants.SECTION_FORMAT, Constants.KEY_START_ROW)
    Logger.LogDebug "ReportMetaData.Init", "Start row: " & CStr(mStartRow)
    
    Logger.LogDebug "ReportMetaData.Init", "Read query path: " & mQueryFilePath
    mValid = False
    mQuery = FileHelper.ReadFileFullPath(mQueryFilePath)
    
    If Len(mQuery) <> 0 Then
        Set mReportSections = New Collection
        mValid = True
        tmpRawSection = Split(mQuery, Constants.SPLIT_LEVEL_1)
        For i = LBound(tmpRawSection) To UBound(tmpRawSection)
            Logger.LogDebug "ReportMetaData.Init", "Found section " & CStr(i + 1)
            Set rpSection = New ReportSection
            tmpStr = Trim(tmpRawSection(i))
            rpSection.Init tmpStr, ss
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

Public Property Get StartRow() As Long
    StartRow = mStartRow
End Property

Public Property Get StartCol() As Long
    StartCol = mStartCol
End Property

Public Property Get WorkSheet() As String
    WorkSheet = mWorkSheet
End Property

Public Property Get FillHeader() As Boolean
    FillHeader = mFillHeader
End Property

Public Property Get StartHeaderRow() As Long
    StartHeaderRow = mStartHeaderRow
End Property

Public Property Get StartHeaderCol() As Long
    StartHeaderCol = mStartHeaderCol
End Property

Public Property Get TemplateFilePath() As String
    TemplateFilePath = mTemplateFilePath
End Property