Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @author Hai Lu
' Report meta data object
Option Explicit

Private mType As String
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

Private mMergeEnable As Boolean
Private mMergeColumes As Collection
Private mMergePrimary As Long
Private mCount As Long
Private mBulkSize As Long
Private mSkipCheckHeader As Boolean
Private mCustomMode As Boolean

Private mComplete As Boolean

Public Function Init(Name As String, Optional ss As SystemSetting, Optional rpType As String)
    If Len(rpType) > 0 Then
        mType = rpType
    Else
        mType = Constants.RP_TYPE_DEFAULT
    End If
    rawName = Name
    Logger.LogDebug "ReportMetaData.Init", "Start init report meta name: " & rawName
    Dim tmpRawSection() As String, tmpStr As String, i As Integer
    Dim v As Variant
    Dim tmpList() As String
    Dim rpSection As ReportSection
    Dim ir As New IniReader
    If StringHelper.IsEqual(mType, Constants.RP_TYPE_DEFAULT, True) Then
        mQueryFilePath = FileHelper.DuplicateAsTemporary(FileHelper.CurrentDbPath & Constants.RP_ROOT_FOLDER & Name & Constants.FILE_EXTENSION_QUERY)
    Else
        mQueryFilePath = FileHelper.DuplicateAsTemporary(FileHelper.CurrentDbPath & Constants.RP_ROOT_FOLDER & Name & Constants.FILE_EXTENSION_QUERY_MAPPING)
    End If
    
    mTemplateFilePath = FileHelper.DuplicateAsTemporary(FileHelper.CurrentDbPath & Constants.RP_ROOT_FOLDER & Name & Constants.FILE_EXTENSION_TEMPLATE)
    mConfigFilePath = FileHelper.DuplicateAsTemporary(FileHelper.CurrentDbPath & Constants.RP_ROOT_FOLDER & Name & Constants.FILE_EXTENSION_CONFIG)
    Logger.LogDebug "ReportMetaData.Init", "Read configuration path: " & mConfigFilePath
    ir.Init mConfigFilePath
    
    mName = ir.ReadKey(Constants.SECTION_GENERAL, Constants.KEY_NAME)
    Logger.LogDebug "ReportMetaData.Init", "Report name: " & mName
    
    mWorkSheet = ir.ReadKey(Constants.SECTION_GENERAL, Constants.KEY_WORK_SHEET)
    Logger.LogDebug "ReportMetaData.Init", "Work sheet: " & mWorkSheet
    mBulkSize = ir.ReadLongKey(Constants.SECTION_FORMAT, Constants.KEY_BULK_SIZE)
    
    mFillHeader = ir.ReadBooleanKey(Constants.SECTION_FORMAT, Constants.KEY_FILL_HEADER)
    Logger.LogDebug "ReportMetaData.Init", "Fill header: " & CStr(mFillHeader)
    mSkipCheckHeader = ir.ReadBooleanKey(Constants.SECTION_FORMAT, Constants.KEY_SKIP_CHECK_HEADER)
    mMergeEnable = ir.ReadBooleanKey(Constants.SECTION_FORMAT, Constants.KEY_MERGE_ENABLE)
    mCustomMode = ir.ReadBooleanKey(Constants.SECTION_FORMAT, Constants.KEY_CUSTOM_MODE)
    Logger.LogDebug "ReportMetaData.Init", "Merge enable: " & CStr(mMergeEnable)
    If mMergeEnable Then
        mMergePrimary = ir.ReadKey(Constants.SECTION_FORMAT, Constants.KEY_MERGE_PRIMARY)
        Logger.LogDebug "ReportMetaData.Init", "Merge primary: " & CStr(mMergePrimary)
        tmpStr = ir.ReadKey(Constants.SECTION_FORMAT, Constants.KEY_MERGE_COLUMES)
        If Len(tmpStr) <> 0 Then
            tmpList = Split(tmpStr, ",")
            Set mMergeColumes = New Collection
            For Each v In tmpList
                mMergeColumes.Add CInt(Trim(CStr(v)))
            Next v
        End If
    End If
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
            rpSection.Init tmpStr, ss, mSkipCheckHeader
            If StringHelper.IsEqual(rpSection.SectionType, Constants.RP_SECTION_TYPE_FIXED, True) _
                Or StringHelper.IsEqual(rpSection.SectionType, Constants.RP_SECTION_TYPE_TMP_TABLE, True) Then
                mCount = rpSection.Count
            End If
            If Not rpSection.Valid Then
                mValid = False
            End If
            mReportSections.Add rpSection
        Next i
    End If
    mComplete = False
End Function

Public Property Get OutputPath() As String
    If Len(mOutputPath) = 0 Then
        Dim tmpDir As String
        tmpDir = FileHelper.tmpDir
        FileHelper.CheckDir tmpDir
        mOutputPath = tmpDir & StringHelper.GetGUID & Constants.FILE_EXTENSION_REPORT
    End If
    OutputPath = mOutputPath
End Property

Public Function OpenReport()
    Dim oExcel As New Excel.Application
    Dim WB As New Excel.Workbook
    Dim WS As Excel.WorkSheet
    Dim rng As Excel.range
    With oExcel
        .Visible = True
        If .CommandBars("Ribbon").Height >= 150 Then
            oExcel.SendKeys "^{F1}"
        End If
        Set WB = .Workbooks.Open(mOutputPath)
    End With
End Function

Public Function OpenSaveAs()
    Dim oExcel As New Excel.Application
    Dim WB As New Excel.Workbook
    Dim WS As Excel.WorkSheet
    Dim rng As Excel.range
    With oExcel
        .Visible = False
        .GetSaveAsFilename InitialFileName:=mName, _
    fileFilter:="Excel Files (*.xlsx), *.xlsx"
    End With
End Function

Public Function Wait()
    ' Waiting forever :)
    WaitForFileClose mOutputPath, 0, 0
End Function

Public Function Recyle()
    Dim dbm As New DbManager
    dbm.Init
    Dim rps As ReportSection
    If Not mReportSections Is Nothing Then
        For Each rps In mReportSections
            If Len(rps.CachedTable) > 0 Then
                dbm.DeleteTable rps.CachedTable
            End If
        Next rps
    End If
    dbm.Recycle
    FileHelper.DeleteFile mQueryFilePath
    FileHelper.DeleteFile mTemplateFilePath
    FileHelper.DeleteFile mConfigFilePath
    FileHelper.DeleteFile mOutputPath
End Function

Public Property Get ReportSections() As Collection
    If mReportSections Is Nothing Then
        Set mReportSections = New Collection
    End If
    Set ReportSections = mReportSections
End Property

Public Property Get query() As String
    query = mQuery
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

Public Property Get MergeEnable() As Boolean
    MergeEnable = mMergeEnable
End Property

Public Property Get CustomMode() As Boolean
    CustomMode = mCustomMode
End Property

Public Property Get MergeColumes() As Collection
    Set MergeColumes = mMergeColumes
End Property

Public Property Get MergePrimary() As Long
    MergePrimary = mMergePrimary
End Property

Public Property Get Count() As Long
    Count = mCount
End Property

Public Function SetComplete(done As Boolean)
    mComplete = done
End Function

Public Property Get Complete() As Boolean
    Complete = mComplete
End Property

Public Property Get SkipCheckHeader() As Boolean
    SkipCheckHeader = mSkipCheckHeader
End Property

Public Property Get BulkSize() As Long
    BulkSize = mBulkSize
End Property