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
Private mWorksheet As String

Private mQuery As String

Private mFillHeader As Boolean
Private mFillCategory As Boolean
Private mStartCategoryRow As Long
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

Private mPivotTable As Boolean
Private mPivotTableName As String
Private mPivotTableWorksheet As String
Private mPivotWordWrapCols As Collection

Private mDateColumes As Collection
Private mDateFormat As String

Private mReportSheets As Scripting.Dictionary

Private mComplete As Boolean

Private mLastModified As String

Private mCustomFilter As Scripting.Dictionary

Public Function SetCustomFilter(mFilter As Scripting.Dictionary)
    Set mCustomFilter = mFilter
End Function

Public Function CustomFilter() As Scripting.Dictionary
    If mCustomFilter Is Nothing Then
        Set mCustomFilter = New Scripting.Dictionary
        mCustomFilter.Add "CUSTOM_FILTER_NAME", " is not null "
        mCustomFilter.Add "CUSTOM_FILTER_ID", " is not null "
    End If
    Set CustomFilter = mCustomFilter
End Function

Public Function Init(Name As String, Optional ss As SystemSetting, Optional rpType As String)
    If Len(rpType) > 0 Then
        mType = rpType
    Else
        mType = Constants.RP_TYPE_DEFAULT
    End If
    rawName = Name
    Logger.LogDebug "ReportMetaData.Init", "Start init report meta name: " & rawName
    Dim mReportSect As Collection
    Dim tmpRawSection() As String, tmpStr As String, i As Integer
    Dim tmpRawSheet() As String
    Dim tmpRawSheetSection() As String
    Dim v As Variant
    Dim tmpList() As String
    Dim rpSection As ReportSection
    Dim ir As New IniReader
    Dim j As Integer
    If StringHelper.IsEqual(mType, Constants.RP_TYPE_DEFAULT, True) Then
        mQueryFilePath = FileHelper.DuplicateAsTemporary(FileHelper.CurrentDbPath & Constants.RP_ROOT_FOLDER & Name & Constants.FILE_EXTENSION_QUERY)
        mTemplateFilePath = FileHelper.DuplicateAsTemporary(FileHelper.CurrentDbPath & Constants.RP_ROOT_FOLDER & Name & Constants.FILE_EXTENSION_TEMPLATE)
        mConfigFilePath = FileHelper.DuplicateAsTemporary(FileHelper.CurrentDbPath & Constants.RP_ROOT_FOLDER & Name & Constants.FILE_EXTENSION_CONFIG)
    Else
        mQueryFilePath = FileHelper.DuplicateAsTemporary(FileHelper.CurrentDbPath & Constants.RP_ROOT_FOLDER & Name & Constants.FILE_EXTENSION_QUERY_MAPPING)
        mTemplateFilePath = FileHelper.DuplicateAsTemporary(FileHelper.CurrentDbPath & Constants.RP_ROOT_FOLDER & Name & Constants.FILE_EXTENSION_TEMPLATE_MAPPING)
        mConfigFilePath = FileHelper.DuplicateAsTemporary(FileHelper.CurrentDbPath & Constants.RP_ROOT_FOLDER & Name & Constants.FILE_EXTENSION_CONFIG_MAPPING)
    End If
    
    
    
    Logger.LogDebug "ReportMetaData.Init", "Read configuration path: " & mConfigFilePath
    ir.Init mConfigFilePath
    
    mName = ir.ReadKey(Constants.SECTION_GENERAL, Constants.KEY_NAME)
    Logger.LogDebug "ReportMetaData.Init", "Report name: " & mName
    
    mWorksheet = ir.ReadKey(Constants.SECTION_GENERAL, Constants.KEY_WORK_SHEET)
    Logger.LogDebug "ReportMetaData.Init", "Work sheet: " & mWorksheet
    mBulkSize = ir.ReadLongKey(Constants.SECTION_FORMAT, Constants.KEY_BULK_SIZE)
    
    mPivotTable = ir.ReadBooleanKey(Constants.SECTION_GENERAL, Constants.KEY_PIVOT_TABLE)
    mPivotTableName = ir.ReadKey(Constants.SECTION_GENERAL, Constants.KEY_PIVOT_TABLE_NAME)
    mPivotTableWorksheet = ir.ReadKey(Constants.SECTION_GENERAL, Constants.KEY_PIVOT_TABLE_WORK_SHEET)
    
    mFillHeader = ir.ReadBooleanKey(Constants.SECTION_FORMAT, Constants.KEY_FILL_HEADER)
    mFillCategory = ir.ReadBooleanKey(Constants.SECTION_FORMAT, Constants.KEY_FILL_CATEGORY)
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
    Set mDateColumes = New Collection
    tmpStr = ir.ReadKey(Constants.SECTION_FORMAT, Constants.KEY_DATE_COLS)
    If Len(tmpStr) <> 0 Then
            tmpList = Split(tmpStr, ",")
            For Each v In tmpList
                mDateColumes.Add CInt(Trim(CStr(v)))
            Next v
    End If
    
    mDateFormat = ir.ReadKey(Constants.SECTION_FORMAT, Constants.KEY_DATE_FORMAT)
    
    If mPivotTable Then
        tmpStr = ir.ReadKey(Constants.SECTION_FORMAT, Constants.KEY_PIVOT_WORD_WRAP_COLS)
        If Len(tmpStr) <> 0 Then
            tmpList = Split(tmpStr, ",")
            Set mPivotWordWrapCols = New Collection
            For Each v In tmpList
                mPivotWordWrapCols.Add CInt(Trim(CStr(v)))
            Next v
        End If
    End If
    mStartHeaderCol = ir.ReadLongKey(Constants.SECTION_FORMAT, Constants.KEY_START_HEADER_COL)
    Logger.LogDebug "ReportMetaData.Init", "Start header column: " & CStr(mStartHeaderCol)
    mStartHeaderRow = ir.ReadLongKey(Constants.SECTION_FORMAT, Constants.KEY_START_HEADER_ROW)
    mStartCategoryRow = ir.ReadLongKey(Constants.SECTION_FORMAT, Constants.KEY_START_CATEGORY_ROW)
    Logger.LogDebug "ReportMetaData.Init", "Start header row: " & CStr(mStartHeaderRow)
    mStartCol = ir.ReadLongKey(Constants.SECTION_FORMAT, Constants.KEY_START_COL)
    Logger.LogDebug "ReportMetaData.Init", "Start column: " & CStr(mStartCol)
    mStartRow = ir.ReadLongKey(Constants.SECTION_FORMAT, Constants.KEY_START_ROW)
    Logger.LogDebug "ReportMetaData.Init", "Start row: " & CStr(mStartRow)
    
    
    
    Logger.LogDebug "ReportMetaData.Init", "Read query path: " & mQueryFilePath
    mValid = False
    mQuery = FileHelper.ReadFileFullPath(mQueryFilePath)
    
    If Len(mQuery) <> 0 And StringHelper.IsContain(mQuery, Constants.SPLIT_LEVEL_S, True) _
            And StringHelper.IsContain(mQuery, Constants.SPLIT_LEVEL_0, True) Then
        Set mReportSheets = New Scripting.Dictionary
        
        mValid = True
        tmpRawSheet = Split(mQuery, Constants.SPLIT_LEVEL_S)
        Dim tmp1 As String
        Dim tmp2 As String
        For i = LBound(tmpRawSheet) To UBound(tmpRawSheet)
            tmpStr = tmpRawSheet(i)
            If StringHelper.IsContain(tmpStr, Constants.SPLIT_LEVEL_0, True) Then
                tmpRawSheetSection = Split(tmpStr, Constants.SPLIT_LEVEL_0)
                tmp1 = StringHelper.TrimNewLine(tmpRawSheetSection(0))
                Logger.LogDebug "ReportMetaData.Init", "Found sheet " & tmp1
                tmp2 = tmpRawSheetSection(1)
                tmpRawSection = Split(tmp2, Constants.SPLIT_LEVEL_1)
                Set mReportSect = New Collection
                For j = LBound(tmpRawSection) To UBound(tmpRawSection)
                    Logger.LogDebug "ReportMetaData.Init", "Found section " & CStr(j + 1)
                    Set rpSection = New ReportSection
                    tmpStr = Trim(tmpRawSection(j))
                    rpSection.Init tmpStr, ss, mSkipCheckHeader
                    If StringHelper.IsEqual(rpSection.SectionType, Constants.RP_SECTION_TYPE_FIXED, True) _
                        Or StringHelper.IsEqual(rpSection.SectionType, Constants.RP_SECTION_TYPE_TMP_TABLE, True) Then
                        mCount = rpSection.count
                    End If
                    If Not rpSection.Valid Then
                        mValid = False
                    End If
                    Logger.LogDebug "ReportMetaData.Init", "Add section to sheet " & tmp1
                    mReportSect.Add rpSection
                Next j
                Logger.LogDebug "ReportMetaData.Init", "Add all sections to sheet " & tmp1
                mReportSheets.Add tmp1, mReportSect
            End If
        Next i
    End If
    mComplete = False
End Function

Public Function SetOutputPath(path As String)
    mOutputPath = path
    mComplete = True
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
    Dim ws As Excel.worksheet
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
    Dim ws As Excel.worksheet
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
    'FileHelper.DeleteFile mOutputPath
    mComplete = False
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

Public Property Get startRow() As Long
    startRow = mStartRow
End Property

Public Property Get StartCol() As Long
    StartCol = mStartCol
End Property

Public Property Get worksheet() As String
    worksheet = mWorksheet
End Property

Public Property Get FillHeader() As Boolean
    FillHeader = mFillHeader
End Property

Public Property Get FillCategory() As Boolean
    FillCategory = mFillCategory
End Property

Public Property Get StartHeaderRow() As Long
    StartHeaderRow = mStartHeaderRow
End Property

Public Property Get StartCategoryRow() As Long
    StartCategoryRow = mStartCategoryRow
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

Public Property Get count() As Long
    count = mCount
End Property

Public Function SetComplete(done As Boolean)
    mComplete = done
    mLastModified = FileHelper.FileLastModified(OutputPath)
End Function

Public Function IsChange() As Boolean
    IsChange = False
    If FileHelper.IsExistFile(OutputPath) Then
        If Not StringHelper.IsEqual(FileHelper.FileLastModified(OutputPath), mLastModified, True) Then
            IsChange = True
        End If
    End If
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

Public Property Get PivotTable() As Boolean
    PivotTable = mPivotTable
End Property

Public Property Get PivotTableName() As String
    PivotTableName = mPivotTableName
End Property

Public Property Get PivotTableWorksheet() As String
    PivotTableWorksheet = mPivotTableWorksheet
End Property

Public Property Get PivotWordWrapCols() As Collection
    If mPivotWordWrapCols Is Nothing Then
        Set mPivotWordWrapCols = New Collection
    End If
    Set PivotWordWrapCols = mPivotWordWrapCols
End Property

Public Property Get RpName() As String
    RpName = rawName
End Property

Public Property Get DateFormat() As String
    DateFormat = mDateFormat
End Property

Public Property Get DateColumes() As Collection
    Set DateColumes = mDateColumes
End Property

Public Property Get ReportSheets() As Scripting.Dictionary
    Set ReportSheets = mReportSheets
End Property