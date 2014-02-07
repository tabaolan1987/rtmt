Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @author Hai Lu
Option Compare Database
Private rawName As String
Private mName As String
Private mWorkSheet As String

Private mStartRowTop As Long
Private mStartRowLeft As Long
Private mStartColTop As Long
Private mStartColLeft As Long

Private mQueryTop As String
Private mQueryLeft As String

Private mQueryCheck As String

Private mQueryUpdate As String
Private mQueryInsert As String

Private mValid As Boolean
Private mConfigFilePath As String
Private mQueryFilePath As String
Private mTemplateFilePath As String

Private mLastModified As String

Private mOutputPath As String

Public Function Recyle()
    FileHelper.Delete mQueryFilePath
    FileHelper.Delete mTemplateFilePath
    FileHelper.Delete mConfigFilePath
    FileHelper.Delete mTemplateFilePath & Constants.FILE_EXTENSION_TEMPLATE
End Function

Public Function Init(mappingName As String, Optional ss As SystemSetting)
    rawName = mappingName
    Logger.LogDebug "MappingMetaData.Init", "Start init mapping meta name: " & rawName
    Dim tmpRawSection() As String, tmpStr As String, i As Integer
    Dim mQuery As String
    Dim rpSection As ReportSection
    Dim ir As New IniReader

    mQueryFilePath = FileHelper.DuplicateAsTemporary(FileHelper.CurrentDbPath & Constants.MAPPING_ROOT_FOLDER & rawName & Constants.FILE_EXTENSION_QUERY)
    mTemplateFilePath = FileHelper.DuplicateAsTemporary(FileHelper.CurrentDbPath & Constants.MAPPING_ROOT_FOLDER & rawName & Constants.FILE_EXTENSION_TEMPLATE)
    mConfigFilePath = FileHelper.DuplicateAsTemporary(FileHelper.CurrentDbPath & Constants.MAPPING_ROOT_FOLDER & rawName & Constants.FILE_EXTENSION_CONFIG)
    Logger.LogDebug "MappingMetaData.Init", "Read configuration path: " & mConfigFilePath
    ir.Init mConfigFilePath
    'RefreshLastModified
    
    mName = ir.ReadKey(Constants.SECTION_GENERAL, Constants.KEY_NAME)
    Logger.LogDebug "MappingMetaData.Init", "Mapping name: " & mName
    
    mWorkSheet = ir.ReadKey(Constants.SECTION_GENERAL, Constants.KEY_WORK_SHEET)
    Logger.LogDebug "MappingMetaData.Init", "Work sheet: " & mWorkSheet
    
    mStartColTop = ir.ReadLongKey(Constants.SECTION_TOP, Constants.KEY_START_COL)
    Logger.LogDebug "MappingMetaData.Init", "Start column top: " & CStr(mStartColTop)
    mStartColLeft = ir.ReadLongKey(Constants.SECTION_LEFT, Constants.KEY_START_COL)
    Logger.LogDebug "MappingMetaData.Init", "Start column left: " & CStr(mStartColLeft)
    mStartRowTop = ir.ReadLongKey(Constants.SECTION_TOP, Constants.KEY_START_ROW)
    Logger.LogDebug "MappingMetaData.Init", "Start row top: " & CStr(mStartRowTop)
    mStartRowLeft = ir.ReadLongKey(Constants.SECTION_LEFT, Constants.KEY_START_ROW)
    Logger.LogDebug "MappingMetaData.Init", "Start row left: " & CStr(mStartRowLeft)
    
    Logger.LogDebug "MappingMetaData.Init", "Start header column: " & CStr(mStartColTop)
    
    Logger.LogDebug "MappingMetaData.Init", "Read query path: " & mQueryFilePath
    mValid = False
    mQuery = FileHelper.ReadFileFullPath(mQueryFilePath)
    
    If Len(mQuery) <> 0 Then
        mValid = False
        tmpRawSection = Split(mQuery, Constants.SPLIT_LEVEL_1)
        If UBound(tmpRawSection) = 4 Then
            mValid = True
            mQueryLeft = Trim(tmpRawSection(0))
            mQueryTop = Trim(tmpRawSection(1))
            mQueryCheck = Trim(tmpRawSection(2))
            mQueryUpdate = Trim(tmpRawSection(3))
            mQueryInsert = Trim(tmpRawSection(4))
            Logger.LogDebug "MappingMetaData.Init", "mQueryLeft: " & mQueryLeft
            Logger.LogDebug "MappingMetaData.Init", "mQueryTop: " & mQueryTop
            Logger.LogDebug "MappingMetaData.Init", "mQueryCheck: " & mQueryCheck
            Logger.LogDebug "MappingMetaData.Init", "mQueryUpdate: " & mQueryUpdate
            Logger.LogDebug "MappingMetaData.Init", "mQueryInsert: " & mQueryInsert
        End If
    End If
End Function

Public Property Get WorkSheet() As String
    WorkSheet = mWorkSheet
End Property

Public Property Get Name() As String
    Name = mName
End Property

Public Property Get Valid() As Boolean
    Valid = mValid
End Property

Public Property Get StartRowTop() As Long
    StartRowTop = mStartRowTop
End Property

Public Property Get StartRowLeft() As Long
    StartRowLeft = mStartRowLeft
End Property

Public Property Get StartColLeft() As Long
    StartColLeft = mStartColLeft
End Property

Public Property Get StartColTop() As Long
    StartColTop = mStartColTop
End Property

Public Property Get TemplateFilePath() As String
    TemplateFilePath = mTemplateFilePath
End Property

Public Property Get LastModified() As String
    LastModified = mLastModified
End Property

Public Function CurrentModifedDate() As String
    CurrentModifedDate = FileHelper.FileLastModified(mTemplateFilePath & Constants.FILE_EXTENSION_TEMPLATE)
End Function

Public Function RefreshLastModified()
    mLastModified = FileHelper.FileLastModified(mTemplateFilePath & Constants.FILE_EXTENSION_TEMPLATE)
End Function

Public Function query(qType As Integer, Optional data As Scripting.Dictionary) As String
    Dim mQuery As String
    Dim v As Variant
    mQuery = ""
    Select Case qType
        Case Constants.Q_INSERT:
            mQuery = mQueryInsert
        Case Constants.Q_UPDATE:
            mQuery = mQueryUpdate
        Case Constants.Q_CHECK:
            mQuery = mQueryCheck
        Case Constants.Q_TOP:
            mQuery = mQueryTop
        Case Constants.Q_LEFT:
            mQuery = mQueryLeft
    End Select
    If Not data Is Nothing Then
        mQuery = StringHelper.GenerateQuery(mQuery, data)
    End If
    'Logger.LogDebug "MappingMetaData.Query", mQuery
    query = mQuery
End Function