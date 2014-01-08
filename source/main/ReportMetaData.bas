Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @author Hai Lu
' Report meta data object
Option Explicit
Private mQuery As String
Private mValid As Boolean
Private mConfigFilePath As String
Private mQueryFilePath As String
Private mTemplateFilePath As String

Private mReportSections As Collection

Public Function Init(name As String)
    Logger.LogDebug "ReportMetaData.Init", "Start init report meta name: " & name
    Dim tmpRawSection() As String, tmpStr As String, i As Integer
    Dim rpSection As ReportSection
    mQueryFilePath = FileHelper.CurrentDbPath & Constants.RP_ROOT_FOLDER & name & Constants.RP_QUERY_FILE_EXTENSION
    mTemplateFilePath = FileHelper.CurrentDbPath & Constants.RP_ROOT_FOLDER & name & Constants.RP_TEMPLATE_FILE_EXTENSION
    mConfigFilePath = FileHelper.CurrentDbPath & Constants.RP_ROOT_FOLDER & name & Constants.RP_CONFIG_FILE_EXTENSION
    
    Logger.LogDebug "ReportMetaData.Init", mQueryFilePath
    mValid = False
    mQuery = FileHelper.ReadFileFullPath(mQueryFilePath)
    
    If Len(mQuery) <> 0 Then
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

Public Property Get query() As String
    query = mQuery
End Property

Public Property Get Valid() As Boolean
    Valid = mValid
End Property