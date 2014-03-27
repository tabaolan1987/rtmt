' @author Hai Lu
' To control the current session data
Option Explicit

Private mCurrentUser As CurrentUser
Private ss As SystemSetting
Private mFlagMapping As Boolean
Private mFlagReports As Boolean
Private mMappingMDCol As Scripting.Dictionary
Private mReportMDCol As Scripting.Dictionary
Private mSelectedCSV As String
Private mAllReportsZip As String
Private mCurrentHelpContent As String
Private mMappingTypes As Scripting.Dictionary
Private mSelectedCurriculum As String

Public Function Recycle()
    Set ss = Nothing
    
End Function

Public Function SelectedCurriculum() As String
    SelectedCurriculum = mSelectedCurriculum
End Function

Public Function SetSelectedCurriculum(path As String)
    mSelectedCurriculum = path
End Function


Public Function SelectedCSV() As String
    SelectedCSV = mSelectedCSV
End Function

Public Function SetSelectedCSV(csvPath As String)
    mSelectedCSV = csvPath
End Function

Public Function RecycleUser()
    Set ss = Nothing
    Set mCurrentUser = Nothing
End Function

Public Function RecyleMapping()
    mFlagMapping = False
    Dim md As MappingMetaData
    Dim v As Variant
    If Not mMappingMDCol Is Nothing Then
        For Each v In mMappingMDCol.keys
            Set md = mMappingMDCol.Item(CStr(v))
            md.Recyle
        Next v
    End If
    Set mMappingMDCol = New Scripting.Dictionary
    
End Function

Public Function RecyleReports()
    mFlagReports = False
    Dim rmd As ReportMetaData
    Dim v As Variant
    ' Remove all cache report
    If Not mReportMDCol Is Nothing Then
        For Each v In mReportMDCol.keys
            Set rmd = mReportMDCol.Item(CStr(v))
            rmd.Recyle
        Next v
    End If
    Set mReportMDCol = Nothing
    FileHelper.DeleteFile mAllReportsZip
    mAllReportsZip = ""
End Function

Public Function RecyleReport(Name As String)
    mFlagReports = False
    Dim rmd As ReportMetaData
    Dim v As Variant
    ' Remove all cache report
    If Not mReportMDCol Is Nothing Then
        Set rmd = mReportMDCol.Item(Name)
        rmd.Recyle
        mReportMDCol.Remove (Name)
    End If
    FileHelper.DeleteFile mAllReportsZip
    mAllReportsZip = ""
End Function

Public Function RenewReports()
    Dim rmd As ReportMetaData
    Set mReportMDCol = New Scripting.Dictionary
    
End Function

Public Function Init()
    Set mCurrentUser = Nothing
    CurrentUser
    Recycle
End Function

Public Function MappingMetaData(mappingName As String) As MappingMetaData
    If mMappingMDCol Is Nothing Then
        RecyleMapping
    End If
    Dim md As MappingMetaData
    If Not mMappingMDCol.Exists(mappingName) Then
        Set md = New MappingMetaData
        md.Init mappingName, Settings
        mMappingMDCol.Add mappingName, md
    End If
    Set MappingMetaData = mMappingMDCol.Item(mappingName)
End Function

Public Function ReportMetaData(reportName As String) As ReportMetaData
    If mReportMDCol Is Nothing Then
        RenewReports
    End If
    Dim rmd As ReportMetaData
    If Not mReportMDCol.Exists(reportName) Then
        If StringHelper.IsEqual(reportName, Constants.RP_AUDIT_LOG, True) Then
            Dim dbm As New DbManager
            dbm.Init
            dbm.DeleteTable Constants.TABLE_AUDIT_LOG
            dbm.SyncTable Settings.ServerName & "," & Settings.Port, Settings.DatabaseName, Constants.TABLE_AUDIT_LOG, Constants.TABLE_AUDIT_LOG, Settings.Username, Settings.Password, False
            dbm.Recycle
        End If
        Set rmd = New ReportMetaData
        rmd.Init reportName
        mReportMDCol.Add reportName, rmd
    End If
    Set ReportMetaData = mReportMDCol.Item(reportName)
End Function

Public Function ReportMDCols() As Collection
     Dim list As New Collection
     list.Add Constants.RP_AUDIT_LOG
     list.Add Constants.RP_END_USER_TO_BB_JOB_ROLE
     list.Add Constants.RP_END_USER_TO_COURSE
     list.Add Constants.RP_ROLE_MAPPING_OUTPUT_OF_TOOL_FOR_SECURITY
     Set ReportMDCols = list
End Function

Public Function CurrentUser() As CurrentUser
    If mCurrentUser Is Nothing Then
        Set mCurrentUser = New CurrentUser
        If Settings().EnableTesting Then
            Logger.LogDebug "Session.CurrentUser", "Enable testing mode"
            mCurrentUser.Init Settings().TestNtid, Settings()
        Else
            Logger.LogDebug "Session.CurrentUser", "Disable testing mode"
            mCurrentUser.Init Ultilities.GetUserName, Settings()
        End If
    End If
   Set CurrentUser = mCurrentUser
End Function

Public Function Settings() As SystemSetting
    On Error GoTo OnError
    If ss Is Nothing Then
        Set ss = New SystemSetting
        ss.Init
    End If
    Set Settings = ss
OnExit:
    Exit Function
OnError:
    Set Settings = Nothing
    Resume OnExit
End Function

Public Function SetFlagMapping(change As Boolean)
    mFlagMapping = change
End Function

Public Function FlagMapping() As Boolean
    FlagMapping = mFlagMapping
End Function

Public Function SetFlagReports(change As Boolean)
    mFlagReports = change
End Function

Public Function FlagReports() As Boolean
    FlagReports = mFlagReports
End Function

Public Function SetAllReportsZip(zipPath As String)
    mAllReportsZip = zipPath
End Function

Public Function AllReportsZip() As String
    AllReportsZip = mAllReportsZip
End Function


Public Function CurrentHelpContent() As String
    If Len(mCurrentHelpContent) = 0 Then
        mCurrentHelpContent = Constants.HELP_UPLOAD_EUDL
    End If
    CurrentHelpContent = mCurrentHelpContent
End Function

Public Function SetCurrentHelpContent(help As String)
    mCurrentHelpContent = help
End Function

Public Function MappingTypes() As Scripting.Dictionary
    If mMappingTypes Is Nothing Then
        Set mMappingTypes = New Scripting.Dictionary
        Dim dbm As New DbManager
        dbm.Init
        dbm.OpenRecordSet "select * from mappingType where deleted=0"
        If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
            dbm.RecordSet.MoveFirst
            Do While Not dbm.RecordSet.EOF
                mMappingTypes.Add dbm.GetFieldValue(dbm.RecordSet, "id"), dbm.GetFieldValue(dbm.RecordSet, "mapp_name")
                dbm.RecordSet.MoveNext
            Loop
        End If
        dbm.Recycle
    End If
    Set MappingTypes = mMappingTypes
End Function

Public Function UpdateChangelog(tblName As String, tblId As String)
    Dim dbm As New DbManager
    dbm.Init
    dbm.OpenRecordSet "select * from [ChangeLog] where [TableName]='" & StringHelper.EscapeQueryString(tblName) _
                                    & "' and [TableId]='" & StringHelper.EscapeQueryString(tblId) & "'"
    If (dbm.RecordSet.BOF And dbm.RecordSet.EOF) Then
        dbm.ExecuteQuery "insert into [ChangeLog]([TableName], [TableId]) values('" & StringHelper.EscapeQueryString(tblName) _
                                    & "', '" & StringHelper.EscapeQueryString(tblId) & "')"
    End If
    dbm.Recycle
End Function