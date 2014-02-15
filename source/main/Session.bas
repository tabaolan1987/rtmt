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

Public Function Recycle()
    Set ss = Nothing
    
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
    Dim md As MappingMetadata
    Dim v As Variant
    If Not mMappingMDCol Is Nothing Then
        For Each v In mMappingMDCol.keys
            Set md = mMappingMDCol.Item(CStr(v))
            md.Recyle
        Next v
    End If
    Set mMappingMDCol = New Scripting.Dictionary
    Set md = New MappingMetadata
    md.Init Constants.MAPPING_ACTIVITIES_BB_JOB_ROLE, Settings
    mMappingMDCol.Add Constants.MAPPING_ACTIVITIES_BB_JOB_ROLE, md
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

Public Function RenewReports()
    Dim rmd As ReportMetaData
    Set mReportMDCol = New Scripting.Dictionary
    Set rmd = New ReportMetaData
    rmd.Init Constants.RP_END_USER_TO_BB_JOB_ROLE
    mReportMDCol.Add Constants.RP_END_USER_TO_BB_JOB_ROLE, rmd
    
    Set rmd = New ReportMetaData
    rmd.Init Constants.RP_END_USER_TO_COURSE
    mReportMDCol.Add Constants.RP_END_USER_TO_COURSE, rmd
    
    Set rmd = New ReportMetaData
    rmd.Init Constants.RP_ROLE_MAPPING_OUTPUT_OF_TOOL_FOR_SECURITY
    mReportMDCol.Add Constants.RP_ROLE_MAPPING_OUTPUT_OF_TOOL_FOR_SECURITY, rmd
    
    Set rmd = New ReportMetaData
    rmd.Init Constants.RP_AUDIT_LOG
    mReportMDCol.Add Constants.RP_AUDIT_LOG, rmd
End Function

Public Function Init()
    Set mCurrentUser = Nothing
    CurrentUser
    Recycle
End Function

Public Function MappingMDCol() As Scripting.Dictionary
    If mMappingMDCol Is Nothing Then
        RecyleMapping
    End If
    Set MappingMDCol = mMappingMDCol
End Function

Public Function ReportMDCol() As Scripting.Dictionary
    If mReportMDCol Is Nothing Then
        RenewReports
    End If
    Set ReportMDCol = mReportMDCol
End Function

Public Function CurrentUser() As CurrentUser
    If mCurrentUser Is Nothing Then
        Set mCurrentUser = New CurrentUser
        If Settings().EnableTesting Then
            mCurrentUser.Init Settings().TestNtid, Settings()
        Else
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