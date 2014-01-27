Option Explicit

Private mCurrentUser As currentUser
Private ss As SystemSetting
Private mFlagMapping As Boolean
Private mFlagReports As Boolean
Private mMappingMDCol As Scripting.Dictionary
Private mReportMDCol As Scripting.Dictionary

Public Function Init()
    Dim mmd As MappingMetadata
    Dim rmd As ReportMetaData
    mFlagMapping = False
    mFlagReports = False
    currentUser
    Set mMappingMDCol = New Scripting.Dictionary
    Set mmd = New MappingMetadata
    mmd.Init Constants.MAPPING_ACTIVITIES_SPECIALISM
    mMappingMDCol.Add Constants.MAPPING_ACTIVITIES_SPECIALISM, mmd
    
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
End Function

Public Function currentUser() As currentUser
    If mCurrentUser Is Nothing Then
        Set mCurrentUser = New currentUser
        If Settings().EnableTesting Then
            mCurrentUser.Init Settings().TestNtid, Settings()
        Else
            mCurrentUser.Init Ultilities.GetUserName, Settings()
        End If
    End If
   Set currentUser = mCurrentUser
End Function

Public Function Settings() As SystemSetting
    If ss Is Nothing Then
        Set ss = New SystemSetting
        ss.Init
    End If
    Set Settings = ss
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