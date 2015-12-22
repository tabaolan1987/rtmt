Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mSelectedProcess As String
Private mNtid As String
Private mFullName As String
Private mFuncRegion As FunctionRegion
Private mListFuncRg As Collection
Private mAuth As Boolean
Private mValid As Boolean
Private mReportCache As Dictionary
Private mSelectedReportFunc As String

Public Property Get IsRole(role As String) As Boolean
    Dim v As Variant
    IsRole = False
    If Not mFuncRegion.role Is Nothing Then
        For Each v In mFuncRegion.role
            If StringHelper.IsEqual(CStr(v), role, True) Then
                Logger.LogDebug "CurrentUser.IsRole", "Found role: " & CStr(v)
                IsRole = True
                Exit For
            End If
        Next v
    End If
End Property

Public Property Get IsPermission(p As String) As Boolean
    Dim v As Variant
    IsPermission = False
    Logger.LogDebug "CurrentUser.IsPermission", "Check permission in Region: " + mFuncRegion.Region
    If Not mFuncRegion.permission Is Nothing Then
        For Each v In mFuncRegion.permission
            If StringHelper.IsEqual(CStr(v), p, True) Then
                Logger.LogDebug "CurrentUser.IsPermission", "Found permission: " & CStr(v)
                IsPermission = True
                Exit For
            End If
        Next v
    End If
    
End Property

Public Function Init(iNtid As String, _
                        Optional ss As SystemSetting)
    'revised code to use alternative BP authentification dominic gardham 01 July 2014
    mSelectedProcess = "false"
    Dim i As Integer
    Dim mData As String
    Dim fields As String
    mNtid = iNtid
    Dim objConnection As ADODB.Connection
    Dim objRecordSet As ADODB.RecordSet
    Dim objCommand As ADODB.Command

    Dim objUser, intUserAccountControl As Variant
    Dim result As String
    Dim tmpNtid, strUserID, strCreated, strSponsor, strUserGPID, strExchLegDN, strUserDN As String
    Dim check, strgivenID, strStaffFlag, intBPint01, strObjType  As String
    Dim tmpDict As Scripting.Dictionary, checkList As Collection
    Const ADS_UF_DONT_EXPIRE_PASSWD = &H10000
    Const ADS_UF_PASSWD_REQ = &H20
    Const ADS_OCS_REQ = &H2
    Const ADS_PROPERTY_APPEND = 3
    Const ADS_SCOPE_SUBTREE = 2
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
    objConnection.Provider = "ADsDSOObject"
    objConnection.Open "Active Directory Provider"
    Set objCommand.ActiveConnection = objConnection
    
    objCommand.Properties("Page Size") = 1000
    objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
    
    If ss Is Nothing Then
        Set ss = Session.Settings()
    End If

    If ss.EnableTesting Or StringHelper.IsEqual(ss.Env, "dev", True) Then
        mAuth = True
    Else
        mAuth = False
        On Error Resume Next
        strgivenID = VBA.Environ("USERNAME")
        'strUserID = Replace(strgivenID, src_chr, tar_chr)
        'strUserID = Trim(strUserID)
        objCommand.CommandText = _
            "SELECT DistinguishedName, LegacyExchangeDN, createTimeStamp, extensionAttribute9, BP-Integer-01, BP-text32-01 FROM 'LDAP://ou=client,dc=bp1,dc=ad,dc=bp,dc=com' WHERE objectCategory='user' " & _
            "AND SamaccountName='" & strgivenID & "'"
        Err.Clear
        Set objRecordSet = objCommand.Execute
        objRecordSet.MoveFirst
        Do Until objRecordSet.EOF
                strCreated = ""
                strSponsor = ""
                strUserGPID = ""
                strExchLegDN = ""
                strUserDN = objRecordSet.fields("DistinguishedName").value
                strCreated = objRecordSet.fields("CreateTimeStamp").value
                strStaffFlag = objRecordSet.fields("extensionAttribute9").value
                intBPint01 = objRecordSet.fields("BP-Integer-01").value
                strUserGPID = objRecordSet.fields("BP-text32-01").value
                strExchLegDN = objRecordSet.fields("LegacyExchangeDN").value
            objRecordSet.MoveNext
        Loop

        If Err.Number = 0 Then
            strObjType = "NO"
            Err.Clear
            Set objUser = GetObject("LDAP://" & strUserDN)
            intUserAccountControl = objUser.Get("userAccountControl")
            
            ' Do this section if Account has Password DOES NOT Expire ticked
            If objUser.AccountDisabled = False Then
                mAuth = True
            Else
                mAuth = False
            End If
        End If
    End If
    ' skip temp
    If Ultilities.CheckTables(Constants.SYNC_TYPE_ROLE) Then
        mValid = False
        Dim dbm As New DbManager
        Dim frg As FunctionRegion
        Dim query As String
        Dim data As New Scripting.Dictionary
        Dim RegionName As String
        Dim FunctionID As String
        Dim roleName As String
        Dim lastRegionName As String
        Dim lastFunctionId As String
        Dim lastRoleName As String
        data.Add Constants.Q_KEY_VALUE, mNtid
        dbm.Init
        query = FileHelper.ReadQuery(Constants.TABLE_USER_PRIVILEGES, Constants.Q_SELECT)
        query = StringHelper.GenerateQuery(query, data)
        Logger.LogDebug "CurrentUser.Init", query
        dbm.OpenRecordSet query
        If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
            dbm.RecordSet.MoveFirst
            mValid = True
            Logger.LogDebug "CurrentUser.Init", "Valid: " & mValid
            Set mListFuncRg = New Collection
            Do Until dbm.RecordSet.EOF = True
                RegionName = dbm.GetFieldValue(dbm.RecordSet, "RegionName")
                roleName = dbm.GetFieldValue(dbm.RecordSet, "roleName")
                Logger.LogDebug "CurrentUser.Init", roleName & " | " & lastRoleName
                Logger.LogDebug "CurrentUser.Init", RegionName & " | " & lastRegionName
                If (Not StringHelper.IsEqual(RegionName, lastRegionName, True)) Then
                    Logger.LogDebug "CurrentUser.Init", "Init new function region " & RegionName
                    Set frg = New FunctionRegion
                    frg.Init RegionName, _
                        roleName, _
                        dbm.GetFieldValue(dbm.RecordSet, "permission")
                Else
                    Logger.LogDebug "CurrentUser.Init", "Add more role " & roleName
                    frg.AddRole roleName
                End If
                If (Not StringHelper.IsEqual(RegionName, lastRegionName, True)) Then
                    Logger.LogDebug "CurrentUser.Init", "Add to list "
                    If mFuncRegion Is Nothing Then
                        Set mFuncRegion = frg
                    End If
                    mListFuncRg.Add frg
                End If
                lastRegionName = RegionName
                lastRoleName = roleName
                dbm.RecordSet.MoveNext
            Loop
            Logger.LogDebug "CurrentUser.Init", "complete..."
        Else
            Logger.LogDebug "CurrentUser.Init", "Not valid"
        End If
    End If
End Function

Public Function ListRegions() As Collection
    Dim frg As FunctionRegion
    Dim list As New Collection
    Dim v As Variant
    Dim check As Boolean
    Dim tmpName As String
    If mValid And Not mListFuncRg Is Nothing Then
        For Each frg In mListFuncRg
            tmpName = frg.Region
            check = False
            For Each v In list
                If StringHelper.IsEqual(CStr(v), tmpName, True) Then
                    check = True
                    Exit For
                End If
            Next v
            If Not check Then
                list.Add tmpName
            End If
        Next frg
    End If
    Set ListRegions = list
End Function

Public Function ListFunctions(Region As String) As Collection
    Dim frg As FunctionRegion
    Dim list As New Collection
    Dim v As Variant
    Dim tmpName As String
    If mValid And Not mListFuncRg Is Nothing Then
        For Each frg In mListFuncRg
            tmpName = frg.Region
            If StringHelper.IsEqual(tmpName, Region, True) Then
                list.Add frg.Name
            End If
        Next frg
    End If
    Set ListFunctions = list
End Function

Public Function SelectRegion(rName As String)
    Dim frg As FunctionRegion
    If mValid And Not mListFuncRg Is Nothing Then
        For Each frg In mListFuncRg
            If StringHelper.IsEqual(frg.Region, rName, True) Then
                Set mFuncRegion = frg
                Exit For
            End If
        Next frg
    End If
End Function

Public Function SelectFunc(fname As String)
    Dim frg As FunctionRegion
    If mValid And Not mListFuncRg Is Nothing Then
        For Each frg In mListFuncRg
            If StringHelper.IsEqual(frg.value, fname, True) Then
                Set mFuncRegion = frg
                Exit For
            End If
        Next frg
    End If
End Function

Public Function LoadReportCache(mRegion As String, func As String)
    Set mReportCache = New Dictionary
    Logger.LogDebug "CurrentUser.GetReportCache", "Try to load from cache info"
    Set mReportCache = FileHelper.ReadDictionary(FileHelper.tmpDir & "/" & mNtid & "_" & mRegion & "_" & func & "_report.cache")
End Function

Public Function SaveReportCache(mRegion, func As String)
    FileHelper.SaveDictionary FileHelper.tmpDir & "/" & mNtid & "_" & mRegion & "_" & func & "_report.cache", mReportCache
End Function

Public Function GetReportCache(rName As String, func As String, mRegion As String)
    LoadReportCache mRegion, func
    If StringHelper.DictExistKey(mReportCache, rName) Then
        GetReportCache = StringHelper.DictGetValue(mReportCache, rName)
    Else
        GetReportCache = ""
    End If
End Function

Public Function AddReportCache(rName As String, path As String, func As String, mRegion As String)
    LoadReportCache mRegion, func
    If StringHelper.DictExistKey(mReportCache, rName) Then
        mReportCache.Remove rName
    End If
    mReportCache.Add rName, path
    SaveReportCache mRegion, func
End Function

Public Function RecheckLocalChangeForReportCacheByTable(rName As String)
    Dim dbm As New DbManager
    Dim tmpName As String
    dbm.Init
    dbm.OpenRecordSet "select TableName from ChangeLog where CacheStatus=0 and TableName='" & StringHelper.EscapeQueryString(rName) & "' group by TableName"
    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
        RemoveReportCacheByTable rName
        dbm.ExecuteQuery "update ChangeLog set CacheStatus=-1 where CacheStatus=0 and TableName='" & StringHelper.EscapeQueryString(rName) & "'"
    End If
    dbm.Recycle
End Function

Public Function RecheckLocalChangeForReportCache(test As Boolean)
    Dim dbm As New DbManager
    Dim tmpName As String
    dbm.Init
    dbm.OpenRecordSet "select TableName from ChangeLog where CacheStatus=0 group by TableName"
    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
        dbm.RecordSet.MoveFirst
        Do While Not dbm.RecordSet.EOF
            tmpName = dbm.GetFieldValue(dbm.RecordSet, "TableName")
            RemoveReportCacheByTable tmpName
            dbm.RecordSet.MoveNext
        Loop
        dbm.ExecuteQuery "update ChangeLog set CacheStatus=-1 where CacheStatus=0"
    End If
    dbm.Recycle
End Function

Public Function RemoveReportCacheTableFunc(rName As String, tmpFid As String, mRegion As String)
    If StringHelper.IsEqual(rName, "user_data", True) _
            Or StringHelper.IsEqual(rName, "BpRoleStandard", True) _
            Or StringHelper.IsEqual(rName, "user_data_mapping_role", True) Then
        RemoveReportCache Constants.RP_AD_HOC_REPORTING, tmpFid, mRegion
    End If
    If StringHelper.IsEqual(rName, "audit_logs", True) Then
        RemoveReportCache Constants.RP_AUDIT_LOG, tmpFid, mRegion
    End If
    If StringHelper.IsEqual(rName, "user_data", True) _
            Or StringHelper.IsEqual(rName, "CourseMappingBpRoleStandard", True) _
            Or StringHelper.IsEqual(rName, "course", True) _
            Or StringHelper.IsEqual(rName, "BpRoleStandard", True) _
            Or StringHelper.IsEqual(rName, "Functions", True) _
            Or StringHelper.IsEqual(rName, "user_data_mapping_role", True) Then
        RemoveReportCache Constants.RP_COURSE_ANALYTICS, tmpFid, mRegion
    End If
    If StringHelper.IsEqual(rName, "user_data_mapping_role", True) _
            Or StringHelper.IsEqual(rName, "user_data", True) _
            Or StringHelper.IsEqual(rName, "BpRoleStandard", True) _
            Or StringHelper.IsEqual(rName, "specialism", True) _
            Or StringHelper.IsEqual(rName, "SpecialismMappingActivity", True) _
            Or StringHelper.IsEqual(rName, "activity", True) Then
        RemoveReportCache Constants.RP_END_USER_TO_BB_ACTIVITY, tmpFid, mRegion
    End If
    If StringHelper.IsEqual(rName, "user_data_mapping_role", True) _
            Or StringHelper.IsEqual(rName, "user_data", True) _
            Or StringHelper.IsEqual(rName, "BpRoleStandard", True) Then
        RemoveReportCache Constants.RP_END_USER_TO_BB_JOB_ROLE, tmpFid, mRegion
    End If
    If StringHelper.IsEqual(rName, "user_data_mapping_qualification", True) _
            Or StringHelper.IsEqual(rName, "user_data", True) _
            Or StringHelper.IsEqual(rName, "Qualifications", True) _
            Or StringHelper.IsEqual(rName, "user_data_mapping_role", True) Then
        RemoveReportCache Constants.RP_END_USER_TO_BB_QUALIFICATION, tmpFid, mRegion
    End If
    If StringHelper.IsEqual(rName, "user_data_mapping_role", True) _
            Or StringHelper.IsEqual(rName, "user_data", True) _
            Or StringHelper.IsEqual(rName, "BpRoleStandard", True) _
            Or StringHelper.IsEqual(rName, "CourseMappingBpRoleStandard", True) _
            Or StringHelper.IsEqual(rName, "Course", True) _
            Or StringHelper.IsEqual(rName, "Functions", True) Then
        RemoveReportCache Constants.RP_END_USER_TO_COURSE, tmpFid, mRegion
    End If
    If StringHelper.IsEqual(rName, "user_data", True) _
            Or StringHelper.IsEqual(rName, "user_data_mapping_role", True) _
            Or StringHelper.IsEqual(rName, "BpRoleStandard", True) _
            Or StringHelper.IsEqual(rName, "Dofa", True) Then
        RemoveReportCache Constants.RP_END_USER_TO_DOFA, tmpFid, mRegion
    End If
    If StringHelper.IsEqual(rName, "user_data", True) _
            Or StringHelper.IsEqual(rName, "user_data_mapping_role", True) _
            Or StringHelper.IsEqual(rName, "BpRoleStandard", True) _
            Or StringHelper.IsEqual(rName, "Dofa", True) _
            Or StringHelper.IsEqual(rName, "specialism", True) _
            Or StringHelper.IsEqual(rName, "SpecialismMappingActivity", True) _
            Or StringHelper.IsEqual(rName, "activity", True) Then
        RemoveReportCache Constants.RP_END_USER_TO_EVERYTHING, tmpFid, mRegion
    End If
    If StringHelper.IsEqual(rName, "user_data_mapping_role", True) _
            Or StringHelper.IsEqual(rName, "user_data", True) _
            Or StringHelper.IsEqual(rName, "BpRoleStandard", True) Then
        RemoveReportCache Constants.RP_END_USER_TO_SYSTEM_ROLE, tmpFid, mRegion
    End If
    If StringHelper.IsEqual(rName, "user_data_mapping_role", True) _
            Or StringHelper.IsEqual(rName, "user_data", True) _
            Or StringHelper.IsEqual(rName, "BpRoleStandard", True) Then
        RemoveReportCache Constants.RP_ROLE_MAPPING_OUTPUT_OF_TOOL_FOR_SECURITY, tmpFid, mRegion
    End If
    If StringHelper.IsEqual(rName, "user_change_log", True) Then
        RemoveReportCache Constants.RP_USER_DATA_CHANGE_LOG, tmpFid, mRegion
    End If
End Function

Public Function RemoveReportCacheByTable(rName As String)
    On Error GoTo OnError
    Dim dbm As New DbManager
    Dim tmpFid As String
    Dim tmpName As String
    Dim v As Variant
    dbm.Init
    dbm.OpenRecordSet "select * from Functions where deleted=0"
    For Each v In ListRegions
        RemoveReportCacheTableFunc rName, Constants.TEXT_DEFAULT_ALL_FUNCTION, CStr(v)
    Next
    
    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
        dbm.RecordSet.MoveFirst
        Do While Not dbm.RecordSet.EOF
            tmpFid = dbm.GetFieldValue(dbm.RecordSet, "id")
            tmpName = dbm.GetFieldValue(dbm.RecordSet, "nameFunction")
            For Each v In ListRegions
                RemoveReportCacheTableFunc rName, tmpFid, CStr(v)
            Next
            dbm.RecordSet.MoveNext
        Loop
    End If
    dbm.Recycle
OnExit:
    Exit Function
OnError:
    'ShowError "An error occurred while processing"
    Logger.LogError "CurrentUser.RemoveReportCacheByTable", "An error occurred while processing", Err
    Resume OnExit
End Function

Public Function RemoveReportCache(rName As String, func As String, mRegion As String)
    LoadReportCache mRegion, func
    If StringHelper.DictExistKey(mReportCache, rName) Then
        Dim path As String
        path = StringHelper.DictGetValue(mReportCache, rName)
        FileHelper.DeleteFile path
        mReportCache.Remove rName
    End If
    SaveReportCache mRegion, func
End Function

Public Property Get Valid() As Boolean
    Valid = mValid
End Property

Public Property Get Auth() As Boolean
    Auth = mAuth
End Property

Public Property Get ntid() As String
    ntid = mNtid
End Property

Public Property Get FuncRegion() As FunctionRegion
    Set FuncRegion = mFuncRegion
End Property

Public Property Get ListFuncRg() As Collection
    Set ListFuncRg = mListFuncRg
End Property

Public Function SetNtid(ntid As String)
    mNtid = ntid
End Function

Public Function SetSelectedProcess(selected As String)
    mSelectedProcess = selected
End Function
Public Property Get selectedProcess() As String
    selectedProcess = mSelectedProcess
End Property