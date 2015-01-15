Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

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
    Dim i As Integer
    Dim mData As String
    Dim fields As String
    mNtid = iNtid
    Dim result As String
    Dim tmpNtid As String
    Dim check As String
    Dim tmpDict As Scripting.Dictionary, checkList As Collection
    If ss Is Nothing Then
        Set ss = Session.Settings()
    End If
    
   ' If ss.EnableTesting Then
    If True Then
        mAuth = True
    Else
        mAuth = False
        For i = 0 To ss.validatorMapping.count - 1
            fields = fields & ss.validatorMapping.Items(i) & ","
        Next i
        If StringHelper.EndsWith(fields, ",", True) Then
            fields = Left(fields, Len(fields) - 1)
        End If
        mData = "token=" & StringHelper.EncodeURL(ss.Token) _
                & "&fields=" & StringHelper.EncodeURL(fields) _
                & "&ntids=" & StringHelper.EncodeURL(mNtid)
        Logger.LogDebug "CurrentUser.Init", "Post valid ntid: " & mNtid
        result = Trim(HttpHelper.PostData(ss.ValidatorURL, mData))
        Logger.LogDebug "CurrentUser.Init", "Result: " & result
        
        If Len(result) > 0 Then
            If StringHelper.IsContain(result, "}", True) And StringHelper.IsContain(result, "{", True) Then
                Set checkList = JSONHelper.parse(result)
                For Each tmpDict In checkList
                    tmpNtid = tmpDict.Item("ntid")
                    check = tmpDict.Item("isvalid")
                    If StringHelper.IsEqual(tmpNtid, mNtid, True) And StringHelper.IsEqual(check, "true", True) Then
                        mAuth = True
                    End If
                Next
            End If
        Else
            mAuth = True
        End If
    End If
    ' skip temp
    If Ultilities.CheckTables(Constants.SYNC_TYPE_ROLE) Then
        mValid = False
        Dim dbm As New DbManager
        Dim frg As FunctionRegion
        Dim query As String
        Dim data As New Scripting.Dictionary
        Dim regionName As String
        Dim functionId As String
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
                regionName = dbm.GetFieldValue(dbm.RecordSet, "RegionName")
                roleName = dbm.GetFieldValue(dbm.RecordSet, "roleName")
                Logger.LogDebug "CurrentUser.Init", roleName & " | " & lastRoleName
                Logger.LogDebug "CurrentUser.Init", regionName & " | " & lastRegionName
                If (Not StringHelper.IsEqual(regionName, lastRegionName, True)) Then
                    Logger.LogDebug "CurrentUser.Init", "Init new function region " & regionName
                    Set frg = New FunctionRegion
                    frg.Init regionName, _
                        roleName, _
                        dbm.GetFieldValue(dbm.RecordSet, "permission")
                Else
                    Logger.LogDebug "CurrentUser.Init", "Add more role " & roleName
                    frg.AddRole roleName
                End If
                If (Not StringHelper.IsEqual(regionName, lastRegionName, True)) Then
                    Logger.LogDebug "CurrentUser.Init", "Add to list "
                    If mFuncRegion Is Nothing Then
                        Set mFuncRegion = frg
                    End If
                    mListFuncRg.Add frg
                End If
                lastRegionName = regionName
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

Public Function SelectRegion(rname As String)
    Dim frg As FunctionRegion
    If mValid And Not mListFuncRg Is Nothing Then
        For Each frg In mListFuncRg
            If StringHelper.IsEqual(frg.Region, rname, True) Then
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

Public Function LoadReportCache(func As String)
    Set mReportCache = New Dictionary
    Logger.LogDebug "CurrentUser.GetReportCache", "Try to load from cache info"
    Set mReportCache = FileHelper.ReadDictionary(FileHelper.tmpDir & "/" & mNtid & "_" & mFuncRegion.Region & "_" & func & "_report.cache")
End Function

Public Function SaveReportCache(func As String)
    FileHelper.SaveDictionary FileHelper.tmpDir & "/" & mNtid & "_" & mFuncRegion.Region & "_" & func & "_report.cache", mReportCache
End Function

Public Function GetReportCache(Name As String, func As String)
    LoadReportCache func
    If StringHelper.DictExistKey(mReportCache, Name) Then
        GetReportCache = StringHelper.DictGetValue(mReportCache, Name)
    Else
        GetReportCache = ""
    End If
End Function

Public Function AddReportCache(Name As String, path As String, func As String)
    LoadReportCache func
    If StringHelper.DictExistKey(mReportCache, Name) Then
        mReportCache.Remove Name
    End If
    mReportCache.Add Name, path
    SaveReportCache func
End Function

Public Function RecheckLocalChangeForReportCacheByTable(Name As String)
    Dim dbm As New DbManager
    Dim tmpName As String
    dbm.Init
    dbm.OpenRecordSet "select TableName from ChangeLog where CacheStatus=0 and TableName='" & StringHelper.EscapeQueryString(Name) & "' group by TableName"
    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
        RemoveReportCacheByTable Name
        dbm.ExecuteQuery "update ChangeLog set CacheStatus=-1 where CacheStatus=0 and TableName='" & StringHelper.EscapeQueryString(Name) & "'"
    End If
    dbm.Recycle
End Function

Public Function RecheckLocalChangeForReportCache()
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

Public Function RemoveReportCacheTableFunc(Name As String, tmpFid As String)
    If StringHelper.IsEqual(Name, "user_data", True) _
            Or StringHelper.IsEqual(Name, "BpRoleStandard", True) _
            Or StringHelper.IsEqual(Name, "user_data_mapping_role", True) Then
        RemoveReportCache Constants.RP_AD_HOC_REPORTING, tmpFid
    End If
    If StringHelper.IsEqual(Name, "audit_logs", True) Then
        RemoveReportCache Constants.RP_AUDIT_LOG, tmpFid
    End If
    If StringHelper.IsEqual(Name, "user_data", True) _
            Or StringHelper.IsEqual(Name, "CourseMappingBpRoleStandard", True) _
            Or StringHelper.IsEqual(Name, "course", True) _
            Or StringHelper.IsEqual(Name, "BpRoleStandard", True) _
            Or StringHelper.IsEqual(Name, "Functions", True) _
            Or StringHelper.IsEqual(Name, "user_data_mapping_role", True) Then
        RemoveReportCache Constants.RP_COURSE_ANALYTICS, tmpFid
    End If
    If StringHelper.IsEqual(Name, "user_data_mapping_role", True) _
            Or StringHelper.IsEqual(Name, "user_data", True) _
            Or StringHelper.IsEqual(Name, "BpRoleStandard", True) _
            Or StringHelper.IsEqual(Name, "specialism", True) _
            Or StringHelper.IsEqual(Name, "SpecialismMappingActivity", True) _
            Or StringHelper.IsEqual(Name, "activity", True) Then
        RemoveReportCache Constants.RP_END_USER_TO_BB_ACTIVITY, tmpFid
    End If
    If StringHelper.IsEqual(Name, "user_data_mapping_role", True) _
            Or StringHelper.IsEqual(Name, "user_data", True) _
            Or StringHelper.IsEqual(Name, "BpRoleStandard", True) Then
        RemoveReportCache Constants.RP_END_USER_TO_BB_JOB_ROLE, tmpFid
    End If
    If StringHelper.IsEqual(Name, "user_data_mapping_qualification", True) _
            Or StringHelper.IsEqual(Name, "user_data", True) _
            Or StringHelper.IsEqual(Name, "Qualifications", True) _
            Or StringHelper.IsEqual(Name, "user_data_mapping_role", True) Then
        RemoveReportCache Constants.RP_END_USER_TO_BB_QUALIFICATION, tmpFid
    End If
    If StringHelper.IsEqual(Name, "user_data_mapping_role", True) _
            Or StringHelper.IsEqual(Name, "user_data", True) _
            Or StringHelper.IsEqual(Name, "BpRoleStandard", True) _
            Or StringHelper.IsEqual(Name, "CourseMappingBpRoleStandard", True) _
            Or StringHelper.IsEqual(Name, "Course", True) _
            Or StringHelper.IsEqual(Name, "Functions", True) Then
        RemoveReportCache Constants.RP_END_USER_TO_COURSE, tmpFid
    End If
    If StringHelper.IsEqual(Name, "user_data", True) _
            Or StringHelper.IsEqual(Name, "user_data_mapping_role", True) _
            Or StringHelper.IsEqual(Name, "BpRoleStandard", True) _
            Or StringHelper.IsEqual(Name, "Dofa", True) Then
        RemoveReportCache Constants.RP_END_USER_TO_DOFA, tmpFid
    End If
    If StringHelper.IsEqual(Name, "user_data", True) _
            Or StringHelper.IsEqual(Name, "user_data_mapping_role", True) _
            Or StringHelper.IsEqual(Name, "BpRoleStandard", True) _
            Or StringHelper.IsEqual(Name, "Dofa", True) _
            Or StringHelper.IsEqual(Name, "specialism", True) _
            Or StringHelper.IsEqual(Name, "SpecialismMappingActivity", True) _
            Or StringHelper.IsEqual(Name, "activity", True) Then
        RemoveReportCache Constants.RP_END_USER_TO_EVERYTHING, tmpFid
    End If
    If StringHelper.IsEqual(Name, "user_data_mapping_role", True) _
            Or StringHelper.IsEqual(Name, "user_data", True) _
            Or StringHelper.IsEqual(Name, "BpRoleStandard", True) Then
        RemoveReportCache Constants.RP_END_USER_TO_SYSTEM_ROLE, tmpFid
    End If
    If StringHelper.IsEqual(Name, "user_change_log", True) Then
        RemoveReportCache Constants.RP_USER_DATA_CHANGE_LOG, tmpFid
    End If
End Function

Public Function RemoveReportCacheByTable(Name As String)
    On Error GoTo OnError
    Dim dbm As New DbManager
    Dim tmpFid As String
    Dim tmpName As String
    dbm.Init
    dbm.OpenRecordSet "select * from Functions where deleted=0"
    RemoveReportCacheTableFunc Name, Constants.TEXT_DEFAULT_ALL_FUNCTION
    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
        dbm.RecordSet.MoveFirst
        Do While Not dbm.RecordSet.EOF
            tmpFid = dbm.GetFieldValue(dbm.RecordSet, "id")
            tmpName = dbm.GetFieldValue(dbm.RecordSet, "nameFunction")
            RemoveReportCacheTableFunc Name, tmpFid
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

Public Function RemoveReportCache(Name As String, func As String)
    LoadReportCache func
    If StringHelper.DictExistKey(mReportCache, Name) Then
        Dim path As String
        path = StringHelper.DictGetValue(mReportCache, Name)
        FileHelper.DeleteFile path
        mReportCache.Remove Name
    End If
    SaveReportCache func
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