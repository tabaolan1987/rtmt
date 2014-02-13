Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Private ss As SystemSetting
Private dbm As DbManager
Private mIsConflict As Boolean
Private mIsDuplicate As Boolean
Private mIsLdapConflict As Boolean
Private mIsLdapNotfound As Boolean
Private mIsFunctionRegionConflict As Boolean

Public Function Init(Optional mss As SystemSetting)
    If mss Is Nothing Then
        Set ss = Session.Settings()
    Else
        Set ss = mss
    End If
    Set dbm = New DbManager
End Function

Public Function IsExistUserData() As Boolean
    Dim query As String
    query = "select * from " & Constants.END_USER_DATA_TABLE_NAME & " where deleted = 0 and SFunction='" & _
                Session.CurrentUser.FuncRegion.name & "'"
                
    dbm.Init
    dbm.OpenRecordSet query
    mIsLdapNotfound = False
    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
        IsExistUserData = True
    Else
        IsExistUserData = False
    End If
    dbm.Recycle
End Function

Public Function CheckRegionFunction()
    Dim query As String
    Dim tmpNtid As String
    query = "SELECT * FROM " & Constants.END_USER_DATA_CACHE_TABLE_NAME & " WHERE [" & Constants.FIELD_REGION_FUNCTION & "] not like '" & StringHelper.EscapeQueryString(Session.CurrentUser.FuncRegion.name) & "'"
    dbm.Init
    dbm.OpenRecordSet query
    mIsFunctionRegionConflict = False
    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
        mIsFunctionRegionConflict = True
        dbm.RecordSet.MoveFirst
        Do Until dbm.RecordSet.EOF = True
            tmpNtid = dbm.GetFieldValue(dbm.RecordSet, Session.Settings.NtidField)
            dbm.ExecuteQuery "DELETE FROM " & Constants.TABLE_USER_DATA_LDAP_CONFLICT & " WHERE [" & Session.Settings.NtidField & "] = '" & StringHelper.EscapeQueryString(tmpNtid) & "'"
            dbm.ExecuteQuery "DELETE FROM " & Constants.TABLE_USER_DATA_LDAP_NOTFOUND & " WHERE [" & Session.Settings.NtidField & "] = '" & StringHelper.EscapeQueryString(tmpNtid) & "'"
            dbm.RecordSet.MoveNext
        Loop
    End If
    dbm.ExecuteQuery "DELETE FROM " & Constants.END_USER_DATA_CACHE_TABLE_NAME & " WHERE [" & Constants.FIELD_REGION_FUNCTION & "] not like '" & StringHelper.EscapeQueryString(Session.CurrentUser.FuncRegion.name) & "' Or " & Session.Settings.NtidField & " is null"
    dbm.Recycle
End Function

Public Function CheckLdapNotFound()
    Dim ntid As String
    Dim query As String
    Dim tmpSelect As String
    query = "SELECT * FROM " & Constants.TABLE_USER_DATA_LDAP_NOTFOUND
    dbm.Init
    dbm.OpenRecordSet query
    mIsLdapNotfound = False
    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
        mIsLdapNotfound = True
    End If
    dbm.Recycle
End Function

Public Function CheckLdapConfict()
    Dim query As String
    Dim tmpSelect As String
    query = "SELECT * FROM " & Constants.TABLE_USER_DATA_LDAP_CONFLICT
    dbm.Init
    dbm.OpenRecordSet query
    mIsLdapConflict = False
    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
        mIsLdapConflict = True
    End If
    dbm.Recycle
End Function

Public Function CheckConflict()
    Dim tblCols As New Collection
    Dim lastUserData As Scripting.Dictionary
    Dim tmpCol As String
    Dim tmpInsertCols As Collection
    Dim tmpInsertData As Scripting.Dictionary
    Dim checkDict As Scripting.Dictionary
    Dim i As Integer
    Dim v As Variant
    Dim lastNtid As String, ntid As String
    Dim str1 As String, str2 As String
    Dim name As String
    Dim check As Boolean
    Dim tmpRst As DAO.RecordSet
    Dim qdf As DAO.QueryDef
    Dim query As String
    mIsConflict = False
    query = "SELECT * FROM " & Constants.END_USER_DATA_CACHE_TABLE_NAME
    Logger.LogDebug "UserManagement.CheckConflict", "Start check conflict records. Query: " & query
    dbm.RecycleTableName Constants.TABLE_USER_DATA_CONFLICT
    dbm.Init
    dbm.OpenRecordSet query
    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
        Set tmpInsertCols = New Collection
        tmpInsertCols.Add "NTID"
        tmpInsertCols.Add "Name"
        tmpInsertCols.Add "Field heading"
        tmpInsertCols.Add "Db field"
        tmpInsertCols.Add "Upload file"
        tmpInsertCols.Add "Data held"
        tmpInsertCols.Add "Select"
        For i = 0 To dbm.RecordSet.fields.Count - 1
            tmpCol = dbm.RecordSet.fields(i).name
             If (Not StringHelper.IsEqual(tmpCol, Constants.FIELD_ID, True)) _
                   And (Not StringHelper.IsEqual(tmpCol, Constants.FIELD_TIMESTAMP, True)) _
                   And (Not StringHelper.IsEqual(tmpCol, Constants.FIELD_DELETED, True)) _
                   And (Not StringHelper.IsEqual(tmpCol, ss.NtidField, True)) _
                   And (Not StringHelper.IsEqual(tmpCol, Constants.END_USER_DATA_CACHE_TABLE_NAME & "." _
                    & ss.NtidField, True)) Then
                 tblCols.Add tmpCol
             End If
        Next i
        dbm.RecordSet.MoveFirst
        Do Until dbm.RecordSet.EOF = True
            ntid = dbm.GetFieldValue(dbm.RecordSet, ss.NtidField)
            
            query = "SELECT * FROM " & Constants.END_USER_DATA_TABLE_NAME _
                                                        & " WHERE " & ss.NtidField & " = '" _
                                                        & StringHelper.EscapeQueryString(ntid) & "'"
            'Logger.LogDebug "UserManagement.CheckConflict", "Compare NTID query: " & query
            Set qdf = dbm.Database.CreateQueryDef("", query)
            Set tmpRst = qdf.OpenRecordSet
            If Not (tmpRst.EOF And tmpRst.BOF) Then
                tmpRst.MoveFirst
                For Each v In tblCols
                    str1 = dbm.GetFieldValue(dbm.RecordSet, CStr(v))
                    str2 = dbm.GetFieldValue(tmpRst, CStr(v))
                    If Not StringHelper.IsEqual(str1, str2, True) Then
                        Set tmpInsertData = New Scripting.Dictionary
                        tmpInsertData.Add "NTID", ntid
                        tmpInsertData.Add "Name", dbm.GetFieldValue(dbm.RecordSet, Constants.FIELD_LAST_NAME) & " " & dbm.GetFieldValue(dbm.RecordSet, Constants.FIELD_FIRST_NAME)
                        tmpInsertData.Add "Field heading", ss.SyncUsers.Item(CStr(v))
                        tmpInsertData.Add "Db field", CStr(v)
                        tmpInsertData.Add "Upload file", str1
                        tmpInsertData.Add "Data held", str2
                        tmpInsertData.Add "Select", "-1"
                        dbm.CreateLocalRecord tmpInsertData, tmpInsertCols, Constants.TABLE_USER_DATA_CONFLICT
                        mIsConflict = True
                    End If
                Next v
            End If
            tmpRst.Close
            dbm.RecordSet.MoveNext
        Loop
    Else
        Logger.LogInfo "UserManagement.CheckConflict", "There are no records in table " & Constants.END_USER_DATA_CACHE_TABLE_NAME
    End If
    dbm.Recycle
End Function

Public Function CheckDuplicate()
    Dim tblCols As New Collection
    Dim lastUserData As Scripting.Dictionary
    Dim tmpCol As String
    Dim tmpInsertCols As Collection
    Dim tmpInsertData As Scripting.Dictionary
    Dim checkDict As Scripting.Dictionary
    Dim i As Integer
    Dim v As Variant
    Dim lastNtid As String, ntid As String
    Dim str1 As String, str2 As String
    Dim name As String
    Dim check As Boolean
    Dim query As String
    mIsDuplicate = False
    query = "SELECT * FROM (" _
                    & Constants.END_USER_DATA_CACHE_TABLE_NAME & " INNER JOIN ( SELECT  " _
                    & ss.NtidField & " FROM " _
                    & Constants.END_USER_DATA_CACHE_TABLE_NAME & " GROUP BY " _
                    & ss.NtidField & " HAVING COUNT(*) > 1) As dupe ON dupe." _
                    & ss.NtidField & " = " _
                    & Constants.END_USER_DATA_CACHE_TABLE_NAME & "." _
                    & ss.NtidField & ")" _
                    & " ORDER BY " _
                    & Constants.END_USER_DATA_CACHE_TABLE_NAME & "." _
                    & ss.NtidField
    Logger.LogDebug "UserManagement.CheckDuplicate", "Start check duplicate " & ss.NtidField & " records. Query: " & query
    dbm.RecycleTableName Constants.TABLE_USER_DATA_DUPLICATE
    dbm.Init
    
    dbm.OpenRecordSet query
    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
        Set tmpInsertCols = New Collection
        tmpInsertCols.Add "NTID"
        tmpInsertCols.Add "Name"
        tmpInsertCols.Add "Field heading"
        tmpInsertCols.Add "Db field"
        tmpInsertCols.Add "Upload file"
        tmpInsertCols.Add "Select"
        For i = 0 To dbm.RecordSet.fields.Count - 1
            tmpCol = dbm.RecordSet.fields(i).name
             If (Not StringHelper.IsEqual(tmpCol, Constants.FIELD_ID, True)) _
                   And (Not StringHelper.IsEqual(tmpCol, Constants.FIELD_TIMESTAMP, True)) _
                   And (Not StringHelper.IsEqual(tmpCol, Constants.FIELD_DELETED, True)) _
                   And (Not StringHelper.IsEqual(tmpCol, ss.NtidField, True)) _
                   And (Not StringHelper.IsEqual(tmpCol, Constants.END_USER_DATA_CACHE_TABLE_NAME & "." _
                    & ss.NtidField, True)) Then
                 tblCols.Add tmpCol
             End If
        Next i
        check = False
        dbm.RecordSet.MoveFirst
        Do Until dbm.RecordSet.EOF = True
            If Not lastUserData Is Nothing Then
                ntid = dbm.GetFieldValue(dbm.RecordSet, Constants.END_USER_DATA_CACHE_TABLE_NAME & "." _
                    & ss.NtidField)
                Logger.LogDebug "UserManagement.CheckDuplicate", "lastNtid: " & lastNtid & ". Current Ntid: " & ntid
                If StringHelper.IsEqual(ntid, lastNtid, True) Then
                    For Each v In tblCols
                        str1 = lastUserData.Item(CStr(v))
                        str2 = dbm.GetFieldValue(dbm.RecordSet, CStr(v))
                        If Not StringHelper.IsEqual(str1, str2, True) Then
                            Set tmpInsertData = New Scripting.Dictionary
                            tmpInsertData.Add "NTID", ntid
                            tmpInsertData.Add "Name", dbm.GetFieldValue(dbm.RecordSet, Constants.FIELD_LAST_NAME) & " " & dbm.GetFieldValue(dbm.RecordSet, Constants.FIELD_FIRST_NAME)
                            tmpInsertData.Add "Field heading", ss.SyncUsers.Item(CStr(v))
                            tmpInsertData.Add "Db field", CStr(v)
                            tmpInsertData.Add "Upload file", str2
                            tmpInsertData.Add "Select", "0"
                            dbm.CreateLocalRecord tmpInsertData, tmpInsertCols, Constants.TABLE_USER_DATA_DUPLICATE
                            mIsDuplicate = True
                            If checkDict Is Nothing Then
                                Set checkDict = New Scripting.Dictionary
                            End If
                            If Not checkDict.Exists(CStr(v)) Then
                                Set tmpInsertData = New Scripting.Dictionary
                                tmpInsertData.Add "NTID", ntid
                                tmpInsertData.Add "Name", lastUserData.Item(Constants.FIELD_LAST_NAME) & " " & lastUserData.Item(dbm.GetFieldValue(dbm.RecordSet, Constants.FIELD_FIRST_NAME))
                                tmpInsertData.Add "Field heading", ss.SyncUsers.Item(CStr(v))
                                tmpInsertData.Add "Db field", CStr(v)
                                tmpInsertData.Add "Upload file", str1
                                tmpInsertData.Add "Select", "-1"
                                dbm.CreateLocalRecord tmpInsertData, tmpInsertCols, Constants.TABLE_USER_DATA_DUPLICATE
                                mIsDuplicate = True
                                checkDict.Add CStr(v), True
                            End If
                        End If
                    Next v
                    'check = True
                Else
                   ' check = False
                   Set checkDict = New Scripting.Dictionary
                End If
            End If
            Set lastUserData = New Scripting.Dictionary
            For Each v In tblCols
                lastUserData.Add CStr(v), dbm.GetFieldValue(dbm.RecordSet, CStr(v))
            Next v
            lastNtid = dbm.GetFieldValue(dbm.RecordSet, Constants.END_USER_DATA_CACHE_TABLE_NAME & "." _
                    & ss.NtidField)
            dbm.RecordSet.MoveNext
        Loop
    Else
        Logger.LogInfo "UserManagement.CheckDuplicate", "There are no duplicate " & ss.NtidField & " records in table " & Constants.END_USER_DATA_CACHE_TABLE_NAME
    End If
    dbm.Recycle
End Function


Public Function ResolveLdapNotFound()
    Dim ntid As String
    Dim query As String
    Dim tmpSelect As String
    query = "SELECT * FROM " & Constants.TABLE_USER_DATA_LDAP_NOTFOUND
    dbm.Init
    dbm.OpenRecordSet query
    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
        dbm.RecordSet.MoveFirst
        Do Until dbm.RecordSet.EOF = True
            ntid = dbm.GetFieldValue(dbm.RecordSet, ss.NtidField)
            tmpSelect = dbm.GetFieldValue(dbm.RecordSet, Constants.FIELD_SELECT)
            Logger.LogDebug "UserManagement.ResolveLdapNotFound", "Select: " & tmpSelect
            If Not StringHelper.IsEqual(tmpSelect, "false", True) Then
                Logger.LogDebug "UserManagement.ResolveLdapNotFound", "delete user NTID " & ntid & " from cache"
                query = "DELETE FROM " & Constants.END_USER_DATA_CACHE_TABLE_NAME & " WHERE " & ss.NtidField & " = '" & StringHelper.EscapeQueryString(ntid) & "'"
                dbm.ExecuteQuery query
            End If
            dbm.RecordSet.MoveNext
        Loop
    Else
        Logger.LogInfo "UserManagement.ResolveLdapNotFound", "There are no selected record in table " & Constants.TABLE_USER_DATA_LDAP_NOTFOUND
    End If
    dbm.RecycleTableName Constants.TABLE_USER_DATA_LDAP_NOTFOUND
    dbm.Recycle
End Function


Public Function ResolveLdapConflict()
    Dim ntid As String
    Dim dbField As String
    Dim tmpValue As String
    Dim query As String
    Dim tmpSelect As String
    query = "SELECT * FROM " & Constants.TABLE_USER_DATA_LDAP_CONFLICT
    dbm.Init
    dbm.OpenRecordSet query
    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
        dbm.RecordSet.MoveFirst
        Do Until dbm.RecordSet.EOF = True
            tmpSelect = dbm.GetFieldValue(dbm.RecordSet, Constants.FIELD_SELECT)
            Logger.LogDebug "UserManagement.ResolveLdapConflict", "Select: " & tmpSelect
            If Not StringHelper.IsEqual(tmpSelect, "false", True) Then
                ntid = dbm.GetFieldValue(dbm.RecordSet, ss.NtidField)
                dbField = dbm.GetFieldValue(dbm.RecordSet, Constants.FIELD_DB_FIELD)
                tmpValue = dbm.GetFieldValue(dbm.RecordSet, "LDAP")
                Logger.LogDebug "UserManagement.ResolveLdapConflict", "Resolve confict user NTID " & ntid & ".Db field: " & dbField & " . New value: " & tmpValue
                query = "UPDATE " & Constants.END_USER_DATA_CACHE_TABLE_NAME & " SET [" & dbField & "] = '" & StringHelper.EscapeQueryString(tmpValue) & "' WHERE " _
                            & ss.NtidField & " = '" & StringHelper.EscapeQueryString(ntid) & "'"
                dbm.ExecuteQuery query
            End If
            dbm.RecordSet.MoveNext
        Loop
    Else
        Logger.LogInfo "UserManagement.ResolveLdapConflict", "There are no selected record in table " & Constants.TABLE_USER_DATA_LDAP_CONFLICT
    End If
    dbm.RecycleTableName Constants.TABLE_USER_DATA_LDAP_CONFLICT
    dbm.Recycle
End Function


Public Function ResolveUserDataConflict()
    Dim ntid As String
    Dim dbField As String
    Dim tmpValue As String
    Dim query As String
    Dim tmpSelect As String
    query = "SELECT * FROM " & Constants.TABLE_USER_DATA_CONFLICT
    dbm.Init
    dbm.OpenRecordSet query
    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
        dbm.RecordSet.MoveFirst
        Do Until dbm.RecordSet.EOF = True
            tmpSelect = dbm.GetFieldValue(dbm.RecordSet, Constants.FIELD_SELECT)
            Logger.LogDebug "UserManagement.ResolveUserDataConflict", "Select: " & tmpSelect
            If StringHelper.IsEqual(tmpSelect, "false", True) Then
                ntid = dbm.GetFieldValue(dbm.RecordSet, ss.NtidField)
                dbField = dbm.GetFieldValue(dbm.RecordSet, Constants.FIELD_DB_FIELD)
                tmpValue = dbm.GetFieldValue(dbm.RecordSet, "Data held")
                Logger.LogDebug "UserManagement.ResolveUserDataConflict", "Resolve confict user NTID " & ntid & ".Db field: " & dbField & " . New value: " & tmpValue
                query = "UPDATE " & Constants.END_USER_DATA_CACHE_TABLE_NAME & " SET [" & dbField & "] = '" & StringHelper.EscapeQueryString(tmpValue) & "' WHERE " _
                                & ss.NtidField & " = '" & StringHelper.EscapeQueryString(ntid) & "'"
                dbm.ExecuteQuery query
            End If
            dbm.RecordSet.MoveNext
        Loop
        
    Else
        Logger.LogInfo "UserManagement.ResolveUserDataConflict", "There are no selected record in table " & Constants.TABLE_USER_DATA_CONFLICT
    End If
    dbm.RecycleTableName Constants.TABLE_USER_DATA_CONFLICT
    dbm.Recycle
End Function


Public Function ResolveUserDataDuplicate()
    Dim ntid As String
    Dim dbField As String
    Dim tmpValue As String
    Dim query As String
    Dim tmpCol As Collection
    Dim lastNtid As String
    Dim c As Integer
    Dim v As Variant
    Dim tmpSelect As String
    query = "SELECT * FROM " & Constants.TABLE_USER_DATA_DUPLICATE
    dbm.Init
    dbm.OpenRecordSet query
    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
        dbm.RecordSet.MoveFirst
        Do Until dbm.RecordSet.EOF = True
            tmpSelect = dbm.GetFieldValue(dbm.RecordSet, Constants.FIELD_SELECT)
            Logger.LogDebug "UserManagement.ResolveUserDataDuplicate", "Select: " & tmpSelect
            If StringHelper.IsEqual(tmpSelect, "false", True) Then
                ntid = dbm.GetFieldValue(dbm.RecordSet, ss.NtidField)
                dbField = dbm.GetFieldValue(dbm.RecordSet, Constants.FIELD_DB_FIELD)
                tmpValue = dbm.GetFieldValue(dbm.RecordSet, "Upload file")
                Logger.LogDebug "UserManagement.ResolveUserDataDuplicate", "Resolve confict user NTID " & ntid & ".Db field: " & dbField & " . New value: " & tmpValue
                query = "UPDATE " & Constants.END_USER_DATA_CACHE_TABLE_NAME & " SET [" & dbField & "] = '" & StringHelper.EscapeQueryString(tmpValue) & "' WHERE " _
                            & ss.NtidField & " = '" & StringHelper.EscapeQueryString(ntid) & "'"
                dbm.ExecuteQuery query
            End If
            dbm.RecordSet.MoveNext
        Loop
    Else
        Logger.LogInfo "UserManagement.ResolveUserDataDuplicate", "There are no selected record in table " & Constants.TABLE_USER_DATA_DUPLICATE
    End If
    dbm.Recycle
    query = "select * from (select *, (select count(" _
                & ss.NtidField & ") from " _
                & Constants.END_USER_DATA_CACHE_TABLE_NAME & " where " _
                & Constants.END_USER_DATA_CACHE_TABLE_NAME & "." _
                & ss.NtidField & " =UD." _
                & ss.NtidField & ") AS count_ntid from " _
                & Constants.END_USER_DATA_CACHE_TABLE_NAME & " AS UD) as tmp_tbl where count_ntid > 1 order by tmp_tbl." & ss.NtidField
    dbm.Init
    dbm.OpenRecordSet query
    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
        Set tmpCol = New Collection
        dbm.RecordSet.MoveFirst
        c = 0
        Do Until dbm.RecordSet.EOF = True
            Logger.LogInfo "UserManagement.ResolveUserDataDuplicate", "Duplicate ntid: " & ntid
            ntid = dbm.GetFieldValue(dbm.RecordSet, ss.NtidField)
            If StringHelper.IsEqual(ntid, lastNtid, True) Then
                If c <> 0 Then
                    tmpCol.Add dbm.GetFieldValue(dbm.RecordSet, Constants.FIELD_ID)
                End If
            Else
                c = 0
            End If
            c = c + 1
            lastNtid = ntid
            dbm.RecordSet.MoveNext
        Loop
        For Each v In tmpCol
            Logger.LogInfo "UserManagement.ResolveUserDataDuplicate", "Delete duplicate id: " & CStr(v)
            query = "DELETE FROM " & Constants.END_USER_DATA_CACHE_TABLE_NAME & " WHERE " & Constants.FIELD_ID & " = '" & StringHelper.EscapeQueryString(CStr(v)) & "'"
            dbm.ExecuteQuery query
        Next v
    Else
        Logger.LogInfo "UserManagement.ResolveUserDataDuplicate", "There are no duplicate record in table " & Constants.END_USER_DATA_CACHE_TABLE_NAME
    End If
    
    dbm.RecycleTableName Constants.TABLE_USER_DATA_DUPLICATE
    dbm.Recycle
End Function

Public Function ListSpecialism() As Collection
    Dim list As New Collection
    Dim query As String
    query = "SELECT [" & Constants.FIELD_SPECIALISM & "] from " & Constants.END_USER_DATA_CACHE_TABLE_NAME _
                    & " group by [" & Constants.FIELD_SPECIALISM & "]"
    dbm.Init
    dbm.OpenRecordSet query
    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
        dbm.RecordSet.MoveFirst
        Do Until dbm.RecordSet.EOF = True
            list.Add dbm.GetFieldValue(dbm.RecordSet, Constants.FIELD_SPECIALISM)
            dbm.RecordSet.MoveNext
        Loop
    End If
    dbm.Recycle
    Set ListSpecialism = list
End Function

Public Function GenereateSpecialismFilter() As String
    Dim list As Collection
    Set list = ListSpecialism
    Dim v As Variant
    Dim filter As String
    filter = ""
    For Each v In list
        filter = filter & "'" & StringHelper.EscapeQueryString(CStr(v)) & "',"
    Next v
    If StringHelper.EndsWith(filter, ",", True) Then
        filter = Left(filter, Len(filter) - 1)
    End If
    GenereateSpecialismFilter = filter
End Function

Public Function GenerateRoleMapping(rm As ReportMetaData)
    Dim i As Integer, j As Long, l As Long, k As Long
    Dim v As Variant, y As Variant
    Dim startMappingCol As Long
    Dim mappingCount As Long
    Dim tmpNtid As String
    Dim tmpRoleId As String
    Dim tmpStr As String
    Dim tmpRole As String
    Dim tmpQuery As String
    Dim oExcel As New Excel.Application
    Dim WB As New Excel.Workbook
    Dim WS As Excel.WorkSheet
    Dim rng As Excel.range
    Dim tmpRps As ReportSection
    Dim header() As String
    Dim isUpdate As Boolean
    i = 0
    For Each tmpRps In rm.ReportSections
        If i = 0 Then
            startMappingCol = rm.StartCol + tmpRps.HeaderCount
        ElseIf i = 1 Then
            header = tmpRps.header
            mappingCount = tmpRps.HeaderCount
        Else
        End If
        i = i + 1
    Next
    
    Logger.LogDebug "UserManagement.GenerateRoleMapping", "Mapping Cols count : " & CStr(mappingCount) _
                                    & ". Start Mapping col: " & CStr(startMappingCol)
    
    With oExcel
            .DisplayAlerts = False
            .Visible = False
            Logger.LogDebug "UserManagement.GenerateRoleMapping", "Open excel : " & rm.OutputPath
            Set WB = .Workbooks.Open(rm.OutputPath)
            With WB
                Logger.LogDebug "UserManagement.GenerateRoleMapping", "Select worksheet: " & rm.WorkSheet
                Set WS = WB.workSheets(rm.WorkSheet)
                l = rm.StartCol
                k = rm.StartRow
                With WS
                    Do While k < 65536
                        Set rng = .Cells(k, l)
                        tmpNtid = rng.value
                        If Len(Trim(tmpNtid)) = 0 Then
                            Exit Do
                        End If
                        Logger.LogDebug "UserManagement.GenerateRoleMapping", "Found ntid: " & tmpNtid
                        
                        dbm.Init
                        ' Remove all mapping
                        dbm.ExecuteQuery "update user_data_mapping_role set Deleted=-1 where idUserdata='" & StringHelper.EscapeQueryString(tmpNtid) _
                                            & "' and idFunction='" & StringHelper.EscapeQueryString(Session.CurrentUser.FuncRegion.FuncRgID) & "'"
                        
                        dbm.Recycle
                        For i = 0 To mappingCount - 1
                            j = startMappingCol + i
                            Set rng = .Cells(rm.StartHeaderRow, j)
                            tmpRole = rng.value
                            'Logger.LogDebug "UserManagement.GenerateRoleMapping", "Found role: " & tmpRole
                            ' Get mapping check
                            Set rng = .Cells(k, j)
                            tmpStr = Trim(rng.value)
                            If Len(tmpStr) <> 0 Then
                                dbm.Init
                                dbm.OpenRecordSet "select BR.[id] from BpRoleStandard As BR Where BR.BpRoleStandardName='" _
                                        & StringHelper.EscapeQueryString(tmpRole) & "' and BR.deleted=0"
                                If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
                                    tmpRoleId = dbm.GetFieldValue(dbm.RecordSet, "id")
                                Else
                                    tmpRoleId = ""
                                End If
                                dbm.Recycle
                                If Len(tmpRoleId) <> 0 Then
                                    Logger.LogDebug "UserManagement.GenerateRoleMapping", "Found mapping Ntid: " & tmpNtid & " to Role: " & tmpRole
                                    
                                    dbm.Init
                                    tmpQuery = "select * from user_data_mapping_role where idUserdata='" & StringHelper.EscapeQueryString(tmpNtid) _
                                                & "' and idFunction='" & StringHelper.EscapeQueryString(Session.CurrentUser.FuncRegion.FuncRgID) & "'" _
                                                & " and idBpRoleStandard='" & StringHelper.EscapeQueryString(tmpRoleId) & "'"
                                    dbm.OpenRecordSet tmpQuery
                                    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
                                        tmpQuery = "update user_data_mapping_role set Deleted=0 where idUserdata='" & StringHelper.EscapeQueryString(tmpNtid) _
                                                & "' and idFunction='" & StringHelper.EscapeQueryString(Session.CurrentUser.FuncRegion.FuncRgID) & "'" _
                                                & " and idBpRoleStandard='" & StringHelper.EscapeQueryString(tmpRoleId) & "'"
                                        ' Update mapping
                                        Logger.LogDebug "UserManagement.GenerateRoleMapping", "Update mapping. Query: " & tmpQuery
                                        dbm.ExecuteQuery tmpQuery
                                    Else
                                        ' Insert mapping
                                        tmpQuery = "insert into user_data_mapping_role(id, idUserdata, idFunction, idBpRoleStandard, Deleted) values('" _
                                                & StringHelper.EscapeQueryString(StringHelper.GetGUID) & "', '" _
                                                & StringHelper.EscapeQueryString(tmpNtid) & "', '" _
                                                & StringHelper.EscapeQueryString(Session.CurrentUser.FuncRegion.FuncRgID) & "', '" _
                                                & StringHelper.EscapeQueryString(tmpRoleId) & "', 0)"
                                                
                                        Logger.LogDebug "UserManagement.GenerateRoleMapping", "Insert mapping. Query: " & tmpQuery
                                        dbm.ExecuteQuery tmpQuery
                                    End If
                                    dbm.Recycle
                                    
                                Else
                                    Logger.LogError "UserManagement.GenerateRoleMapping", "Found mapping Ntid: " & tmpNtid & " to Role: " & tmpRole & ". But role ID not found", Nothing
                                End If
                                
                            End If
                        Next i
                        k = k + 1
                    Loop
                End With
                Logger.LogDebug "UserManagement.GenerateRoleMapping", "Close excel file " & rm.OutputPath
            End With
            .Quit
        End With
    dbm.Recycle
End Function
 
Public Function MergeUserData()
    Dim ntid As String
    Dim v As Variant
    Dim check As Boolean
    Dim tmpNtid As String
    Dim tmpId As String
    Dim dbField As String
    Dim tmpValue As String
    Dim tmpRst As DAO.RecordSet
    Dim tmpQdf As DAO.QueryDef
    Dim tmpCols As Collection
    Dim tmpData As Scripting.Dictionary
    Dim tmpCol As String
    Dim query As String
    
    query = "SELECT * FROM " & Constants.END_USER_DATA_CACHE_TABLE_NAME
    dbm.Init
    dbm.OpenRecordSet query
    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
        dbm.RecordSet.MoveFirst
        Do Until dbm.RecordSet.EOF = True
            Set tmpCols = New Collection
            For i = 0 To dbm.RecordSet.fields.Count - 1
                tmpCol = dbm.RecordSet.fields(i).name
                If (Not StringHelper.IsEqual(tmpCol, Constants.FIELD_TIMESTAMP, True)) _
                    And (Not StringHelper.IsEqual(tmpCol, Constants.FIELD_DELETED, True)) _
                    And (Not StringHelper.IsEqual(tmpCol, Constants.FIELD_ID, True)) Then
                     tmpCols.Add tmpCol
                End If
            Next i
            Set tmpData = New Scripting.Dictionary
            For Each v In tmpCols
                tmpData.Add CStr(v), dbm.GetFieldValue(dbm.RecordSet, CStr(v))
            Next v
            tmpCols.Add FIELD_DELETED
            tmpCols.Add FIELD_ID
            tmpData.Add Constants.FIELD_DELETED, "0"
            ntid = dbm.GetFieldValue(dbm.RecordSet, ss.NtidField)
            query = "SELECT * FROM " & Constants.END_USER_DATA_TABLE_NAME & " WHERE " & ss.NtidField _
                    & " = '" & StringHelper.EscapeQueryString(ntid) & "'"
                    
            Set tmpQdf = dbm.Database.CreateQueryDef("", query)
            Set tmpRst = tmpQdf.OpenRecordSet
            
            If Not (tmpRst.EOF And tmpRst.BOF) Then
                tmpRst.MoveFirst
                tmpNtid = dbm.GetFieldValue(tmpRst, ss.NtidField)
                tmpId = dbm.GetFieldValue(tmpRst, Constants.FIELD_ID)
                Logger.LogDebug "UserManagement.MergeUserData", "Update old record NTID" & tmpNtid & ". ID: " & tmpId
                tmpData.Add Constants.FIELD_ID, tmpId
                dbm.UpdateLocalRecord tmpData, tmpCols, Constants.END_USER_DATA_TABLE_NAME
            Else
                tmpData.Add Constants.FIELD_ID, StringHelper.GetGUID
                Logger.LogDebug "UserManagement.MergeUserData", "Create new record"
                dbm.CreateLocalRecord tmpData, tmpCols, Constants.END_USER_DATA_TABLE_NAME
            End If
            dbm.RecordSet.MoveNext
        Loop
    Else
        Logger.LogInfo "UserManagement.MergeUserData", "There are no record in table " & Constants.END_USER_DATA_CACHE_TABLE_NAME
    End If
    dbm.Recycle
End Function

Public Property Get IsConflict() As Boolean
    IsConflict = mIsConflict
End Property

Public Property Get IsDuplicate() As Boolean
    IsDuplicate = mIsDuplicate
End Property

Public Property Get IsLdapConflict() As Boolean
    IsLdapConflict = mIsLdapConflict
End Property

Public Property Get IsLdapNotfound() As Boolean
    IsLdapNotfound = mIsLdapNotfound
End Property

Public Property Get IsFunctionRegionConflict() As Boolean
    IsFunctionRegionConflict = mIsFunctionRegionConflict
End Property