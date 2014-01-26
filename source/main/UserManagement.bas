Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Private ss As SystemSettings
Private dbm As DbManager

Public Function Init(mss As SystemSettings)
    Set ss = mss
    Set dbm = New DbManager
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
    Dim lastNtid As String, NTID As String
    Dim str1 As String, str2 As String
    Dim name As String
    Dim check As Boolean
    Dim tmpRst As DAO.RecordSet
    Dim qdf As DAO.QueryDef
    Dim Query As String
    
    Query = "SELECT * FROM " & Constants.END_USER_DATA_CACHE_TABLE_NAME
    Logger.LogDebug "UserManagement.CheckConflict", "Start check conflict records. Query: " & Query
    dbm.RecycleTableName Constants.TABLE_USER_DATA_CONFLICT
    dbm.Init
    dbm.OpenRecordSet Query
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
            NTID = dbm.GetFieldValue(dbm.RecordSet, ss.NtidField)
            
            Query = "SELECT * FROM " & Constants.END_USER_DATA_TABLE_NAME _
                                                        & " WHERE " & ss.NtidField & " = '" _
                                                        & StringHelper.EscapeQueryString(NTID) & "'"
            'Logger.LogDebug "UserManagement.CheckConflict", "Compare NTID query: " & query
            Set qdf = dbm.Database.CreateQueryDef("", Query)
            Set tmpRst = qdf.OpenRecordSet
            If Not (tmpRst.EOF And tmpRst.BOF) Then
                tmpRst.MoveFirst
                For Each v In tblCols
                    str1 = dbm.GetFieldValue(dbm.RecordSet, CStr(v))
                    str2 = dbm.GetFieldValue(tmpRst, CStr(v))
                    If Not StringHelper.IsEqual(str1, str2, True) Then
                        Set tmpInsertData = New Scripting.Dictionary
                        tmpInsertData.Add "NTID", NTID
                        tmpInsertData.Add "Name", dbm.GetFieldValue(dbm.RecordSet, Constants.FIELD_LAST_NAME) & " " & dbm.GetFieldValue(dbm.RecordSet, Constants.FIELD_FIRST_NAME)
                        tmpInsertData.Add "Field heading", ss.SyncUsers.Item(CStr(v))
                        tmpInsertData.Add "Db field", CStr(v)
                        tmpInsertData.Add "Upload file", str1
                        tmpInsertData.Add "Data held", str2
                        tmpInsertData.Add "Select", "-1"
                        dbm.CreateLocalRecord tmpInsertData, tmpInsertCols, Constants.TABLE_USER_DATA_CONFLICT
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
    Dim lastNtid As String, NTID As String
    Dim str1 As String, str2 As String
    Dim name As String
    Dim check As Boolean
    Dim Query As String
    
    Query = "SELECT * FROM (" _
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
    Logger.LogDebug "UserManagement.CheckDuplicate", "Start check duplicate " & ss.NtidField & " records. Query: " & Query
    dbm.RecycleTableName Constants.TABLE_USER_DATA_DUPLICATE
    dbm.Init
    
    dbm.OpenRecordSet Query
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
                NTID = dbm.GetFieldValue(dbm.RecordSet, Constants.END_USER_DATA_CACHE_TABLE_NAME & "." _
                    & ss.NtidField)
                Logger.LogDebug "UserManagement.CheckDuplicate", "lastNtid: " & lastNtid & ". Current Ntid: " & NTID
                If StringHelper.IsEqual(NTID, lastNtid, True) Then
                    For Each v In tblCols
                        str1 = lastUserData.Item(CStr(v))
                        str2 = dbm.GetFieldValue(dbm.RecordSet, CStr(v))
                        If Not StringHelper.IsEqual(str1, str2, True) Then
                            Set tmpInsertData = New Scripting.Dictionary
                            tmpInsertData.Add "NTID", NTID
                            tmpInsertData.Add "Name", dbm.GetFieldValue(dbm.RecordSet, Constants.FIELD_LAST_NAME) & " " & dbm.GetFieldValue(dbm.RecordSet, Constants.FIELD_FIRST_NAME)
                            tmpInsertData.Add "Field heading", ss.SyncUsers.Item(CStr(v))
                            tmpInsertData.Add "Db field", CStr(v)
                            tmpInsertData.Add "Upload file", str2
                            tmpInsertData.Add "Select", "0"
                            dbm.CreateLocalRecord tmpInsertData, tmpInsertCols, Constants.TABLE_USER_DATA_DUPLICATE
                            If checkDict Is Nothing Then
                                Set checkDict = New Scripting.Dictionary
                            End If
                            If Not checkDict.Exists(CStr(v)) Then
                                Set tmpInsertData = New Scripting.Dictionary
                                tmpInsertData.Add "NTID", NTID
                                tmpInsertData.Add "Name", lastUserData.Item(Constants.FIELD_LAST_NAME) & " " & lastUserData.Item(dbm.GetFieldValue(dbm.RecordSet, Constants.FIELD_FIRST_NAME))
                                tmpInsertData.Add "Field heading", ss.SyncUsers.Item(CStr(v))
                                tmpInsertData.Add "Db field", CStr(v)
                                tmpInsertData.Add "Upload file", str1
                                tmpInsertData.Add "Select", "-1"
                                dbm.CreateLocalRecord tmpInsertData, tmpInsertCols, Constants.TABLE_USER_DATA_DUPLICATE
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
    Dim NTID As String
    Dim Query As String
    Dim tmpSelect As String
    Query = "SELECT * FROM " & Constants.TABLE_USER_DATA_LDAP_NOTFOUND
    dbm.Init
    dbm.OpenRecordSet Query
    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
        dbm.RecordSet.MoveFirst
        Do Until dbm.RecordSet.EOF = True
            NTID = dbm.GetFieldValue(dbm.RecordSet, ss.NtidField)
            tmpSelect = dbm.GetFieldValue(dbm.RecordSet, Constants.FIELD_SELECT)
            Logger.LogDebug "UserManagement.ResolveLdapNotFound", "Select: " & tmpSelect
            If Not StringHelper.IsEqual(tmpSelect, "false", True) Then
                Logger.LogDebug "UserManagement.ResolveLdapNotFound", "delete user NTID " & NTID & " from cache"
                Query = "DELETE FROM " & Constants.END_USER_DATA_CACHE_TABLE_NAME & " WHERE " & ss.NtidField & " = '" & StringHelper.EscapeQueryString(NTID) & "'"
                dbm.ExecuteQuery Query
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
    Dim NTID As String
    Dim dbField As String
    Dim tmpValue As String
    Dim Query As String
    Dim tmpSelect As String
    Query = "SELECT * FROM " & Constants.TABLE_USER_DATA_LDAP_CONFLICT
    dbm.Init
    dbm.OpenRecordSet Query
    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
        dbm.RecordSet.MoveFirst
        Do Until dbm.RecordSet.EOF = True
            tmpSelect = dbm.GetFieldValue(dbm.RecordSet, Constants.FIELD_SELECT)
            Logger.LogDebug "UserManagement.ResolveLdapConflict", "Select: " & tmpSelect
            If Not StringHelper.IsEqual(tmpSelect, "false", True) Then
                NTID = dbm.GetFieldValue(dbm.RecordSet, ss.NtidField)
                dbField = dbm.GetFieldValue(dbm.RecordSet, Constants.FIELD_DB_FIELD)
                tmpValue = dbm.GetFieldValue(dbm.RecordSet, "LDAP")
                Logger.LogDebug "UserManagement.ResolveLdapConflict", "Resolve confict user NTID " & NTID & ".Db field: " & dbField & " . New value: " & tmpValue
                Query = "UPDATE " & Constants.END_USER_DATA_CACHE_TABLE_NAME & " SET [" & dbField & "] = '" & StringHelper.EscapeQueryString(tmpValue) & "' WHERE " _
                            & ss.NtidField & " = '" & StringHelper.EscapeQueryString(NTID) & "'"
                dbm.ExecuteQuery Query
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
    Dim NTID As String
    Dim dbField As String
    Dim tmpValue As String
    Dim Query As String
    Dim tmpSelect As String
    Query = "SELECT * FROM " & Constants.TABLE_USER_DATA_CONFLICT
    dbm.Init
    dbm.OpenRecordSet Query
    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
        dbm.RecordSet.MoveFirst
        Do Until dbm.RecordSet.EOF = True
            tmpSelect = dbm.GetFieldValue(dbm.RecordSet, Constants.FIELD_SELECT)
            Logger.LogDebug "UserManagement.ResolveUserDataConflict", "Select: " & tmpSelect
            If StringHelper.IsEqual(tmpSelect, "false", True) Then
                NTID = dbm.GetFieldValue(dbm.RecordSet, ss.NtidField)
                dbField = dbm.GetFieldValue(dbm.RecordSet, Constants.FIELD_DB_FIELD)
                tmpValue = dbm.GetFieldValue(dbm.RecordSet, "Data held")
                Logger.LogDebug "UserManagement.ResolveUserDataConflict", "Resolve confict user NTID " & NTID & ".Db field: " & dbField & " . New value: " & tmpValue
                Query = "UPDATE " & Constants.END_USER_DATA_CACHE_TABLE_NAME & " SET [" & dbField & "] = '" & StringHelper.EscapeQueryString(tmpValue) & "' WHERE " _
                                & ss.NtidField & " = '" & StringHelper.EscapeQueryString(NTID) & "'"
                dbm.ExecuteQuery Query
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
    Dim NTID As String
    Dim dbField As String
    Dim tmpValue As String
    Dim Query As String
    Dim tmpCol As Collection
    Dim lastNtid As String
    Dim c As Integer
    Dim v As Variant
    Dim tmpSelect As String
    Query = "SELECT * FROM " & Constants.TABLE_USER_DATA_DUPLICATE
    dbm.Init
    dbm.OpenRecordSet Query
    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
        dbm.RecordSet.MoveFirst
        Do Until dbm.RecordSet.EOF = True
            tmpSelect = dbm.GetFieldValue(dbm.RecordSet, Constants.FIELD_SELECT)
            Logger.LogDebug "UserManagement.ResolveUserDataDuplicate", "Select: " & tmpSelect
            If StringHelper.IsEqual(tmpSelect, "false", True) Then
                NTID = dbm.GetFieldValue(dbm.RecordSet, ss.NtidField)
                dbField = dbm.GetFieldValue(dbm.RecordSet, Constants.FIELD_DB_FIELD)
                tmpValue = dbm.GetFieldValue(dbm.RecordSet, "Upload file")
                Logger.LogDebug "UserManagement.ResolveUserDataDuplicate", "Resolve confict user NTID " & NTID & ".Db field: " & dbField & " . New value: " & tmpValue
                Query = "UPDATE " & Constants.END_USER_DATA_CACHE_TABLE_NAME & " SET [" & dbField & "] = '" & StringHelper.EscapeQueryString(tmpValue) & "' WHERE " _
                            & ss.NtidField & " = '" & StringHelper.EscapeQueryString(NTID) & "'"
                dbm.ExecuteQuery Query
            End If
            dbm.RecordSet.MoveNext
        Loop
    Else
        Logger.LogInfo "UserManagement.ResolveUserDataDuplicate", "There are no selected record in table " & Constants.TABLE_USER_DATA_DUPLICATE
    End If
    dbm.Recycle
    Query = "select * from (select *, (select count(" _
                & ss.NtidField & ") from " _
                & Constants.END_USER_DATA_CACHE_TABLE_NAME & " where " _
                & Constants.END_USER_DATA_CACHE_TABLE_NAME & "." _
                & ss.NtidField & " =UD." _
                & ss.NtidField & ") AS count_ntid from " _
                & Constants.END_USER_DATA_CACHE_TABLE_NAME & " AS UD) as tmp_tbl where count_ntid > 1 order by tmp_tbl." & ss.NtidField
    dbm.Init
    dbm.OpenRecordSet Query
    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
        Set tmpCol = New Collection
        dbm.RecordSet.MoveFirst
        c = 0
        Do Until dbm.RecordSet.EOF = True
            Logger.LogInfo "UserManagement.ResolveUserDataDuplicate", "Duplicate ntid: " & NTID
            NTID = dbm.GetFieldValue(dbm.RecordSet, ss.NtidField)
            If StringHelper.IsEqual(NTID, lastNtid, True) Then
                If c <> 0 Then
                    tmpCol.Add dbm.GetFieldValue(dbm.RecordSet, Constants.FIELD_ID)
                End If
            Else
                c = 0
            End If
            c = c + 1
            lastNtid = NTID
            dbm.RecordSet.MoveNext
        Loop
        For Each v In tmpCol
            Logger.LogInfo "UserManagement.ResolveUserDataDuplicate", "Delete duplicate id: " & CStr(v)
            Query = "DELETE FROM " & Constants.END_USER_DATA_CACHE_TABLE_NAME & " WHERE " & Constants.FIELD_ID & " = '" & StringHelper.EscapeQueryString(CStr(v)) & "'"
            dbm.ExecuteQuery Query
        Next v
    Else
        Logger.LogInfo "UserManagement.ResolveUserDataDuplicate", "There are no duplicate record in table " & Constants.END_USER_DATA_CACHE_TABLE_NAME
    End If
    
    dbm.RecycleTableName Constants.TABLE_USER_DATA_DUPLICATE
    dbm.Recycle
End Function

Public Function MergeUserData()
    Dim NTID As String
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
    Dim Query As String
    
    Query = "SELECT * FROM " & Constants.END_USER_DATA_CACHE_TABLE_NAME
    dbm.Init
    dbm.OpenRecordSet Query
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
            NTID = dbm.GetFieldValue(dbm.RecordSet, ss.NtidField)
            Query = "SELECT * FROM " & Constants.END_USER_DATA_TABLE_NAME & " WHERE " & ss.NtidField _
                    & " = '" & StringHelper.EscapeQueryString(NTID) & "'"
            Set tmpQdf = dbm.Database.CreateQueryDef("", Query)
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