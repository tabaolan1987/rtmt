Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Public Function CheckConflict()
    Dim ss As New SystemSettings
    Dim tblCols As New Collection
    Dim lastUserData As Scripting.Dictionary
    Dim tmpCol As String
    Dim tmpInsertCols As Collection
    Dim tmpInsertData As Scripting.Dictionary
    Dim checkDict As Scripting.Dictionary
    Dim dbm As New DbManager
    Dim i As Integer
    Dim v As Variant
    Dim lastNtid As String, ntid As String
    Dim str1 As String, str2 As String
    Dim Name As String
    Dim check As Boolean
    Dim tmpRst As DAO.RecordSet
    Dim qdf As DAO.QueryDef
    ss.Init
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
        For i = 0 To dbm.RecordSet.fields.count - 1
            tmpCol = dbm.RecordSet.fields(i).Name
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
            
            Query = "SELECT * FROM " & Constants.END_USER_DATA_TABLE_NAME _
                                                        & " WHERE " & ss.NtidField & " = '" _
                                                        & StringHelper.EscapeQueryString(ntid) & "'"
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
                        tmpInsertData.Add "NTID", ntid
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
    Dim ss As New SystemSettings
    Dim tblCols As New Collection
    Dim lastUserData As Scripting.Dictionary
    Dim tmpCol As String
    Dim tmpInsertCols As Collection
    Dim tmpInsertData As Scripting.Dictionary
    Dim checkDict As Scripting.Dictionary
    Dim dbm As New DbManager
    Dim i As Integer
    Dim v As Variant
    Dim lastNtid As String, ntid As String
    Dim str1 As String, str2 As String
    Dim Name As String
    Dim check As Boolean
    ss.Init
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
        For i = 0 To dbm.RecordSet.fields.count - 1
            tmpCol = dbm.RecordSet.fields(i).Name
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