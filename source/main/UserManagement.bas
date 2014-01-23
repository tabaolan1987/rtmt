Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database


Public Function CompareUserData()
    Dim ss As New SystemSettings
    Dim tblCols As New Collection
    Dim lastUserData As Scripting.Dictionary
    Dim tmpCol As String
    Dim tmpInsertCols As Collection
    Dim tmpInsertData As Scripting.Dictionary
    Dim dbm As New DbManager
    Dim i As Integer
    Dim v As Variant
    Dim lastNtid As String, ntid As String
    Dim str1 As String, str2 As String
    Dim name As String
    Dim check As Boolean
    ss.Init
    Dim query As String
    ' ===== DUPLICATION BLOCK =====
    
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
    Logger.LogDebug "UserManagement.CompareUserData", "Start check duplicate " & ss.NtidField & " records. Query: " & query
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
        For i = 0 To dbm.RecordSet.fields.count - 1
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
                Logger.LogDebug "UserManagement.CompareUserData", "lastNtid: " & lastNtid & ". Current Ntid: " & ntid
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
                            If Not check Then
                                Set tmpInsertData = New Scripting.Dictionary
                                tmpInsertData.Add "NTID", ntid
                                tmpInsertData.Add "Name", lastUserData.Item(Constants.FIELD_LAST_NAME) & " " & lastUserData.Item(dbm.GetFieldValue(dbm.RecordSet, Constants.FIELD_FIRST_NAME))
                                tmpInsertData.Add "Field heading", ss.SyncUsers.Item(CStr(v))
                                tmpInsertData.Add "Db field", CStr(v)
                                tmpInsertData.Add "Upload file", str1
                                tmpInsertData.Add "Select", "-1"
                                dbm.CreateLocalRecord tmpInsertData, tmpInsertCols, Constants.TABLE_USER_DATA_DUPLICATE
                            End If
                        End If
                    Next v
                    check = True
                Else
                    check = False
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
        Logger.LogInfo "UserManagement.CompareUserData", "There are no duplicate " & ss.NtidField & " records in table " & Constants.END_USER_DATA_CACHE_TABLE_NAME
    End If
    
    dbm.Recycle
    ' ===== DUPLICATION BLOCK =====
End Function