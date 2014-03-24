Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
' @author Hai Lu

Private dbs As DAO.Database
Private qdf As DAO.QueryDef
Private rst As DAO.RecordSet
Private rs As ADODB.RecordSet
Private cn As ADODB.Connection

Private mConnString As String
Private mTableName As String
Private mLocalTimestamp As String
Private mHeaders As Collection
Private mFieldTypes As Scripting.Dictionary
Private mIdCol As Collection
Private dbm As DbManager

Public Function init(tblName As String)
    Set dbs = CurrentDb
    mTableName = tblName
    Logger.LogDebug "SyncHelper.init", "Start sync table " & tblName
    Set dbm = New DbManager
    If Len(Session.Settings.Username) <> 0 Then
    
        mConnString = "DRIVER=SQL Server;SERVER=" & Session.Settings.ServerName & "," & Session.Settings.Port _
                                                & ";DATABASE=" & Session.Settings.DatabaseName _
                                                & ";UID=" & Session.Settings.Username _
                                                & ";PWD=" & Session.Settings.Password
    Else
        mConnString = "DRIVER=SQL Server;SERVER=" & Session.Settings.ServerName & "," & Session.Settings.Port _
                                & ";DATABASE=" & Session.Settings.DatabaseName
    End If
    
End Function

Public Function sync()
    If Ultilities.IfTableExists(mTableName) Then
        GetLocalTimestamp
        CompareLocal
        CompareServer
        PushLocalChange
        UpdateLocalTimestamp
    Else
        Dim cs As String
        cs = "ODBC;DRIVER=SQL Server;SERVER=" & Session.Settings.ServerName & "," & Session.Settings.Port _
                                                 & ";DATABASE=" & Session.Settings.DatabaseName _
                                                & ";UID=" & Session.Settings.Username _
                                                & ";PWD=" & Session.Settings.Password
        DoCmd.TransferDatabase acImport, "ODBC Database", cs, acTable, mTableName, mTableName, False, True
    End If
End Function

Public Function Recycle()
    RecycleLocal
    RecycleServer
    If Not dbs Is Nothing Then
        dbs.TableDefs.Refresh
        dbs.Close
        Set dbs = Nothing
    End If
End Function

Private Function RecycleLocal()
    On Error GoTo OnError
    If Not qdf Is Nothing Then
        qdf.Close
        Set qdf = Nothing
    End If
    If Not rst Is Nothing Then
        rst.Close
        Set rst = Nothing
    End If
OnExit:
    Exit Function
OnError:
    Logger.LogError "SyncHelper.RecycleLocal", "Could close current object", Err
    Resume OnExit
End Function

Private Function RecycleServer()
    On Error GoTo OnError
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not cn Is Nothing Then
        cn.Close
        Set cn = Nothing
    End If
OnExit:
    Exit Function
OnError:
    Logger.LogError "SyncHelper.RecycleServer", "Could close current object", Err
    Resume OnExit
End Function

Private Function GetLocalTimestamp()
    On Error GoTo OnError
    Set qdf = dbs.CreateQueryDef("", "select top 1 [timestamp] from [" & mTableName & "] order by [timestamp] desc")
    Set rst = qdf.OpenRecordSet
    If Not (rst.EOF And rst.BOF) Then
        rst.MoveFirst
        mLocalTimestamp = dbm.GetFieldValue(rst, "timestamp")
    End If
OnExit:
    RecycleLocal
    Exit Function
OnError:
    Logger.LogError "SyncHelper.GetLocalTimestamp", "Could not get local timestamp table " & mTableName, Err
    Resume OnExit
End Function

Private Function CompareServer()
    Set cn = New ADODB.Connection
    Dim tmpId As String
    Dim tmpData As Scripting.Dictionary
    Set mFieldTypes = New Scripting.Dictionary
    Dim tmpValue As String
    Dim query As String
    Dim i As Integer
    Dim v As Variant
    cn.Open mConnString
    cn.BeginTrans

    
    If Len(mLocalTimestamp) > 0 Then
        query = "select * from [" & mTableName _
                & "] where CONVERT(DATETIME, CONVERT(VARCHAR(MAX), [timestamp], 120), 120) > '" & StringHelper.EscapeQueryString(mLocalTimestamp) _
                & "'"
    Else
        query = "select * from [" & StringHelper.EscapeQueryString(mTableName) _
            & "] where [deleted]=0"
    End If
    Logger.LogDebug "SyncHelper.CompareServer", "Query: " & query
    Set rs = cn.Execute(query)
    Set mHeaders = New Collection
    For i = 0 To rs.fields.Count - 1
        mHeaders.Add rs.fields(i).Name
        mFieldTypes.Add rs.fields(i).Name, rs.fields(i).Type
    Next i
    Logger.LogDebug "SyncHelper.CompareServer", "Found header:"
    For Each v In mHeaders
        Logger.LogDebug "SyncHelper.CompareServer", "-> " & CStr(v)
    Next v
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        Do Until rs.EOF = True
            Set tmpData = New Scripting.Dictionary
            Logger.LogDebug "SyncHelper.CompareServer", "==========================="
            For Each v In mHeaders
                If IsNull(rs(CStr(v))) Then
                    tmpValue = ""
                Else
                    tmpValue = rs(CStr(v))
                End If
                tmpData.Add CStr(v), tmpValue
            Next v
            tmpId = rs("id")
            Logger.LogDebug "SyncHelper.CompareServer", "Found id: " & tmpId
            query = "select * from [" & mTableName _
                & "] where [id] = '" & StringHelper.EscapeQueryString(tmpId) _
                & "'"
            Logger.LogDebug "SyncHelper.CompareServer", "Check table " & mTableName & " id " & tmpId & ". Query: " & query
            Set qdf = dbs.CreateQueryDef("", query)
            Set rst = qdf.OpenRecordSet
            If Not (rst.EOF And rst.BOF) Then
                query = dbm.UpdateRecordQuery(tmpData, mHeaders, mTableName, mFieldTypes, False)
            Else
                query = dbm.CreateRecordQuery(tmpData, mHeaders, mTableName, mFieldTypes, False)
                
            End If
            Logger.LogDebug "SyncHelper.CompareServer", "Post execute query: " & query
            Set qdf = dbs.CreateQueryDef("", query)
            qdf.Execute
            RemoveChangeLog (tmpId)
            qdf.Close
            Set qdf = Nothing
            rst.Close
            Set rst = Nothing
            
            rs.MoveNext
        Loop
    End If
    cn.CommitTrans
OnExit:
    RecycleLocal
    RecycleServer
    Exit Function
OnError:
    cn.RollbackTrans
    Logger.LogError "SyncHelper.CompareServer", "Could not compare " & mTableName, Err
    Resume OnExit
End Function

Private Function CompareLocal()
    On Error GoTo OnError
    Dim v As Variant
    Set mIdCol = New Collection
    Set qdf = dbs.CreateQueryDef("", "select [TableId] from [ChangeLog] where [TableName]='" _
                        & StringHelper.EscapeQueryString(mTableName) & "'")
    Set rst = qdf.OpenRecordSet
    If Not (rst.EOF And rst.BOF) Then
        rst.MoveFirst
        Do Until rst.EOF = True
            mIdCol.Add dbm.GetFieldValue(rst, "TableId")
            rst.MoveNext
        Loop
    End If
    
    For Each v In mIdCol
        Logger.LogDebug "SyncHelper.CompareLocal", "Found " & mTableName & " | id: " & CStr(v)
    Next v
OnExit:
    RecycleLocal
    Exit Function
OnError:
    Logger.LogError "SyncHelper.CompareLocal", "Could not get changelog of table " & mTableName, Err
    Resume OnExit
End Function

Private Function PushLocalChange()
    If mIdCol.Count = 0 Then
        Exit Function
    End If
    Set cn = New ADODB.Connection
    Dim qBatch As Collection
    Dim mFilter As String
    Dim tmpData As Scripting.Dictionary
    Dim adData As Scripting.Dictionary
    Dim adCol As New Collection
    Dim tmpId As String
    Dim tmpTs As String
    Dim query As String
    Dim i As Integer
    Dim v As Variant
    
    adCol.Add "id"
    adCol.Add "ntid"
    adCol.Add "idFunction"
    adCol.Add "userAction"
    adCol.Add "description"
    adCol.Add "data_fields"
    adCol.Add "prev_value"
    adCol.Add "new_value"
    adCol.Add "table_name"
    
    mFilter = GetChangeLogFilter
    query = "select * from [" & mTableName & "] where [id] in (" & mFilter & ")"
    Logger.LogDebug "SyncHelper.PushLocalChange", "List changed data query: " & query
    cn.Open mConnString
    cn.BeginTrans
    
    Set rs = cn.Execute(query)
    Dim v1 As String
    Dim v2 As String
    Set qBatch = New Collection
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        Logger.LogDebug "SyncHelper.PushLocalChange", "Found record in server"
        Do Until rs.EOF = True
            
            tmpId = rs("id")
            tmpTs = rs("timestamp")
            Set qdf = dbs.CreateQueryDef("", "select * from [" & mTableName & "] where [id]='" & StringHelper.EscapeQueryString(tmpId) & "'")
            Set rst = qdf.OpenRecordSet
            If Not (rst.EOF And rst.BOF) Then
                rst.MoveFirst
                For Each v In mHeaders
                    If Not StringHelper.IsEqual(CStr(v), "id", True) _
                         And Not StringHelper.IsEqual(CStr(v), "timestamp", True) Then
                        
                        If IsNull(rs(CStr(v))) Then
                            v1 = ""
                        Else
                            v1 = rs(CStr(v))
                        End If
                        v2 = dbm.GetFieldValue(rst, CStr(v))
                        Logger.LogDebug "SyncHelper.PushLocalChange", "Compare column " & CStr(v) & ". Local: " & v2 & ". Server: " & v1
                        If Not StringHelper.IsEqual(v1, v2, False) Then
                            Set adData = New Scripting.Dictionary
                            adData.Add "id", StringHelper.GetGUID
                            adData.Add "ntid", Session.CurrentUser.ntid
                            adData.Add "idFunction", ""
                            adData.Add "userAction", "Update central store record"
                            adData.Add "description", "update [" & mTableName & "] set [" & CStr(v) _
                                    & "]='" & StringHelper.EscapeQueryString(v2) & "', [timestamp]=getdate() where [id]='" _
                                    & StringHelper.EscapeQueryString(tmpId) & "'"
                            adData.Add "data_fields", CStr(v)
                            adData.Add "prev_value", v1
                            adData.Add "new_value", v2
                            adData.Add "table_name", mTableName
                            query = dbm.CreateRecordQuery(adData, adCol, "audit_logs", IsServer:=True)
                            qBatch.Add query
                            
                        End If
                    End If
                Next v
            End If
            rst.Close
            Set rst = Nothing
            rs.MoveNext
        Loop
    End If
    rs.Close
    
    Set rs = Nothing
    If qBatch.Count > 0 Then
        For Each v In qBatch
            cn.Execute CStr(v)
        Next v
    End If
    Set qBatch = New Collection
    query = "select * from [" & mTableName & "] where [id] in (" & mFilter & ")"
    Set qdf = dbs.CreateQueryDef("", query)
    Set rst = qdf.OpenRecordSet
    Logger.LogDebug "SyncHelper.PushLocalChange", "Check for new record"
    Dim tmpDeleted As String
    If Not (rst.EOF And rst.BOF) Then
        rst.MoveFirst
        Do Until rst.EOF = True
            
            tmpTs = Trim(dbm.GetFieldValue(rst, "timestamp"))
            tmpId = dbm.GetFieldValue(rst, "id")
            tmpDeleted = dbm.GetFieldValue(rst, "deleted")
            Logger.LogDebug "SyncHelper.PushLocalChange", "deleted status: " & tmpDeleted
            'Logger.LogDebug "SyncHelper.PushLocalChange", "Found id: " & tmpId & ". Timestamp: " & tmpTs
            If Len(tmpTs) = 0 And StringHelper.IsEqual(tmpDeleted, "false", True) Then
                Set tmpData = New Scripting.Dictionary
                For Each v In mHeaders
                    tmpData.Add CStr(v), dbm.GetFieldValue(rst, CStr(v))
                Next v
                
                query = dbm.CreateRecordQuery(tmpData, mHeaders, mTableName, mFieldTypes, True)
                Set adData = New Scripting.Dictionary
                adData.Add "id", StringHelper.GetGUID
                adData.Add "ntid", Session.CurrentUser.ntid
                adData.Add "idFunction", ""
                adData.Add "userAction", "Create central store record"
                adData.Add "description", query
                adData.Add "data_fields", ""
                adData.Add "prev_value", ""
                adData.Add "new_value", ""
                adData.Add "table_name", mTableName
                query = dbm.CreateRecordQuery(adData, adCol, "audit_logs", IsServer:=True)
                'cn.Execute query
                qBatch.Add query
            End If
            rst.MoveNext
        Loop
    End If
    rst.Close
    Set rst = Nothing
    If qBatch.Count > 0 Then
        For Each v In qBatch
            cn.Execute CStr(v)
        Next v
    End If
    Set qBatch = Nothing
    cn.CommitTrans
OnExit:
    RecycleLocal
    RecycleServer
    Exit Function
OnError:
    cn.RollbackTrans
    Logger.LogError "SyncHelper.PushLocalChange", "Could not pust local change " & mTableName, Err
    Resume OnExit
End Function

Private Function UpdateLocalTimestamp()
    If mIdCol.Count = 0 Then
        Exit Function
    End If
    
    Set cn = New ADODB.Connection
    Dim mFilter As String
    Dim tmpData As Scripting.Dictionary
    Dim tmpId As String
    Dim tmpTs As String
    Dim query As String
    Dim i As Integer
    Dim v As Variant
    
    mFilter = GetChangeLogFilter
    query = "select * from [" & mTableName & "] where [id] in (" & mFilter & ")"
    Logger.LogDebug "SyncHelper.UpdateLocalTimestamp", "Query: " & query
    cn.Open mConnString
    cn.BeginTrans
    
    Set rs = cn.Execute(query)
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        Do Until rs.EOF = True
            tmpId = rs("id")
            tmpTs = rs("timestamp")
            Logger.LogDebug "SyncHelper.UpdateLocalTimestamp", "Update new local timestamp: " & tmpTs & ". Id: " & tmpId
            Set qdf = dbs.CreateQueryDef("", "delete from [ChangeLog] where [TableName]='" & StringHelper.EscapeQueryString(mTableName) _
                                            & "' and [TableId]='" & StringHelper.EscapeQueryString(tmpId) & "'")
            qdf.Execute
            Set qdf = dbs.CreateQueryDef("", "update [" & mTableName & "] set [timestamp]='" _
                                & StringHelper.EscapeQueryString(tmpTs) & "' where [id]='" & StringHelper.EscapeQueryString(tmpId) & "'")
            qdf.Execute
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    cn.CommitTrans
OnExit:
    RecycleLocal
    RecycleServer
    Exit Function
OnError:
    cn.RollbackTrans
    Logger.LogError "SyncHelper.UpdateLocalTimestamp", "Could not update local timestamp " & mTableName, Err
    Resume OnExit
End Function

Private Function RemoveChangeLog(id As String)
    Dim v As Variant
    Dim index As Integer
    Dim i As Integer
    i = 0
    index = -1
    If mIdCol.Count > 0 Then
        For Each v In mIdCol
            If StringHelper.IsEqual(CStr(v), id, True) Then
                index = i
                Exit For
            End If
            i = i + 1
        Next v
        If index <> -1 Then
            Logger.LogDebug "SyncHelper.CheckConflict", "Remove conflict id: " & id
            mIdCol.Remove (index)
        End If
    End If
End Function


Private Function GetChangeLogFilter() As String
    Dim v As Variant
    Dim mFilter As String
    mFilter = ""
    If mIdCol.Count > 0 Then
        For Each v In mIdCol
            mFilter = mFilter & "'" & StringHelper.EscapeQueryString(CStr(v)) & "',"
        Next v
        mFilter = Trim(mFilter)
        If StringHelper.EndsWith(mFilter, ",", True) Then
            mFilter = Left(mFilter, Len(mFilter) - 1)
        End If
    End If
    GetChangeLogFilter = mFilter
End Function