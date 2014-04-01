Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' @author Hai Lu

Private Const BULK_SIZE = 50
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
Private mIdCol As Scripting.Dictionary
Private mIdTs As Scripting.Dictionary
Private dbm As DbManager
Private mUncompleteId As Scripting.Dictionary
Private mEnablePrimary As Boolean
Private mEnableRegion As Boolean
Public Function Init(tblName As String)
    If Session.EnablePrimarySync.Exists(LCase(tblName)) Then
        mEnablePrimary = True
    Else
        mEnablePrimary = False
    End If
    
    If Session.SyncByRegion.Exists(LCase(tblName)) Then
        mEnableRegion = True
    Else
        mEnableRegion = False
    End If
    
    Set mUncompleteId = New Scripting.Dictionary
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
    Set mIdTs = New Scripting.Dictionary
End Function

Public Function sync()
    If Ultilities.IfTableExists(mTableName) Then
        GetLocalTimestamp
        CompareServer
        CompareLocal
        RollbackId
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
    Dim query As String
    
    query = "select top 1 [timestamp] from [" & mTableName & "]"
    If mEnablePrimary Then
        
    End If
    query = query & " order by [timestamp] desc"
    
    Set qdf = dbs.CreateQueryDef("", query)
    Set rst = qdf.OpenRecordSet
    If Not (rst.EOF And rst.BOF) Then
        rst.MoveFirst
        mLocalTimestamp = dbm.GetFieldValue(rst, "timestamp")
        'Logger.LogDebug "SyncHelper.GetLocalTimestamp", "Date format: " & Session.Settings.DateFormat
        If Len(mLocalTimestamp) > 0 Then
            mLocalTimestamp = Format(CDate(mLocalTimestamp), Session.Settings.DateFormat)
        End If
        Logger.LogDebug "SyncHelper.GetLocalTimestamp", "mLocalTimestamp: " & mLocalTimestamp
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
    Dim extraId As String
    Dim tmpLocalId As String
    Dim tmpRegion As String
    Dim i As Integer
    Dim v As Variant
    cn.Open mConnString
    'cn.BeginTrans

    
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
    Dim tmpFName As String
    For i = 0 To rs.fields.count - 1
        tmpFName = rs.fields(i).Name
        mHeaders.Add LCase(tmpFName)
        mFieldTypes.Add LCase(tmpFName), rs.fields(i).Type
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
                Logger.LogDebug "SyncHelper.CompareServer", "get header value: " & CStr(v)
                If IsNull(rs(CStr(v))) Then
                    tmpValue = ""
                Else
                    tmpValue = Trim(rs(CStr(v)))
                End If
                If Not tmpData.Exists(CStr(v)) Then
                    tmpData.Add LCase(CStr(v)), tmpValue
                    If mEnablePrimary Then
                        If StringHelper.IsEqual(CStr(v), Session.EnablePrimarySync.Item(LCase(mTableName)), True) Then
                            extraId = tmpValue
                        End If
                    End If
                    If mEnableRegion Then
                        If StringHelper.IsEqual(CStr(v), Session.SyncByRegion.Item(LCase(mTableName)), True) Then
                            tmpRegion = tmpValue
                        End If
                    End If
                End If
            Next v
            tmpId = rs("id")
            Logger.LogDebug "SyncHelper.CompareServer", "Found id: " & tmpId
            query = "select * from [" & mTableName _
                & "] where "
            If mEnablePrimary Then
                query = query & " [" & Session.EnablePrimarySync.Item(LCase(mTableName)) & "] = '" & StringHelper.EscapeQueryString(extraId) & "'"
                If mEnableRegion Then
                    query = query & " and [" & Session.SyncByRegion.Item(LCase(mTableName)) & "] = '" & StringHelper.EscapeQueryString(tmpRegion) & "'"
                End If
            Else
                query = query & " [id] = '" & StringHelper.EscapeQueryString(tmpId) & "'"
            End If
            Logger.LogDebug "SyncHelper.CompareServer", "Check table " & mTableName & " id " & tmpId & ". Query: " & query
            Set qdf = dbs.CreateQueryDef("", query)
            Set rst = qdf.OpenRecordSet
            If Not (rst.EOF And rst.BOF) Then
                If mEnablePrimary Then
                    tmpLocalId = dbm.GetFieldValue(rst, "id")
                    tmpData.Remove LCase("id")
                    tmpData.Add LCase("id"), tmpLocalId
                End If
                query = dbm.UpdateRecordQuery(tmpData, mHeaders, mTableName, mFieldTypes, False)
            Else
                query = dbm.CreateRecordQuery(tmpData, mHeaders, mTableName, mFieldTypes, False)
            End If
            Logger.LogDebug "SyncHelper.CompareServer", "Post execute query: " & query
            qdf.Close
            Set qdf = Nothing
            rst.Close
            Set rst = Nothing
            
            Set qdf = dbs.CreateQueryDef("", query)
            qdf.Execute
            qdf.Close
            Set qdf = Nothing
            Set qdf = dbs.CreateQueryDef("", "delete from [ChangeLog] where [TableName]='" & StringHelper.EscapeQueryString(mTableName) _
                                            & "' and [TableId]='" & StringHelper.EscapeQueryString(tmpId) & "'")
            qdf.Execute
            qdf.Close
            Set qdf = Nothing
            rs.MoveNext
        Loop
    End If
   ' cn.CommitTrans
OnExit:
    RecycleLocal
    RecycleServer
    Exit Function
OnError:
   ' cn.RollbackTrans
    Logger.LogError "SyncHelper.CompareServer", "Could not compare " & mTableName, Err
    Resume OnExit
End Function

Private Function RollbackId()
    On Error GoTo OnError
    Dim v As Variant
    Set qdf = dbs.CreateQueryDef("", "delete from [ChangeLog] where [TableName]='" _
                        & StringHelper.EscapeQueryString(mTableName) & "'")
    qdf.Execute
    If mUncompleteId.count > 0 Then
        For Each v In mUncompleteId
            Set qdf = dbs.CreateQueryDef("", "insert into [ChangeLog]([TableName], [TableId]) values('" _
                            & StringHelper.EscapeQueryString(mTableName) & "', '" & StringHelper.EscapeQueryString(CStr(v)) & "')")
            qdf.Execute
            Set qdf = Nothing
        Next v
    End If
OnExit:
    RecycleLocal
    Exit Function
OnError:
    Logger.LogError "SyncHelper.RollbackId", "Could not rollback " & mTableName, Err
    Resume OnExit
End Function

Private Function CompareLocal()
  '  On Error GoTo OnError
    Dim mQdf As DAO.QueryDef
    Dim mRst As DAO.RecordSet
    Dim v As Variant
    Dim tmpId As String
    Logger.LogDebug "SyncHelper.CompareLocal", "Start compare local"
    
    Set mQdf = dbs.CreateQueryDef("", "select [TableId] from [ChangeLog] where [TableName]='" _
                        & StringHelper.EscapeQueryString(mTableName) & "' group by [TableId]")
    Set mRst = mQdf.OpenRecordSet
    Dim count As Integer
    count = 0
    Set mIdCol = New Scripting.Dictionary
    Set mIdTs = New Scripting.Dictionary
    If Not (mRst.EOF And mRst.BOF) Then
        mRst.MoveFirst
        Do Until mRst.EOF = True
            tmpId = dbm.GetFieldValue(mRst, "TableId")
            If Not mIdCol.Exists(tmpId) Then
                mIdCol.Add tmpId, tmpId
            End If
            If count >= BULK_SIZE Then
                Logger.LogDebug "SyncHelper.CompareLocal", "Push bulk"
               ' CompareLocal
                PushLocalChange
                UpdateLocalTimestamp
                'RollbackId
                If mIdCol.count > 0 Then
                    For Each v In mIdCol
                        If Not mUncompleteId.Exists(CStr(v)) Then
                            mUncompleteId.Add CStr(v), CStr(v)
                        End If
                    Next v
                End If
                Set mIdTs = New Scripting.Dictionary
                Set mIdCol = New Scripting.Dictionary
                count = 0
            End If
            count = count + 1
            mRst.MoveNext
        Loop
        Logger.LogDebug "SyncHelper.CompareLocal", "Push bulk"
        'CompareLocal
        PushLocalChange
        UpdateLocalTimestamp
        If mIdCol.count > 0 Then
                    For Each v In mIdCol
                        If Not mUncompleteId.Exists(CStr(v)) Then
                            mUncompleteId.Add CStr(v), CStr(v)
                        End If
                    Next v
        End If
    End If
    
    For Each v In mIdCol
        Logger.LogDebug "SyncHelper.CompareLocal", "Found " & mTableName & " | id: " & CStr(v)
    Next v
    
OnExit:
    On Error Resume Next
    mQdf.Close
    Set mQdf = Nothing
    mRst.Close
    Set mRst = Nothing
    RecycleLocal
    Exit Function
OnError:
    Logger.LogError "SyncHelper.CompareLocal", "Could not get changelog of table " & mTableName, Err
    Resume OnExit
End Function

Private Function PushLocalChange()
    If mIdCol.count = 0 Then
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
                            v1 = Trim(rs(CStr(v)))
                        End If
                        v2 = Trim(dbm.GetFieldValue(rst, CStr(v)))
                        Logger.LogDebug "SyncHelper.PushLocalChange", "Compare column " & CStr(v) & ". Local: " & v2 & ". Server: " & v1
                        If Not StringHelper.IsEqual(v1, v2, True) Then
                            Set adData = New Scripting.Dictionary
                            adData.Add "id", StringHelper.GetGUID
                            adData.Add "ntid", Session.CurrentUser.ntid
                            adData.Add "idFunction", Session.CurrentUser.FuncRegion.Region
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
                    If Not mIdTs.Exists(tmpId) Then
                        mIdTs.Add tmpId, tmpId
                    End If
                    RemoveChangeLog tmpId
                Next v
            End If
            qdf.Close
            Set qdf = Nothing
            rst.Close
            Set rst = Nothing
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
  '  If qBatch.count > 0 Then
  '      For Each v In qBatch
  '          cn.Execute CStr(v)
  '      Next v
  '  End If
  '  Set qBatch = New Collection
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
            'Logger.LogDebug "SyncHelper.PushLocalChange", "deleted status: " & tmpDeleted
            'Logger.LogDebug "SyncHelper.PushLocalChange", "Found id: " & tmpId & ". Timestamp: " & tmpTs
            If Len(tmpTs) = 0 And StringHelper.IsEqual(tmpDeleted, "false", True) Then
                Set tmpData = New Scripting.Dictionary
                For Each v In mHeaders
                    If Not tmpData.Exists(CStr(v)) Then
                        tmpData.Add CStr(v), dbm.GetFieldValue(rst, CStr(v))
                    End If
                Next v
                Logger.LogDebug "SyncHelper.PushLocalChange", "Create query for new record"
                query = dbm.CreateRecordQuery(tmpData, mHeaders, mTableName, mFieldTypes, True)
                Set adData = New Scripting.Dictionary
                adData.Add "id", StringHelper.GetGUID
                adData.Add "ntid", Session.CurrentUser.ntid
                adData.Add "idFunction", Session.CurrentUser.FuncRegion.Region
                adData.Add "userAction", "Create central store record"
                adData.Add "description", query
                adData.Add "data_fields", ""
                adData.Add "prev_value", ""
                adData.Add "new_value", ""
                adData.Add "table_name", mTableName
                Logger.LogDebug "SyncHelper.PushLocalChange", "Create query for audit log"
                query = dbm.CreateRecordQuery(adData, adCol, "audit_logs", IsServer:=True)
                'cn.Execute query
                Logger.LogDebug "SyncHelper.PushLocalChange", "Add to collection"
                qBatch.Add query
            End If
            If Not mIdTs.Exists(tmpId) Then
                mIdTs.Add tmpId, tmpId
            End If
            RemoveChangeLog tmpId
            rst.MoveNext
        Loop
    End If
    qdf.Close
    Set qdf = Nothing
    rst.Close
    Set rst = Nothing
    cn.BeginTrans
    If qBatch.count > 0 Then
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
    If mIdTs.count = 0 Then
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
    
    mFilter = GetChangeLogFilter(mIdTs)
    query = "select * from [" & mTableName & "] where [id] in (" & mFilter & ")"
    Logger.LogDebug "SyncHelper.UpdateLocalTimestamp", "Query: " & query
    cn.Open mConnString
    'cn.BeginTrans
    
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
            qdf.Close
            Set qdf = dbs.CreateQueryDef("", "update [" & mTableName & "] set [timestamp]='" _
                                & StringHelper.EscapeQueryString(tmpTs) & "' where [id]='" & StringHelper.EscapeQueryString(tmpId) & "'")
            qdf.Execute
            qdf.Close
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    'cn.CommitTrans
OnExit:
    RecycleLocal
    RecycleServer
    Exit Function
OnError:
    'cn.RollbackTrans
    Logger.LogError "SyncHelper.UpdateLocalTimestamp", "Could not update local timestamp " & mTableName, Err
    Resume OnExit
End Function

Private Function RemoveChangeLog(id As String)
    On Error Resume Next
    Dim v As Variant
    If mIdCol.count > 0 Then
        Logger.LogDebug "SyncHelper.CheckConflict", "Remove conflict id: " & id
        If mIdCol.Exists(id) Then
            mIdCol.Remove (id)
        End If
    End If
End Function


Private Function GetChangeLogFilter(Optional col As Scripting.Dictionary) As String
    Dim v As Variant
    Dim mFilter As String
    mFilter = ""
    If col Is Nothing Then
        Set col = mIdCol
    End If
    If col.count > 0 Then
        For Each v In col.keys
            mFilter = mFilter & "'" & StringHelper.EscapeQueryString(CStr(v)) & "',"
        Next v
        mFilter = Trim(mFilter)
        If StringHelper.EndsWith(mFilter, ",", True) Then
            mFilter = Left(mFilter, Len(mFilter) - 1)
        End If
    End If
    GetChangeLogFilter = mFilter
End Function