Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @author: Hai Lu
' Database management object
Option Compare Database
Private dbs As DAO.Database
Private qdf As DAO.QueryDef
Private rst As DAO.RecordSet
Private mWarnings() As String
Private countWarning As Integer

Public Property Get warnings() As String()
    warnings = mWarnings
End Property

Public Property Get RecordSet() As DAO.RecordSet
    Set RecordSet = rst
End Property

Public Property Get QueryDef() As DAO.QueryDef
    Set QueryDef = qdf
End Property

Public Function Init()
    If dbs Is Nothing Then
        Set dbs = CurrentDb
    End If
End Function

Public Function Recycle()
    On Error GoTo OnError
    If Not qdf Is Nothing Then
        qdf.Close
        Set qdf = Nothing
    End If
    If Not rst Is Nothing Then
        rst.Close
        Set rst = Nothing
    End If
    If Not dbs Is Nothing Then
        dbs.Close
        Set dbs = Nothing
    End If
OnExit:
    Exit Function
OnError:
    Logger.LogError "DbManager.Recycle", "Could close db object: ", Err
    Resume OnExit
End Function

Public Function ExecuteQuery(query As String, Optional params As Scripting.Dictionary)
    On Error GoTo OnError
    Dim key As String, value As Variant
    Set qdf = dbs.CreateQueryDef("", query)
    If Not params Is Nothing Then
        'Logger.LogDebug "DbManager.OpenRecordSet", "Param cound: " & params.count
        For i = 0 To params.count - 1
            On Error Resume Next
            key = params.keys(i)
            value = params.Items(i)
            'Logger.LogDebug "DbManager.OpenRecordSet", "Param key: " & key & ". Value: " & value
            qdf.Parameters(key).value = value
            On Error GoTo 0
        Next i
    End If
    qdf.Execute
    dbs.TableDefs.Refresh
OnExit:
    Exit Function
OnError:
    Logger.LogError "DbManager.ExecuteQuery", "Could execute query: " & query, Err
    Resume OnExit
End Function

Public Function OpenRecordSet(query As String, Optional params As Scripting.Dictionary)
    Dim prm As DAO.Parameter, i As Integer, key As String
    On Error GoTo OnError
    Set qdf = dbs.CreateQueryDef("", query)
    If Not params Is Nothing Then
        
        'Logger.LogDebug "DbManager.OpenRecordSet", "Param cound: " & params.count
        For i = 0 To params.count - 1
            On Error Resume Next
            key = params.keys(i)
            'Logger.LogDebug "DbManager.OpenRecordSet", "Param key: " & params.Keys(i) & ". Value: " & params.Items(i)
            If params.Exists(key) Then
                qdf.Parameters(key).value = params.Items(key)
            End If
            On Error GoTo 0
        Next i
    End If
    Set rst = qdf.OpenRecordSet
OnExit:
    Exit Function
OnError:
    Logger.LogError "DbManager.OpenRecordSet", "Could execute query: " & query, Err
    Resume OnExit
End Function

Public Function RecycleTable(s As SystemSettings)
    Dim i As Integer
    Dim TableNames() As String
    Dim tblName As Variant, tmp As String
    
    TableNames = s.TableNames
    Logger.LogDebug "DbManager.RecycleTable", "Start check table. Size: " & CStr(UBound(TableNames))
    Dim query As String
    
    For i = LBound(TableNames) To UBound(TableNames)
        tmp = Trim(CStr(TableNames(i)))
        Logger.LogDebug "DbManager.RecycleTable", "Check table name " & tmp
        If Ultilities.IfTableExists(tmp) = True Then
            'ExecuteQuery FileHelper.ReadQuery(tmp, Constants.Q_DELETE_ALL)
            DoCmd.DeleteObject acTable, tmp
        End If
        ExecuteQuery FileHelper.ReadQuery(tmp, Constants.Q_CREATE)
    Next
End Function

Private Function GetHeaderIndex(name As String) As Integer
    Dim index As Integer
    index = -1
    For i = 0 To rst.fields.count - 1
        If StringHelper.IsEqual(Trim(rst.fields(i).name), Trim(name), True) Then
            'Logger.LogDebug "DbManager.SyncUserData", "## HEADER: " & rst.fields(i).name
            index = i
        End If
    Next i
    GetHeaderIndex = index
End Function

Private Function GetFieldValue(rs As RecordSet, name As String) As String
    If Len(name) <> 0 Then
        Dim index As Integer
        index = -1
        For i = 0 To rs.fields.count - 1
            If StringHelper.IsEqual(Trim(rs.fields(i).name), Trim(name), True) Then
                index = i
            End If
        Next i
        If index <> -1 And Len(rs.fields(index).value) <> 0 Then
            GetFieldValue = Trim(rs.fields(index).value)
        Else
            GetFieldValue = ""
        End If
    Else
        GetFieldValue = ""
    End If
End Function

Private Function AddWarning(mes As String)
    ReDim Preserve mWarnings(countWarning)
    mWarnings(countWarning) = mes
    Logger.LogError "DbManager.AddWarning", mes, Nothing
    countWarning = countWarning + 1
End Function

Private Function ValidateNtid(s As SystemSettings, ntids As String, Optional userData As Scripting.Dictionary)
    Dim validatorMapping As Scripting.Dictionary
    Dim checkList As Collection
    Dim tmpDict As Scripting.Dictionary
    Dim fields As String
    Dim mData As String
    Dim result As String
    Dim ntid As String
    Dim str1 As String, str2 As String, key As String, value As String
    Dim check As Boolean
    Dim v As Variant
    Dim tmpUserData As Scripting.Dictionary
    Set validatorMapping = s.validatorMapping
    Dim i As Integer
    For i = 0 To validatorMapping.count - 1
        fields = fields & validatorMapping.Items(i) & ","
    Next i
    If StringHelper.EndsWith(fields, ",", True) Then
        fields = Left(fields, Len(fields) - 1)
    End If
    
    If StringHelper.EndsWith(ntids, ",", True) Then
        ntids = Left(ntids, Len(ntids) - 1)
    End If
    mData = "token=" & StringHelper.EncodeURL(s.Token) _
                & "&fields=" & StringHelper.EncodeURL(fields) _
                & "&ntids=" & StringHelper.EncodeURL(ntids)
    Logger.LogDebug "DbManager.SyncUserData", "Post valid ntids: " & ntids
    result = HttpHelper.PostData(s.ValidatorURL, mData)
    Logger.LogDebug "DbManager.SyncUserData", "Result: " & result
    
    If Len(result) > 0 Then
        If StringHelper.IsContain(result, "}", True) And StringHelper.IsContain(result, "{", True) Then
            Set checkList = JSONHelper.parse(result)
            For Each tmpDict In checkList
                ntid = tmpDict.Item("ntid")
                check = tmpDict.Item("isvalid")
                Logger.LogDebug "DbManager.SyncUserData", "Is valid: " & CStr(check)
                If check Then
                    Logger.LogDebug "DbManager.SyncUserData", "check ntid: " & ntid
                    If Not userData Is Nothing And userData.count > 0 Then
                        Set tmpUserData = userData.Item(ntid)
                        For Each v In validatorMapping
                            key = v
                            value = validatorMapping.Item(key)
                            str1 = tmpUserData.Item(key)
                            str2 = tmpDict.Item(value)
                            Logger.LogDebug "DbManager.SyncUserData", "User data key: " & key & " = " & str1
                            Logger.LogDebug "DbManager.SyncUserData", "Mapping data key: " & value & " = " & str2
                            If StringHelper.IsEqual(str1, str2, True) Then
                                Logger.LogDebug "DbManager.SyncUserData", "validated!!!"
                            Else
                                AddWarning ("Validation failed !!! NTID: " & ntid & " . Field name: " & key & ". Local: " & str1 & ". LDAP: " & str2)
                            End If
                        Next v
                    End If
                Else
                    AddWarning ("Validation failed !!! NTID: " & ntid & " not found!")
                End If
            Next tmpDict
        Else
            Logger.LogDebug "DbManager.SyncUserData", "Error: " & result
        End If
    Else
        
    End If
    
End Function

Public Function SyncUserData()
    Dim s As SystemSettings: Set s = New SystemSettings
    Dim flag As Boolean: flag = False
    Dim dictMapping As Scripting.Dictionary, i As Integer, j As Integer, k As Integer, key As String, value As String
    Dim tmpDict As Scripting.Dictionary
    Dim dictParams As Scripting.Dictionary
    Dim queryInsertData As String
    Dim queryCustomInsert As String, checkValue As String, tmpSplit() As String, tmpValues() As String
    Dim tmpValue As String
    Dim tmpCache As String
    ' Init database
    Init
    ' Init settings
    s.Init
    '
    RecycleTable s
    ' Read the dict mapping
    Set dictMapping = s.SyncUsers
    ' Read query insert user data
    'queryInsertData = FileHelper.ReadQuery(Constants.END_USER_DATA_TABLE_NAME, Constants.Q_INSERT)
    ' Open tmp table user data from CSV file
    OpenRecordSet "select * from " & Constants.TMP_END_USER_TABLE_NAME
    
    Logger.LogDebug "DbManager.SyncUserData", "ntidField: " & s.NtidField
    Logger.LogDebug "DbManager.SyncUserData", "validatorUrl: " & s.ValidatorURL
    Logger.LogDebug "DbManager.SyncUserData", "token: " & s.Token
    Logger.LogDebug "DbManager.SyncUserData", "bulkSize: " & CStr(s.BulkSize)
    Dim ntids As String, tmpNtid As String, size As Long, userData As Scripting.Dictionary
    Dim tmpUserData As Scripting.Dictionary
    Dim tmpCols As Collection
    size = 0
    If Not (rst.EOF And rst.BOF) Then
        rst.MoveFirst
        Do Until rst.EOF = True
            Dim tmpList() As String, arraySize As Integer
            Logger.LogDebug "DbManager.SyncUserData", "###########################################"
            Set tmpCols = New Collection
            ' List all mapping column
            Set dictParams = New Scripting.Dictionary
            Set tmpUserData = New Scripting.Dictionary
            For i = 0 To dictMapping.count - 1
                 key = dictMapping.keys(i)
                 value = dictMapping.Items(i)
                 tmpCols.Add key
                 If StringHelper.IsContain(key, "insert into", True) _
                    And StringHelper.IsContain(key, "|", True) Then
                    If flag = False And Len(value) <> 0 And Len(key) <> 0 Then
                        ' Add custom insert key to list
                        ReDim Preserve tmpList(arraySize)
                        Logger.LogDebug "DbManager.SyncUserData", "Why we import that?" & key
                        tmpList(arraySize) = key
                        arraySize = arraySize + 1
                    End If
                 Else
                    ' Add params value
                    On Error Resume Next
                    tmpCache = GetFieldValue(rst, value)
                    Logger.LogDebug "DbManager.SyncUserData", key & " = " & tmpCache
                    If Len(value) <> 0 Then
                        tmpUserData.Add value, tmpCache
                    End If
                    dictParams.Add key, tmpCache
                 End If
            Next i
            ' Add custom insert key one time
            flag = True
            ' Insert data to user_data table
            'ExecuteQuery queryInsertData, dictParams
            Logger.LogDebug "DbManager.SyncUserData", "Insert user record"
            tmpCols.Add Constants.FIELD_ID
            dictParams.Add Constants.FIELD_ID, StringHelper.GetGUID
            CreateLocalRecord dictParams, tmpCols, Constants.END_USER_DATA_TABLE_NAME
            
            Logger.LogDebug "DbManager.SyncUserData", "Number custom insert: " & CStr(UBound(tmpList) + 1)
            ' Insert mapping data
            For i = LBound(tmpList) To UBound(tmpList)
                If StringHelper.IsContain(key, "insert into", True) _
                    And StringHelper.IsContain(key, "|", True) Then
                    key = tmpList(i)
                    tmpSplit = Split(key, "|")
                    queryCustomInsert = Trim(tmpSplit(0))
                    checkValue = Trim(tmpSplit(1))
                    Logger.LogDebug "DbManager.SyncUserData", "checkValue: " & checkValue
                    Logger.LogDebug "DbManager.SyncUserData", "custom import query: " & queryCustomInsert & ". Check value: " & checkValue & " . From list: " & value
                    value = dictMapping.Items(key)
                    tmpValues = Split(value, ",")
                    Logger.LogDebug "DbManager.SyncUserData", "Number of column to check: " & CStr(UBound(tmpValues) + 1)
                    ' Loop to check all column
                    'Logger.LogDebug "DbManager.SyncUserData", "====== ITID: " & rst("ntid") & " ======"
                    For j = LBound(tmpValues) To UBound(tmpValues)
                        tmpValue = Trim(tmpValues(j))
                        Logger.LogDebug "DbManager.SyncUserData", "Column to compare: " & tmpValue
                        Dim index As Integer
                        index = GetHeaderIndex(tmpValue)
                        Logger.LogDebug "DbManager.SyncUserData", "Index: " & index
                        If Len(rst.fields(index).value) <> 0 Then
                            tmpCache = Trim(rst.fields(index).value)
                        Else
                            tmpCache = ""
                        End If
                        
                        Logger.LogDebug "DbManager.SyncUserData", "Value: " & tmpCache
                        'Logger.LogDebug "DbManager.SyncUserData", "column name: " & tmpValue
                        If StringHelper.IsEqual(tmpCache, checkValue, True) Then
                            ' If value is valid, get parameter and execute query
                            Set tmpDict = New Scripting.Dictionary
                            For k = 0 To dictParams.count - 1
                                tmpDict.Add dictParams.keys(k), dictParams.Items(k)
                            Next k
                            tmpDict.Add "value", tmpValue
                            tmpDict.Add "region_name", s.RegionName
                            ExecuteQuery queryCustomInsert, tmpDict
                        End If
                    Next j
                End If
            Next i
            
            ' === VALIDATION BLOCK ===
            If size >= s.BulkSize Then
                Logger.LogDebug "DbManager.SyncUserData", "check bulk user data size: " & userData.count
                ValidateNtid s, ntids, userData
                ntids = ""
                size = 0
                Set userData = Nothing
            End If
            If userData Is Nothing Then
                Set userData = New Scripting.Dictionary
            End If
            tmpNtid = rst(s.NtidField)
            userData.Add tmpNtid, tmpUserData
            Set tmpUserData = Nothing
            
            ntids = ntids & tmpNtid & ","
            size = size + 1
            ' === VALIDATION BLOCK ===
            
            rst.MoveNext
        Loop
        ' === VALIDATION BLOCK ===
        If Len(ntids) <> 0 And size <> 0 And Not userData Is Nothing And userData.count > 0 Then
            ValidateNtid s, ntids, userData
            ntids = ""
            size = 0
        End If
        ' === VALIDATION BLOCK ===
    Else
        Logger.LogInfo "DbManager.SyncUserData", "There are no records in table " & Constants.TMP_END_USER_TABLE_NAME
    End If
End Function

Public Function ImportData(csvPath As String)
    On Error GoTo OnError
    If Ultilities.IfTableExists(Constants.TMP_END_USER_TABLE_NAME) Then
        dbs.TableDefs.Delete Constants.TMP_END_USER_TABLE_NAME
    End If
    dbs.TableDefs.Refresh
    DoCmd.TransferText TransferType:=acLinkDelim, TableName:=Constants.TMP_END_USER_TABLE_NAME, _
        FileName:=csvPath, HasFieldNames:=True
    dbs.TableDefs.Refresh
OnExit:
    On Error Resume Next
        ' Delete table if name is correct or not
        DoCmd.DeleteObject acTable, "Name AutoCorrect Save Failures"
    Exit Function
OnError:
    If Ultilities.IfTableExists(Constants.TMP_END_USER_TABLE_NAME) Then
        DoCmd.DeleteObject acTable, Constants.TMP_END_USER_TABLE_NAME
    End If
    Logger.LogError "DbManager.ImportData", "Could not import user data from CSV file " & csvPath, Err
    Resume OnExit
End Function

Public Function ImportSqlTable(Server As String, _
                                    DatabaseName As String, _
                                    fromTable As String, _
                                    desTable As String, _
                                    Optional Username As String, _
                                    Optional Password As String)
    Dim check As Boolean: check = False
    Logger.LogDebug "DbManager.ImportSqlTable", "Server: " & Server _
                                                & ", Database: " & DatabaseName _
                                                & ", FromTable: " & fromTable _
                                                & ", ToTable: " & desTable _
                                                & ", Username: " & Username
    On Error GoTo OnError
    If Ultilities.IfTableExists(desTable) Then
        Logger.LogDebug "DbManager.ImportSqlTable", "Create cached table " & desTable & "_tmp"
        DoCmd.Rename desTable & "_tmp", acTable, desTable
    End If
    Dim stConnect As String
    If Len(Username) <> 0 Then
        stConnect = "ODBC;DRIVER=SQL Server;SERVER=" & Server & ";DATABASE=" & DatabaseName _
                                                & ";UID=" & Username _
                                                & ";PWD=" & Password
    Else
        stConnect = "ODBC;DRIVER=SQL Server;SERVER=" & Server & ";DATABASE=" & DatabaseName
    End If
    Logger.LogDebug "DbManager.ImportSqlTable", "Start import table " & desTable & " from table " & fromTable
    DoCmd.TransferDatabase acImport, "ODBC Database", stConnect, acTable, desTable, fromTable, False, True
    check = True
    Logger.LogDebug "DbManager.ImportSqlTable", "Check: " & check
    
OnExit:
    On Error GoTo Quit
    If check = True Then
        If Ultilities.IfTableExists(desTable & "_tmp") Then
            Logger.LogDebug "DbManager.ImportSqlTable", "Delete cached table " & desTable & "_tmp"
            DoCmd.DeleteObject acTable, desTable & "_tmp"
        End If
    Else
        If Ultilities.IfTableExists(desTable & "_tmp") Then
            Logger.LogDebug "DbManager.ImportSqlTable", "Rename cached table " & desTable & "_tmp" & " to " & desTable
            DoCmd.Rename desTable, acTable, desTable & "_tmp"
        End If
    End If
Quit:
    Exit Function
OnError:
    Logger.LogError "DbManager.ImportSqlTable", "Could not import table " _
                        & desTable & " data from table " & fromTable, Err
    Resume OnExit
End Function

Public Function RecycleTableName(name As String)
    Init
        Logger.LogDebug "DbManager.SyncTable", "Recycle table name " & name
        dbs.TableDefs.Refresh
        If Ultilities.IfTableExists(name) Then
            Logger.LogDebug "DbManager.SyncTable", "Delete all record table " & name
            ExecuteQuery FileHelper.ReadQuery(name, Constants.Q_DELETE_ALL)
            'DoCmd.DeleteObject acTable, name
        Else
            Logger.LogDebug "DbManager.SyncTable", "Create new table " & name
            ExecuteQuery FileHelper.ReadQuery(name, Constants.Q_CREATE)
        End If
        
        dbs.TableDefs.Refresh
    Recycle
End Function

Public Function SyncTable(Server As String, _
                                    DatabaseName As String, _
                                    fromTable As String, _
                                    desTable As String, _
                                    Optional Username As String, _
                                    Optional Password As String, _
                                    Optional CheckConflict As Boolean)
    Logger.LogDebug "DbManager.SyncTable", "Start sync table " & fromTable
    If Ultilities.IfTableExists(desTable) Then
        Init
        Logger.LogDebug "DbManager.SyncTable", "Table " & fromTable & " is existed!"
        Dim tblCached As String
        Dim tmpTimestampServer As String
        Dim tmpTimestampLocal As String
        Dim tmpRst As RecordSet
        Dim tmpId As String
        Dim tmpQuery As String
        Dim i As Integer
        Dim tmpCol As String, tmpType As String
        Dim tmpColType As Scripting.Dictionary
        Dim tmpDataServer As Scripting.Dictionary
        Dim tmpDataLocal As Scripting.Dictionary
        Dim tmpCols As Collection
        Dim str1 As String, str2 As String
        Dim c As Integer
        tblCached = desTable & "_cached"
        If Ultilities.IfTableExists(tblCached) Then
            DoCmd.DeleteObject acTable, tblCached
        End If
        DoCmd.Rename tblCached, acTable, desTable
        
        ImportSqlTable Server, DatabaseName, fromTable, desTable, Username, Password
        
        OpenRecordSet "SELECT * FROM " & desTable
        If Not (rst.EOF And rst.BOF) Then
            rst.MoveFirst
            Do Until rst.EOF = True
                Set tmpCols = New Collection
                tmpTimestampServer = GetFieldValue(rst, Constants.FIELD_TIMESTAMP)
                tmpId = GetFieldValue(rst, Constants.FIELD_ID)
                If Len(tmpId) <> 0 And Len(tmpTimestampServer) <> 0 Then
                    'Logger.LogDebug "DbManager.SyncTable", "ID: " & tmpId & " | Timestamp: " & tmpTimestamp
                    Set qdf = dbs.CreateQueryDef("", "SELECT * FROM " & tblCached & " WHERE ID=[VALUE]")
                    qdf.Parameters("VALUE").value = tmpId
                    Set tmpRst = qdf.OpenRecordSet
                    If Not (tmpRst.EOF And tmpRst.BOF) Then
                        Dim check As Boolean: check = False
                        tmpTimestampLocal = GetFieldValue(tmpRst, Constants.FIELD_TIMESTAMP)
                        c = TimerHelper.Compare(tmpTimestampLocal, tmpTimestampServer)
                        tmpRst.MoveFirst
                        For i = 0 To tmpRst.fields.count - 1
                           ' Logger.LogDebug "field type:", tmpRst.fields(i).Type & " === " & dbBoolean
                            tmpCol = tmpRst.fields(i).name
                            tmpType = tmpRst.fields(i).Type
                            str1 = Trim(GetFieldValue(tmpRst, tmpCol))
                            str2 = Trim(GetFieldValue(rst, tmpCol))
                            
                            If Not StringHelper.IsEqual(str1, str2, False) Then
                                Logger.LogDebug "DbManager.SyncTable", tmpCol & " | " & StringHelper.IsEqual(tmpCol, Constants.FIELD_ID, True) & " | " & StringHelper.IsEqual(tmpCol, Constants.FIELD_TIMESTAMP, True)
                                Logger.LogDebug "DbManager.SyncTable", tmpId & " | " & str1 & " | " & str2
                                
                                If (tmpRst.fields(i).Type = dbBoolean) Then
                                    If StringHelper.IsEqual(str1, "True", True) Then
                                        str1 = "-1"
                                    Else
                                        str1 = "0"
                                    End If
                                    If StringHelper.IsEqual(str2, "True", True) Then
                                        str2 = "-1"
                                    Else
                                        str2 = "0"
                                    End If
                                End If
                                Set tmpDataLocal = New Scripting.Dictionary
                                Set tmpDataServer = New Scripting.Dictionary
                                Set tmpCols = New Collection
                                Logger.LogDebug "DbManager.SyncTable", tmpCol & " | " & StringHelper.IsEqual(tmpCol, Constants.FIELD_ID, True) & " | " & StringHelper.IsEqual(tmpCol, Constants.FIELD_TIMESTAMP, True)
                                If (Not StringHelper.IsEqual(tmpCol, Constants.FIELD_ID, True)) _
                                    And (Not StringHelper.IsEqual(tmpCol, Constants.FIELD_TIMESTAMP, True)) Then
                                    check = True
                                    If c = 0 Then
                                        '============== UPDATE SERVER RECORD BLOCK ===============
                                        tmpDataLocal.Add tmpCol, str1
                                        tmpCols.Add tmpCol
                                        tmpCols.Add Constants.FIELD_ID
                                        tmpCols.Add Constants.FIELD_TIMESTAMP
                                        tmpDataLocal.Add Constants.FIELD_ID, tmpId
                                        Logger.LogDebug "DbManager.SyncTable", "Update server. Field: " & tmpCol & ". Local: " & str1 & " | Server: " & str2
                                        UpdateServerRecord tmpDataLocal, tmpCols, fromTable, tblCached, Server, DatabaseName, Username, Password
                                        '============== UPDATE SERVER RECORD BLOCK ===============
                                    Else
                                        '============== UPDATE LOCAL RECORD & CONFLICT RECORD BLOCK ===============
                                        ' if local < server -> update local db
                                        
                                        'UpdateLocalRecord tmpDataServer, tmpCols, tblCached
                                        tmpDataLocal.Add "Table name", desTable
                                        tmpDataLocal.Add "Field name", tmpCol
                                        tmpDataLocal.Add "Field type", CStr(tmpType)
                                        tmpDataLocal.Add "Local data", str1
                                        tmpDataLocal.Add "Server data", str2
                                        tmpDataLocal.Add "Local timestamp", tmpTimestampLocal
                                        tmpDataLocal.Add "Server timestamp", tmpTimestampServer
                                        tmpDataLocal.Add "Row ID", tmpId
                                        Logger.LogDebug "DbManager.SyncTable", "Insert conflict. Field: " & tmpCol & " Local: " & str1 & " | Server: " & str2
                                        ExecuteQuery FileHelper.ReadQuery(Constants.TABLE_SYNC_CONFLICT, Constants.Q_INSERT), tmpDataLocal
                                        If Not CheckConflict Then
                                            tmpDataServer.Add Constants.FIELD_TIMESTAMP, tmpTimestampServer
                                            tmpDataServer.Add Constants.FIELD_ID, tmpId
                                            tmpCols.Add Constants.FIELD_ID
                                            tmpCols.Add Constants.FIELD_TIMESTAMP
                                            tmpCols.Add tmpCol
                                            tmpDataServer.Add tmpCol, str2
                                            UpdateLocalRecord tmpDataServer, tmpCols, tblCached
                                        End If
                                         '============== UPDATE LOCAL RECORD & CONFLICT RECORD BLOCK ===============
                                    End If
                                End If
                            End If
                        Next i
                        'Logger.LogDebug "DbManager.SyncTable", "Check: " & check & ". Compare timestamp: " & c
                        If Not check And c <> 0 Then
                            '============== ONLY UPDATE TIMESTAMP BLOCK ===============
                            ' not equal -> check for update local timestamp
                            Logger.LogDebug "DbManager.SyncTable", "Update local timestamp. Table: " & desTable & ". Row ID: " & tmpId & ". Local timestamp:" & tmpTimestampLocal & ". Server timestamp: " & tmpTimestampServer
                            Set tmpDataServer = New Scripting.Dictionary
                            Set tmpCols = New Collection
                            tmpDataServer.Add Constants.FIELD_TIMESTAMP, tmpTimestampServer
                            tmpDataServer.Add Constants.FIELD_ID, tmpId
                            tmpCols.Add Constants.FIELD_ID
                            tmpCols.Add Constants.FIELD_TIMESTAMP
                            UpdateLocalRecord tmpDataServer, tmpCols, tblCached
                            '============== ONLY UPDATE TIMESTAMP BLOCK ===============
                        End If
                    Else
                        '============== ADD NEW LOCAL RECORD BLOCK ===============
                        ' If not exist in local db. Create new record
                        Set tmpCols = New Collection
                        Set tmpDataServer = New Scripting.Dictionary
                        For i = 0 To rst.fields.count - 1
                            tmpCol = rst.fields(i).name
                            str2 = GetFieldValue(rst, tmpCol)
                            tmpDataServer.Add tmpCol, str2
                            tmpCols.Add tmpCol
                        Next i
                        CreateLocalRecord tmpDataServer, tmpCols, tblCached
                        '============== ADD NEW LOCAL RECORD BLOCK ===============
                    End If
                End If
                rst.MoveNext
            Loop
        Else
        End If
        dbs.TableDefs.Refresh
        Recycle
        Init
        '============== ADD NEW SERVER RECORD BLOCK ===============
        OpenRecordSet "SELECT * FROM [" & tblCached _
                    & "] WHERE [" & Constants.FIELD_TIMESTAMP & "] IS NULL "
        If (Not rst Is Nothing) And (Not (rst.EOF And rst.BOF)) Then
            rst.MoveFirst
            Do Until rst.EOF = True
                Set tmpDataLocal = New Scripting.Dictionary
                Set tmpCols = New Collection
                Set tmpColType = New Scripting.Dictionary
                For i = 0 To rst.fields.count - 1
                    tmpCol = rst.fields(i).name
                    tmpColType.Add tmpCol, rst.fields(i).Type
                    If Not StringHelper.IsEqual(tmpCol, Constants.FIELD_TIMESTAMP, True) Then
                        tmpCols.Add tmpCol
                        If StringHelper.IsEqual(tmpCol, Constants.FIELD_ID, True) Then
                            tmpDataLocal.Add tmpCol, StringHelper.GetGUID
                        Else
                            tmpDataLocal.Add tmpCol, GetFieldValue(rst, tmpCol)
                        End If
                    End If
                Next i
                rst.Edit
                rst(Constants.FIELD_ID).value = CStr(tmpDataLocal.Item(Constants.FIELD_ID))
                rst.Update
                CreateServerRecord tmpDataLocal, tmpColType, tmpCols, fromTable, tblCached, Server, DatabaseName, Username, Password
                rst.MoveNext
            Loop
        Else
            Logger.LogDebug "DbManager.SyncTable", "No new record need to upload to server!"
        End If
        Recycle
        '============== ADD NEW SERVER RECORD BLOCK ===============
        
        DoCmd.DeleteObject acTable, desTable
        DoCmd.Rename desTable, acTable, tblCached
    Else
        Logger.LogDebug "DbManager.SyncTable", "Table " & fromTable & " is not existed!" _
                                    & " . Create new table ..."
        ImportSqlTable Server, DatabaseName, fromTable, desTable, Username, Password
    End If
    
End Function

Private Function CreateRecordQuery(datas As Scripting.Dictionary, cols As Collection _
                                , table As String, Optional colsType As Scripting.Dictionary _
                                , Optional IsServer As Boolean) As String
    Dim query As String
    Dim tmpCol As String, tmpVal As String
    Dim i As Integer
    Dim val As Variant
    Dim value As String
    tmpCol = ""
    tmpVal = ""
    For Each val In cols
        tmpCol = tmpCol & "[" & val & "],"
        If IsServer Then
            value = datas.Item(val)
            If colsType Is Nothing Then
                tmpVal = tmpVal & "'" & StringHelper.EscapeQueryString(value) & "',"
            Else
                If colsType.Item(val) = dbBoolean Then
                    Logger.LogDebug "DbManager.CreateRecordQuery", "boolean value: " & value
                    If StringHelper.IsEqual(value, "False", True) Then
                        tmpVal = tmpVal & "'0',"
                    Else
                        tmpVal = tmpVal & "'1',"
                    End If
                Else
                    tmpVal = tmpVal & "'" & StringHelper.EscapeQueryString(value) & "',"
                End If
            End If
        End If
    Next
    tmpCol = Trim(tmpCol)
    If StringHelper.EndsWith(tmpCol, ",", True) Then
        tmpCol = Left(tmpCol, Len(tmpCol) - 1)
    End If
    tmpVal = Trim(tmpVal)
    If StringHelper.EndsWith(tmpVal, ",", True) And IsServer Then
        tmpVal = Left(tmpVal, Len(tmpVal) - 1)
    End If
    
    If IsServer Then
        query = "INSERT INTO [" & table & "](" & tmpCol & ")" & " VALUES(" & tmpVal & ")"
    Else
        query = "INSERT INTO [" & table & "](" & tmpCol & ")" & " VALUES(" & tmpCol & ")"
    End If
    Logger.LogDebug "DbManager.CreateRecordQuery", "Query: " & query
    CreateRecordQuery = query
End Function

Private Function UpdateRecordQuery(datas As Scripting.Dictionary, cols As Collection, table As String, Optional IsServer As Boolean) As String
    Dim query As String
    Dim tmpCol As String
    Dim i As Integer
    Dim val As Variant
    tmpCol = ""
    For Each val In cols
        If Not StringHelper.IsEqual(CStr(val), Constants.FIELD_ID, True) Then
            If IsServer = True And StringHelper.IsEqual(CStr(val), Constants.FIELD_TIMESTAMP, True) Then
                tmpCol = tmpCol & "[" & CStr(val) & "] = GETDATE()" & " ,"
            Else
                tmpCol = tmpCol & "[" & CStr(val) & "] = '" & StringHelper.EscapeQueryString(datas.Item(CStr(val))) & "'" & " ,"
            End If
        End If
    Next
    tmpCol = Trim(tmpCol)
    If StringHelper.EndsWith(tmpCol, ",", True) Then
        tmpCol = Left(tmpCol, Len(tmpCol) - 1)
    End If
    query = "UPDATE [" & table & "] SET " & tmpCol & "" & " WHERE [id]='" & StringHelper.EscapeQueryString(datas.Item("id")) & "'"
    Logger.LogDebug "DbManager.UpdateLocalRecord", "Query: " & query
    UpdateRecordQuery = query
End Function

Private Function CreateLocalRecord(datas As Scripting.Dictionary, cols As Collection, table As String)
    Dim query As String: query = CreateRecordQuery(datas, cols, table)
    ExecuteQuery query, datas
End Function

Private Function UpdateLocalRecord(datas As Scripting.Dictionary, cols As Collection, table As String)
    Dim query As String: query = UpdateRecordQuery(datas, cols, table)
    ExecuteQuery query, datas
End Function

Private Function CreateServerRecord(datas As Scripting.Dictionary, colsType As Scripting.Dictionary, cols As Collection, table As String, _
                                            desTable As String, _
                                            Server As String, _
                                    DatabaseName As String, _
                                    Optional Username As String, _
                                    Optional Password As String)
    On Error GoTo OnError
    Dim rs As ADODB.RecordSet
    Dim cn As ADODB.Connection
    Set cn = New ADODB.Connection
    Dim query As String
    Dim stConnect As String
    Dim tmpTimestamp As String, tmpId As String
    If Len(Username) <> 0 Then
        stConnect = "DRIVER=SQL Server;SERVER=" & Server & ";DATABASE=" & DatabaseName _
                                                & ";UID=" & Username _
                                                & ";PWD=" & Password
    Else
        stConnect = "DRIVER=SQL Server;SERVER=" & Server & ";DATABASE=" & DatabaseName
    End If
    Logger.LogDebug "DbManager.CreateServerRecord", "Connection String: " & stConnect
    query = CreateRecordQuery(datas, cols, table, colsType, True)
    cn.Open stConnect
    cn.BeginTrans
    cn.Execute query
    cn.CommitTrans
    tmpId = datas.Item(Constants.FIELD_ID)
    query = "SELECT [" & Constants.FIELD_TIMESTAMP & "] FROM " & table & " WHERE [" & Constants.FIELD_ID & "]='" & StringHelper.EscapeQueryString(tmpId) & "'"
    Logger.LogDebug "DbManager.CreateServerRecord", "Query: " & query
    Set rs = cn.Execute(query)
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        tmpTimestamp = rs(Constants.FIELD_TIMESTAMP)
        Logger.LogDebug "DbManager.CreateServerRecord", "New timestamp: " & tmpTimestamp
        query = "UPDATE [" & desTable & "] SET [" _
                                    & Constants.FIELD_TIMESTAMP & "] = '" & StringHelper.EscapeQueryString(tmpTimestamp) _
                                    & "' WHERE [" & Constants.FIELD_ID & "] = '" & StringHelper.EscapeQueryString(tmpId) & "'"
        Logger.LogDebug "DbManager.CreateServerRecord", "Query: " & query
        ExecuteQuery query
    End If
OnExit:
    On Error Resume Next
    cn.Close
    Exit Function
OnError:
    Logger.LogError "DbManager.UpdateServerRecord", "Could create table " & table & " records. Query: " & query, Err
    Resume OnExit
End Function

Private Function UpdateServerRecord(datas As Scripting.Dictionary, cols As Collection, table As String, _
                                            desTable As String, _
                                            Server As String, _
                                    DatabaseName As String, _
                                    Optional Username As String, _
                                    Optional Password As String)
    On Error GoTo OnError
    Dim rs As ADODB.RecordSet
    Dim cn As ADODB.Connection
    Set cn = New ADODB.Connection
    Dim query As String
    Dim stConnect As String
    Dim tmpTimestamp As String, tmpId As String
    If Len(Username) <> 0 Then
        stConnect = "DRIVER=SQL Server;SERVER=" & Server & ";DATABASE=" & DatabaseName _
                                                & ";UID=" & Username _
                                                & ";PWD=" & Password
    Else
        stConnect = "DRIVER=SQL Server;SERVER=" & Server & ";DATABASE=" & DatabaseName
    End If
    Logger.LogDebug "DbManager.UpdateServerRecord", "Connection String: " & stConnect
    query = UpdateRecordQuery(datas, cols, table, True)
    
    cn.Open stConnect
    cn.BeginTrans
    cn.Execute query
    cn.CommitTrans
    tmpId = datas.Item(Constants.FIELD_ID)
    query = "SELECT [" & Constants.FIELD_TIMESTAMP & "] FROM " & table & " WHERE [id]='" & StringHelper.EscapeQueryString(tmpId) & "'"
    Logger.LogDebug "DbManager.UpdateServerRecord", "Query: " & query
    Set rs = cn.Execute(query)
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        tmpTimestamp = rs(Constants.FIELD_TIMESTAMP)
        Logger.LogDebug "DbManager.UpdateServerRecord", "New timestamp: " & tmpTimestamp
        query = "UPDATE [" & desTable & "] SET [" _
                                    & Constants.FIELD_TIMESTAMP & "] = '" & StringHelper.EscapeQueryString(tmpTimestamp) _
                                    & "' WHERE [id] = '" & StringHelper.EscapeQueryString(tmpId) & "'"
        Logger.LogDebug "DbManager.UpdateServerRecord", "Query: " & query
        ExecuteQuery query
    End If
OnExit:
    On Error Resume Next
    cn.Close
    Exit Function
OnError:
    Logger.LogError "DbManager.UpdateServerRecord", "Could update table " & table & " records. Query: " & query, Err
    Resume OnExit
End Function