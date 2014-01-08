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

Public Function ExecuteQuery(Query As String, Optional params As Scripting.Dictionary)
    On Error GoTo OnError
    Dim key As String, value As Variant
    Set qdf = dbs.CreateQueryDef("", Query)
    If Not params Is Nothing Then
        'Logger.LogDebug "DbManager.OpenRecordSet", "Param cound: " & params.count
        For i = 0 To params.count - 1
            On Error Resume Next
            key = params.Keys(i)
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
    Logger.LogError "DbManager.ExecuteQuery", "Could execute query: " & Query, Err
    Resume OnExit
End Function

Public Function OpenRecordSet(Query As String, Optional params As Scripting.Dictionary)
    Dim prm As DAO.Parameter, i As Integer, key As String
    On Error GoTo OnError
    Set qdf = dbs.CreateQueryDef("", Query)
    If Not params Is Nothing Then
        
        'Logger.LogDebug "DbManager.OpenRecordSet", "Param cound: " & params.count
        For i = 0 To params.count - 1
            On Error Resume Next
            key = params.Keys(i)
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
    Logger.LogError "DbManager.OpenRecordSet", "Could execute query: " & Query, Err
    Resume OnExit
End Function

Public Function RecycleTable(s As SystemSettings)
    Dim i As Integer
    Dim TableNames() As String
    Dim tblName As Variant, tmp As String
    
    TableNames = s.TableNames
    Logger.LogDebug "DbManager.RecycleTable", "Start check table. Size: " & CStr(UBound(TableNames))
    Dim Query As String
    
    For i = LBound(TableNames) To UBound(TableNames)
        tmp = Trim(CStr(TableNames(i)))
        Logger.LogDebug "DbManager.RecycleTable", "Check table name " & tmp
        If Ultilities.ifTableExists(tmp) = True Then
            'ExecuteQuery FileHelper.ReadQuery(tmp, Constants.Q_DELETE_ALL)
            DoCmd.DeleteObject acTable, tmp
        End If
        
        ExecuteQuery FileHelper.ReadQuery(tmp, Constants.Q_CREATE)
    Next
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
    ' Init database
    Init
    ' Init settings
    s.Init
    '
    RecycleTable s
    ' Read the dict mapping
    Set dictMapping = s.SyncUsers
    ' Read query insert user data
    queryInsertData = FileHelper.ReadQuery(Constants.END_USER_DATA_TABLE_NAME, Constants.Q_INSERT)
    ' Open tmp table user data from CSV file
    OpenRecordSet "select * from " & Constants.TMP_END_USER_TABLE_NAME

    If Not (rst.EOF And rst.BOF) Then
        rst.MoveFirst
        Do Until rst.EOF = True
            Dim tmpList() As String, arraySize As Integer
            Logger.LogDebug "DbManager.SyncUserData", "###########################################"
            ' List all mapping column
            Set dictParams = New Scripting.Dictionary
            For i = 0 To dictMapping.count - 1
                 key = dictMapping.Keys(i)
                 value = dictMapping.Items(i)
                 If Not (StringHelper.StartsWith(key, "insert into", True)) Then
                    ' Add params value
                    On Error Resume Next
                    Logger.LogDebug "DbManager.SyncUserData", key & " = " & rst(key)
                    dictParams.Add value, rst(key)
                 Else
                    If flag = False Then
                        ' Add custom insert key to list
                        ReDim Preserve tmpList(arraySize)
                        tmpList(arraySize) = key
                        arraySize = arraySize + 1
                    End If
                 End If
            Next i
            ' Add custom insert key one time
            flag = True
            ' Insert data to user_data table
            ExecuteQuery queryInsertData, dictParams
            ' Insert mapping data
            For i = LBound(tmpList) To UBound(tmpList)
                key = tmpList(i)
                tmpSplit = Split(key, "|")
                queryCustomInsert = Trim(tmpSplit(0))
                checkValue = Trim(tmpSplit(1))
                Logger.LogDebug "DbManager.SyncUserData", "custom import query: " & queryCustomInsert & ". Check value: " & checkValue & " . From list: " & value
                value = dictMapping.Items(key)
                tmpValues = Split(value, ",")
                Logger.LogDebug "DbManager.SyncUserData", "Number of column to check: " & CStr(UBound(tmpValues) + 1)
                ' Loop to check all column
                For j = LBound(tmpValues) To UBound(tmpValues)
                    tmpValue = Trim(tmpValues(j))
                    'Logger.LogDebug "DbManager.SyncUserData", "column name: " & tmpValue
                    If StringHelper.IsEqual(rst(tmpValue), checkValue, True) Then
                        ' If value is valid, get parameter and execute query
                        Set tmpDict = New Scripting.Dictionary
                        For k = 0 To dictParams.count - 1
                            tmpDict.Add dictParams.Keys(k), dictParams.Items(k)
                        Next k
                        tmpDict.Add "value", tmpValue
                        tmpDict.Add "region_name", s.RegionName
                        ExecuteQuery queryCustomInsert, tmpDict
                    End If
                Next j
            Next i
            rst.MoveNext
        Loop
    Else
        Logger.LogInfo "DbManager.SyncUserData", "There are no records in table " & Constants.TMP_END_USER_TABLE_NAME
    End If
End Function

Public Function ImportData(tblName As String, csvPath As String)
    On Error GoTo OnError
    If Ultilities.ifTableExists(Constants.TMP_END_USER_TABLE_NAME) Then
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
    If Ultilities.ifTableExists(Constants.TMP_END_USER_TABLE_NAME) Then
        DoCmd.DeleteObject acTable, Constants.TMP_END_USER_TABLE_NAME
    End If
    Logger.LogError "DbManager.ImportData", "Could not import table " _
                        & tblName & " data from CSV file " & csvPath, Err
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
    If Ultilities.ifTableExists(desTable) Then
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
OnExit:
    On Error GoTo Quit
    If check = True Then
        If Ultilities.ifTableExists(desTable & "_tmp") Then
            Logger.LogDebug "DbManager.ImportSqlTable", "Delete cached table " & desTable & "_tmp"
            DoCmd.DeleteObject acTable, desTable & "_tmp"
        End If
    Else
        If Ultilities.ifTableExists(desTable & "_tmp") Then
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