Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Public Function CreateTable()
    Dim query As String
    On Error GoTo OnError
    query = FileHelper.ReadFile(Constants.CREATE_TABLE_END_USER_QUERY)
    CurrentDb.Execute query, dbFailOnError
    CurrentDb.TableDefs.Refresh
OnExit:
    
    Exit Function
OnError:
    Logger.LogError "DbManager.CreateTable", "Could not create table end user", Err
    Resume OnExit
End Function

Public Function DropTable()
    Dim query As String
    On Error GoTo OnError
    query = FileHelper.ReadFile(Constants.DROP_TABLE_END_USER_QUERY)
    CurrentDb.Execute query, dbFailOnError
    CurrentDb.TableDefs.Refresh
OnExit:
    Exit Function
OnError:
    Logger.LogError "DbManager.DropTable", "Could not drop table end user", Err
    Resume OnExit
End Function

Public Function DeleteAllData(stTable As String)
    Dim query As String
    On Error GoTo OnError
    If Ultilities.ifTableExists(stTable) Then
        CurrentDb.Execute "DELETE * FROM [" & stTable & "]", dbFailOnError
        CurrentDb.TableDefs.Refresh
    End If
OnExit:
    Exit Function
OnError:
    Logger.LogError "DbManager.DeleteAllData", "Could not delete all table end user data", Err
    Resume OnExit
End Function

Public Function ImportData(tblName As String, csvPath As String)
    Dim db As DAO.Database
    ' Re-link the CSV Table
    Set db = CurrentDb
    On Error GoTo OnError
    If Ultilities.ifTableExists(Constants.TMP_END_USER_TABLE_NAME) Then
        db.TableDefs.Delete Constants.TMP_END_USER_TABLE_NAME
    End If
    db.TableDefs.Refresh
    DoCmd.TransferText TransferType:=acLinkDelim, TableName:=Constants.TMP_END_USER_TABLE_NAME, _
        FileName:=csvPath, HasFieldNames:=True
    db.TableDefs.Refresh
    
    ' Perform the import
    'db.Execute "INSERT INTO someTable SELECT col1, col2, ... FROM tblImport " _
       & "WHERE NOT F1 IN ('A1', 'A2', 'A3')"
OnExit:
    db.TableDefs.Delete "Name AutoCorrect Save Failures"
    db.Close:   Set db = Nothing
    Exit Function
OnError:
    If Ultilities.ifTableExists(Constants.TMP_END_USER_TABLE_NAME) Then
        db.TableDefs.Delete Constants.TMP_END_USER_TABLE_NAME
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