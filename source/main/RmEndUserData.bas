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
    Logger.LogError "RmEndUserData.CreateTable", "Could not create table end user", Err
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
    Logger.LogError "RmEndUserData.DropTable", "Could not drop table end user", Err
    Resume OnExit
End Function

Public Function DeleteAllData()
    Dim query As String
    On Error GoTo OnError
    query = FileHelper.ReadFile(Constants.DELETE_ALL_TABLE_END_USER_DATA_QUERY)
    CurrentDb.Execute query, dbFailOnError
    CurrentDb.TableDefs.Refresh
OnExit:
    Exit Function
OnError:
    Logger.LogError "RmEndUserData.DeleteAllData", "Could not delete all table end user data", Err
    Resume OnExit
End Function

Public Function ImportData(dbName As String, csvPath As String)
    
End Function