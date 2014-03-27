Option Explicit

Public Sub OnTest()
   ' Run
    'Dim sh As New SyncHelper
    'sh.init "mappingType"
    'sh.sync
    'sh.Recycle
    'Dim dbm As New DbManager
    'dbm.init
    'dbm.ExecuteQuery "delete [user_data].[mappingBpRoles].value from user_data"
    'dbm.ExecuteQuery "insert into user_data([mappingBpRoles].value) values('SRM Lead Requester') where ntid='ABDUST'"
    'dbm.ExecuteQuery "insert into user_data([mappingBpRoles].value) values('Contract Display & Reporting') where ntid='ABDUST'"
    'dbm.Recycle
    
    'Dim dbm As New DbManager
    'dbm.Init
    'dbm.RecycleTableName "user_data_mapping_role"
    'dbm.Recycle
    Dim tmpFid As String
    Dim tmpName As String
    Dim dbm As New DbManager
    dbm.Init
    dbm.OpenRecordSet "select * from Functions where deleted=0"
   
                If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
                    dbm.RecordSet.MoveFirst
                    Do While Not dbm.RecordSet.EOF
                        tmpFid = dbm.GetFieldValue(dbm.RecordSet, "id")
                        tmpName = dbm.GetFieldValue(dbm.RecordSet, "nameFunction")
                        MsgBox tmpFid & " " & tmpName
                        dbm.RecordSet.MoveNext
                    Loop
                End If
End Sub

Public Sub Run(Optional TestClassName As String)
    Dim r As TestRunner
    Set r = New TestRunner
    r.Run TestClassName
End Sub

Public Sub Prep(Optional className As String)
    Dim AG As AutoGen
    Set AG = New AutoGen
    AG.Prep className
End Sub

Public Function QW(s As String) As String
    QW = Chr(34) & s & Chr(34)
End Function