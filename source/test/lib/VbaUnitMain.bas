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
    Dim c As New CourseHelper
    c.Init
    c.PrepareCurriculumSheet
    c.Validation
    c.ImportCourse
    c.ImportMapping
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