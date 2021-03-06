Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private dbm As DbManager
Private qdf As DAO.QueryDef
Private rst As DAO.RecordSet
Private mPath As String
Private mWorksheet As String
Private mCols As Collection
Private mData As Scripting.Dictionary
Private missingbbRolesDb As Scripting.Dictionary
Private missingbbRolesCur As Scripting.Dictionary

Public Function Init(Optional path As String, Optional worksheet As String)
    If Len(worksheet) > 0 Then
        mWorksheet = worksheet
    Else
        mWorksheet = "Course to Roles"
    End If
    If Len(path) > 0 Then
        mPath = path
    Else
        mPath = FileHelper.GetCSVFile("Select Curriculum")
    End If
    mPath = FileHelper.DuplicateAsTemporary(mPath)
    Set dbm = New DbManager
    dbm.Init
    
    If Ultilities.IfTableExists("tmp_curriculum") Then
        dbm.RecycleTableName "tmp_curriculum"
    Else
        dbm.ExecuteQuery FileHelper.ReadQuery("tmp_curriculum", Constants.Q_CREATE)
    End If
    dbm.Init
    dbm.TableDefsRefresh
    Set mCols = New Collection
    mCols.Add "Course ID"
    mCols.Add "Course Title"
    mCols.Add "Course Duration"
    mCols.Add "Spare Column"
    mCols.Add "Role Name"
    mCols.Add "P/S"
    mCols.Add "Roles Concatenate"
    mCols.Add "Course Type"
    mCols.Add "For sorting only"
    mCols.Add "Delivery Timing"
    mCols.Add "Area"
    Logger.LogDebug "test", "step1"
End Function

Public Function PrepareCurriculumSheet()
    dbm.Init
    Dim oExcel As New Excel.Application
    Dim WB As New Excel.Workbook
    Dim WS As Excel.worksheet
    Dim rng As Excel.range
    Dim l, k As Long
    Dim tmpValue As String
    Dim query As String
    
    ' Save all state
    Logger.LogDebug "CourseHelper.PrepareCurriculumSheet", "Save all excel state"
    Dim screenUpdateState, statusBarState, calcState, eventsState, displayPageBreakState As Boolean
    screenUpdateState = oExcel.ScreenUpdating
    Logger.LogDebug "CourseHelper.PrepareCurriculumSheet", "Save state ScreenUpdating"
    statusBarState = oExcel.DisplayStatusBar
    Logger.LogDebug "CourseHelper.PrepareCurriculumSheet", "Save state DisplayStatusBar"
    eventsState = oExcel.EnableEvents
    Logger.LogDebug "CourseHelper.PrepareCurriculumSheet", "Save state EnableEvents"
    
    With oExcel
            .DisplayAlerts = False
            .Visible = False
            Logger.LogDebug "CourseHelper.PrepareCurriculumSheet", "Open excel template: " & mPath
            Set WB = .Workbooks.Open(mPath)
            With WB
                Logger.LogDebug "CourseHelper.PrepareCurriculumSheet", "Select worksheet: " & mWorksheet
                Set WS = WB.workSheets(mWorksheet)
                
                'Turn off some Excel functionality so the code runs faster
                Logger.LogDebug "CourseHelper.PrepareCurriculumSheet", "Turn off ScreenUpdating"
                oExcel.ScreenUpdating = False
                Logger.LogDebug "CourseHelper.PrepareCurriculumSheet", "Turn off DisplayStatusBar"
                oExcel.DisplayStatusBar = False
                Logger.LogDebug "CourseHelper.PrepareCurriculumSheet", "Turn off EnableEvents"
                oExcel.EnableEvents = False
                Logger.LogDebug "CourseHelper.PrepareCurriculumSheet", "Turn off DisplayPageBreaks"
                WS.DisplayPageBreaks = False
                
                With WS
                    If .FilterMode Then
                        .ShowAllData
                    End If
                    l = 2
                    Do While l < 65000
                        Set rng = .Cells(l, 1)
                        tmpValue = Trim(rng.value)
                        If Len(tmpValue) = 0 Then
                            Exit Do
                        End If
                        Set mData = New Scripting.Dictionary
                        mData.Add "Course ID", tmpValue
                        mData.Add "Course Title", Trim(.Cells(l, 2).value)
                        mData.Add "Course Duration", Trim(.Cells(l, 3).value)
                        mData.Add "Spare Column", Trim(.Cells(l, 4).value)
                        mData.Add "Role Name", Trim(.Cells(l, 5).value)
                        mData.Add "P/S", Trim(.Cells(l, 6).value)
                        mData.Add "Roles Concatenate", Trim(.Cells(l, 7).value)
                        mData.Add "Course Type", Trim(.Cells(l, 8).value)
                        mData.Add "For sorting only", Trim(.Cells(l, 9).value)
                        mData.Add "Delivery Timing", Trim(.Cells(l, 10).value)
                        mData.Add "Area", Trim(.Cells(l, 11).value)
                        query = dbm.CreateRecordQuery(mData, mCols, "tmp_curriculum")
                        dbm.ExecuteQuery query
                        l = l + 1
                    Loop
                End With
                 ' Restore state
                oExcel.ScreenUpdating = screenUpdateState
                oExcel.DisplayStatusBar = statusBarState
                oExcel.EnableEvents = eventsState
                WS.DisplayPageBreaks = displayPageBreakState
                
                Logger.LogDebug "CourseHelper.PrepareCurriculumSheet", "Close excel file " & mPath
            End With
            .Quit
        End With
    dbm.TableDefsRefresh
    dbm.Recycle
End Function

Public Function CourseCount() As Long
    Dim ct As Long
    ct = 0
    dbm.Init
    dbm.OpenRecordSet "select count(*) as ct from tmp_curriculum"
    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
        dbm.RecordSet.MoveFirst
        ct = CLng(dbm.GetFieldValue(dbm.RecordSet, "ct"))
        
    End If
    dbm.RecycleRecordSet
    CourseCount = ct
End Function

Public Function Validation()
    
    dbm.RecycleTableName "w_in_db_not_in_curriculum"
    dbm.RecycleTableName "w_in_curriculum_not_in_db"
    dbm.Init
    Set missingbbRolesCur = New Scripting.Dictionary
    Dim role As String
    dbm.OpenRecordSet "select BpRoleStandard.BpRoleStandardName from BpRoleStandard left join tmp_curriculum on BpRoleStandard.BpRoleStandardName = tmp_curriculum.[Role Name] " _
                        & " where tmp_curriculum.[Role Name] Is Null and BpRoleStandard.deleted=0"
    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
        dbm.RecordSet.MoveFirst
        Do While Not dbm.RecordSet.EOF
            role = dbm.GetFieldValue(dbm.RecordSet, "BpRoleStandardName")
            dbm.ExecuteQuery "insert into w_in_db_not_in_curriculum(role) values('" & StringHelper.EscapeQueryString(role) & "')"
            Logger.LogDebug "CourseHelper.Validation", "Not found in curiculum role: " & role
            missingbbRolesCur.Add role, role
            dbm.RecordSet.MoveNext
        Loop
    End If
    dbm.RecycleRecordSet
    Set missingbbRolesDb = New Scripting.Dictionary
    dbm.OpenRecordSet "select tmp_curriculum.[Role Name] from tmp_curriculum left join BpRoleStandard on BpRoleStandard.BpRoleStandardName = tmp_curriculum.[Role Name] " _
                        & " where BpRoleStandard.BpRoleStandardName Is Null or BpRoleStandard.deleted=-1 " _
                        & " group by tmp_curriculum.[Role Name]"
    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
        dbm.RecordSet.MoveFirst
        Do While Not dbm.RecordSet.EOF
            role = dbm.GetFieldValue(dbm.RecordSet, "Role Name")
            dbm.ExecuteQuery "insert into w_in_curriculum_not_in_db(role) values('" & StringHelper.EscapeQueryString(role) & "')"
            Logger.LogDebug "CourseHelper.Validation", "Not found in db role: " & role
            missingbbRolesDb.Add role, role
            dbm.RecordSet.MoveNext
        Loop
    End If
    dbm.Recycle
End Function

Public Function ImportCourse()
    Dim tmpId As String
    Dim query As String
    Dim tmpCourseId As String
    Set mCols = New Collection
    mCols.Add "id"
    mCols.Add "courseId"
    mCols.Add "courseArena"
    mCols.Add "courseTitle"
    mCols.Add "courseType"
    mCols.Add "courseDuration"
    mCols.Add "courseDelivery"
    mCols.Add "idRegion"
    mCols.Add "idFunction"
    mCols.Add "spare"
    mCols.Add "deleted"
    dbm.Init
    dbm.ExecuteQuery "update course set deleted=-1 where idRegion='" _
                       & StringHelper.EscapeQueryString(Session.currentUser.FuncRegion.Region) _
                       & "'" _
                      & " and idFunction='" _
                      & StringHelper.EscapeQueryString(Session.currentUser.FuncRegion.FuncRgID) & "' and deleted=0"
    dbm.OpenRecordSet "select [Course ID], [Course Title], [Course Duration], [Spare Column]," _
      & "[Course Type],[Delivery Timing],[Area] from tmp_curriculum " _
      & " group by [Course ID], [Course Title], [Course Duration], [Spare Column]," _
      & "[Course Type],[Delivery Timing],[Area]"
    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
        dbm.RecordSet.MoveFirst
        Do While Not dbm.RecordSet.EOF
            Set mData = New Scripting.Dictionary
            tmpCourseId = dbm.GetFieldValue(dbm.RecordSet, "Course ID")
            mData.Add "courseId", tmpCourseId
            mData.Add "courseTitle", dbm.GetFieldValue(dbm.RecordSet, "Course Title")
            mData.Add "courseDuration", dbm.GetFieldValue(dbm.RecordSet, "Course Duration")
            mData.Add "spare", dbm.GetFieldValue(dbm.RecordSet, "Spare Column")
            mData.Add "courseType", dbm.GetFieldValue(dbm.RecordSet, "Course Type")
            mData.Add "courseDelivery", dbm.GetFieldValue(dbm.RecordSet, "Delivery Timing")
            mData.Add "courseArena", dbm.GetFieldValue(dbm.RecordSet, "Area")
            mData.Add "idRegion", Session.currentUser.FuncRegion.Region
            mData.Add "idFunction", Session.currentUser.FuncRegion.FuncRgID
            mData.Add "deleted", "0"
            Set qdf = dbm.Database.CreateQueryDef("", "select [id] from course where courseId='" _
                       & StringHelper.EscapeQueryString(tmpCourseId) & "' and idRegion='" _
                       & StringHelper.EscapeQueryString(Session.currentUser.FuncRegion.Region) _
                       & "'" _
                      & " and idFunction='" _
                      & StringHelper.EscapeQueryString(Session.currentUser.FuncRegion.FuncRgID) & "'")
            Set rst = qdf.OpenRecordSet
            If Not (rst.EOF And rst.BOF) Then
                mData.Add "id", dbm.GetFieldValue(rst, "id")
                query = dbm.UpdateRecordQuery(mData, mCols, "course")
            Else
                mData.Add "id", StringHelper.GetGUID
                query = dbm.CreateRecordQuery(mData, mCols, "course")
            End If
            rst.Close
            Set rst = Nothing
            dbm.ExecuteQuery query
            dbm.RecordSet.MoveNext
        Loop
    End If
    dbm.Recycle
End Function

Public Function ImportMapping()
    Dim tmpId As String
    Dim query As String
    Dim tmpIdCourse As String
    Dim tmpIdBpRole As String
    Set mCols = New Collection
    mCols.Add "id"
    mCols.Add "idCourse"
    mCols.Add "idBpRole"
    mCols.Add "ps"
    mCols.Add "idRegion"
    mCols.Add "idFunction"
    mCols.Add "deleted"
    dbm.Init
    dbm.ExecuteQuery "update CourseMappingBpRoleStandard set deleted=-1 where idRegion='" _
                       & StringHelper.EscapeQueryString(Session.currentUser.FuncRegion.Region) _
                       & "'" _
                      & " and idFunction='" _
                      & StringHelper.EscapeQueryString(Session.currentUser.FuncRegion.FuncRgID) & "' and deleted=0"
    dbm.OpenRecordSet "select BpRoleStandard.id as [idBpRole], course.id As [idCourse], tmp_curriculum.[P/S] from ((BpRoleStandard inner join tmp_curriculum on BpRoleStandard.BpRoleStandardName = tmp_curriculum.[Role Name]) " _
                        & " inner join course on course.courseId = tmp_curriculum.[Course ID])" _
                        & " where BpRoleStandard.deleted=0 and course.deleted=0 " _
                        & " and course.idRegion='" _
                        & StringHelper.EscapeQueryString(Session.currentUser.FuncRegion.Region) _
                        & "'" _
                        & " and course.idFunction='" _
                        & StringHelper.EscapeQueryString(Session.currentUser.FuncRegion.FuncRgID) & "'"
    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
        dbm.RecordSet.MoveFirst
        Do While Not dbm.RecordSet.EOF
            Set mData = New Scripting.Dictionary
            tmpIdCourse = dbm.GetFieldValue(dbm.RecordSet, "idCourse")
            tmpIdBpRole = dbm.GetFieldValue(dbm.RecordSet, "idBpRole")
            mData.Add "idCourse", tmpIdCourse
            mData.Add "idBpRole", tmpIdBpRole
            mData.Add "ps", dbm.GetFieldValue(dbm.RecordSet, "P/S")
            mData.Add "idRegion", Session.currentUser.FuncRegion.Region
            mData.Add "idFunction", Session.currentUser.FuncRegion.FuncRgID
            mData.Add "deleted", "0"
            Set qdf = dbm.Database.CreateQueryDef("", "select [id] from CourseMappingBpRoleStandard where idCourse='" _
                       & StringHelper.EscapeQueryString(tmpIdCourse) & "'" _
                       & " and idBpRole='" & StringHelper.EscapeQueryString(tmpIdBpRole) & "'" _
                       & " and idRegion='" _
                       & StringHelper.EscapeQueryString(Session.currentUser.FuncRegion.Region) _
                       & "'" _
                      & " and idFunction='" _
                      & StringHelper.EscapeQueryString(Session.currentUser.FuncRegion.FuncRgID) & "'")
            Set rst = qdf.OpenRecordSet
            If Not (rst.EOF And rst.BOF) Then
                mData.Add "id", dbm.GetFieldValue(rst, "id")
                query = dbm.UpdateRecordQuery(mData, mCols, "CourseMappingBpRoleStandard")
            Else
                mData.Add "id", StringHelper.GetGUID
                query = dbm.CreateRecordQuery(mData, mCols, "CourseMappingBpRoleStandard")
            End If
            rst.Close
            Set rst = Nothing
            dbm.ExecuteQuery query
            dbm.RecordSet.MoveNext
        Loop
    End If
    dbm.Recycle
End Function

Public Function GetMissingbbRolesDb() As Scripting.Dictionary
    Set GetMissingbbRolesDb = missingbbRolesDb
End Function

Public Function GetMissingbbRolesCur() As Scripting.Dictionary
    Set GetMissingbbRolesCur = missingbbRolesCur
End Function

Public Function Recycle()
     FileHelper.DeleteFile mPath
     dbm.Recycle
End Function