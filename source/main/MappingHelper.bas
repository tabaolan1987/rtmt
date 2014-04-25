Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @author Hai Lu
' @create_date 24/01/2014

Option Compare Database
Private dictTop As Scripting.Dictionary
Private dictTopComment As Scripting.Dictionary
Private dictLeft As Scripting.Dictionary
Private dictLeftComment As Scripting.Dictionary
Private dbm As DbManager
Private mmd As MappingMetaData
Private ss As SystemSetting
Private Valid As Boolean
Private mFilterTop As String
Private mFilterLeft As String
Private mWorkingFile As String

Public Function Init(md As MappingMetaData, Optional mss As SystemSetting, Optional filterTop As String, _
                            Optional filterLeft As String)
    Set mmd = md
    Set ss = mss
    If Len(Trim(filterTop)) = 0 Then
        mFilterTop = "'" & StringHelper.GetGUID & "'"
    Else
        mFilterTop = filterTop
    End If
    If Len(Trim(filterLeft)) = 0 Then
        mFilterLeft = "'" & StringHelper.GetGUID & "'"
    Else
        mFilterLeft = filterLeft
    End If
   
    If mmd.Valid Then
        Set dbm = New DbManager
        Set mmd = md
        If ss Is Nothing Then
            Set ss = Session.Settings()
        End If
        PrepareMappingActivitesBBJobRoles
        PrepareData
        If dictTop.count <> 0 And dictTopComment.count <> 0 _
            And dictLeft.count <> 0 And dictLeftComment.count <> 0 Then
            Valid = True
        Else
            ' Add warning here
            Valid = False
        End If
        
    Else
        Logger.LogError "MappingHelper.Init", "Mapping meta data " & mmd.Name & " is not valid", Nothing
        ' Add warning here
        Valid = False
    End If
End Function

Public Function CheckExistMapping() As Boolean
    Dim query As String
    query = FileHelper.ReadQuery(Constants.TABLE_MAPPING_SPECIALISM_ACITIVITY, Q_SELECT)
    Dim data As New Scripting.Dictionary
    data.Add Constants.Q_KEY_REGION_NAME, Session.Settings.regionName
    data.Add Constants.Q_KEY_FUNCTION_REGION_ID, Session.Settings.RegionFunctionId
    query = StringHelper.GenerateQuery(query, data)
    dbm.Init
    dbm.OpenRecordSet query
    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
        CheckExistMapping = True
    Else
        CheckExistMapping = False
    End If
    dbm.Recycle
End Function

Private Function PrepareData()
    Dim tmpId As String, tmpIdLeft As String
    Dim query As String
    Dim tmpComment As String
    Dim tmpValue As String
    Dim data As Scripting.Dictionary
    '========= TOP FIELD DATA BLOCK ========
    Set dictTop = New Scripting.Dictionary
    Set dictTopComment = New Scripting.Dictionary
    dbm.Init
    Set data = New Scripting.Dictionary
    data.Add Constants.Q_KEY_FILTER, mFilterTop
    query = mmd.query(Q_TOP, data)
    Logger.LogDebug "MappingHelper.PrepareData", "Query top: " & query
    dbm.OpenRecordSet query
    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
        
        dbm.RecordSet.MoveFirst
        Do Until dbm.RecordSet.EOF = True
            
            tmpId = dbm.GetFieldValue(dbm.RecordSet, Constants.Q_KEY_ID)
            tmpComment = dbm.GetFieldValue(dbm.RecordSet, Constants.Q_KEY_COMMENT)
            tmpValue = dbm.GetFieldValue(dbm.RecordSet, Constants.Q_KEY_VALUE)
            'Logger.LogDebug "MappingHelper.PrepareData", "Found field name: " & tmpValue & ". Comment: " & tmpComment & ". ID: " & tmpId
            dictTop.Add tmpId, tmpValue
            dictTopComment.Add tmpId, tmpComment
            dbm.RecordSet.MoveNext
        Loop
    Else
        Logger.LogError "MappingHelper.PrepareData", "No top column data found", Nothing
    End If
    dbm.Recycle
    '========= TOP FIELD DATA BLOCK ========
    
    '========= LEFT FIELD DATA BLOCK ========
    Set dictLeft = New Scripting.Dictionary
    Set dictLeftComment = New Scripting.Dictionary
    dbm.Init
    Set data = New Scripting.Dictionary
    data.Add Constants.Q_KEY_FILTER, mFilterLeft
    query = mmd.query(Q_LEFT, data)
    Logger.LogDebug "MappingHelper.PrepareData", "Query left: " & query
    dbm.OpenRecordSet query
    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
        
        dbm.RecordSet.MoveFirst
        Do Until dbm.RecordSet.EOF = True
            tmpId = dbm.GetFieldValue(dbm.RecordSet, Constants.Q_KEY_ID)
            tmpComment = dbm.GetFieldValue(dbm.RecordSet, Constants.Q_KEY_COMMENT)
            tmpValue = dbm.GetFieldValue(dbm.RecordSet, Constants.Q_KEY_VALUE)
            'Logger.LogDebug "MappingHelper.PrepareData", "Found field name: " & tmpValue & ". Comment: " & tmpComment & ". ID: " & tmpId
            dictLeft.Add tmpId, tmpValue
            dictLeftComment.Add tmpId, tmpComment
            dbm.RecordSet.MoveNext
        Loop
    Else
        Logger.LogError "MappingHelper.PrepareData", "No left column data found", Nothing
    End If
    dbm.Recycle
    '========= LEFT FIELD DATA BLOCK ========
End Function

Public Function GenerateMapping()
    If Not mmd.Complete Then
        Dim mappingChar As String
        Dim i As Long, j As Long, l As Long, k As Long
        Dim tmpId As String, tmpComment As String, tmpValue As String
        Dim tmpData As Scripting.Dictionary
        Dim v As Variant, y As Variant
        Dim oExcel As New Excel.Application
        Dim WB As New Excel.Workbook
        Dim ws As Excel.worksheet
        Dim rng As Excel.range
        Dim check As String
    
        With oExcel
            .DisplayAlerts = False
            .Visible = False
            'Create new workbook from the template file
            mWorkingFile = mmd.TemplateFilePath
            Logger.LogDebug "MappingHelper.GenerateMapping", "Open excel template: " & mWorkingFile
            Set WB = .Workbooks.Open(mWorkingFile)
            With WB
                Logger.LogDebug "MappingHelper.GenerateMapping", "Select worksheet: " & mmd.worksheet
                Set ws = WB.workSheets(mmd.worksheet)
                With ws
                    ' Fill top field heading
                    l = mmd.StartColTop
                    For Each v In dictTop.keys
                        tmpId = CStr(v)
                        tmpValue = dictTop.Item(tmpId)
                        tmpComment = dictTopComment.Item(tmpId)
                        Set rng = .Cells(mmd.StartRowTop, l)
                        rng.value = tmpValue
                        If Len(Trim(tmpComment)) <> 0 Then
                            rng.ClearComments
                            rng.AddComment tmpComment
                            rng.Locked = True
                        End If
                        l = l + 1
                    Next v
                    ' Fill left field heading
                    l = mmd.StartRowLeft
                    For Each v In dictLeft.keys
                        tmpId = CStr(v)
                        tmpValue = dictLeft.Item(tmpId)
                        tmpComment = dictLeftComment.Item(tmpId)
                        Set rng = .Cells(l, mmd.StartColLeft)
                        rng.value = tmpValue
                        If Len(Trim(tmpComment)) <> 0 Then
                            rng.ClearComments
                            rng.AddComment tmpComment
                            rng.Locked = True
                        End If
                        l = l + 1
                    Next v
                    ' Fill mapping
                    l = mmd.StartColTop
                    For Each v In dictTop.keys
                        k = mmd.StartRowLeft
                        tmpId = CStr(v)
                        For Each y In dictLeft.keys
                            tmpIdLeft = CStr(y)
                            Set tmpData = New Scripting.Dictionary
                            tmpData.Add Q_KEY_FUNCTION_REGION_ID, ss.RegionFunctionId
                            tmpData.Add Q_KEY_ID_LEFT, tmpIdLeft
                            tmpData.Add Q_KEY_ID_TOP, tmpId
                            dbm.Init
                            dbm.OpenRecordSet mmd.query(Q_CHECK, tmpData)
                            If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
                                dbm.RecordSet.MoveFirst
                                Do Until dbm.RecordSet.EOF = True
                                    check = dbm.GetFieldValue(dbm.RecordSet, Constants.FIELD_DELETED)
                                    'Logger.LogDebug "MappingHelper.GenerateMapping", "Check: " & check
                                    If StringHelper.IsEqual(check, "false", True) Then
                                        If Len(mmd.mappingChar) > 0 Then
                                            mappingChar = mmd.mappingChar
                                        Else
                                            mappingChar = dbm.GetFieldValue(dbm.RecordSet, Constants.FIELD_MAPPING_CHAR)
                                        End If
                                        Set rng = .Cells(k, l)
                                        rng.value = mappingChar
                                        'rng.Borders.LineStyle = xlContinuous
                                        'rng.Interior.Color = RGB(255, 255, 0)
                                    End If
                                    dbm.RecordSet.MoveNext
                                Loop
                            End If
                            dbm.Recycle
                            k = k + 1
                        Next y
                        l = l + 1
                    Next v
                End With
                Logger.LogDebug "MappingHelper.GenerateMapping", "Save & Close excel file " & mmd.TemplateFilePath
                .SaveAs mmd.TemplateFilePath
            End With
            .Quit
        End With
        mmd.RefreshLastModified
        mmd.SetComplete (True)
        TimerHelper.Sleep 1000
    End If
End Function

Public Function OpenMapping()
    Dim oExcel As New Excel.Application
    Dim WB As New Excel.Workbook
    Dim ws As Excel.worksheet
    Dim rng As Excel.range
    With oExcel
        .Visible = True
        If .CommandBars("Ribbon").Height >= 150 Then
            oExcel.SendKeys "^{F1}"
        End If
        'Create new workbook from the template file
        mWorkingFile = mmd.TemplateFilePath & Constants.FILE_EXTENSION_TEMPLATE
        Logger.LogDebug "MappingHelper.OpenMapping", "Open excel template: " & mWorkingFile
        Set WB = .Workbooks.Open(mWorkingFile)
    End With
End Function

Public Function Wait()
    ' Waiting forever :)
    WaitForFileClose WorkingFile, 0, 0
End Function

Public Function ParseMapping()
    If Not StringHelper.IsEqual(mmd.CurrentModifedDate, mmd.LastModified, True) Then
        
        Dim i As Integer, j As Integer, l As Long, k As Long
        Dim tmpId As String, tmpComment As String, tmpValue As String
        Dim tmpIdLeft As String, tmpIdTop As String
        Dim tmpData As Scripting.Dictionary
        Dim tmpCols As Collection
        Dim v As Variant, y As Variant
        Dim oExcel As New Excel.Application
        Dim WB As New Excel.Workbook
        Dim ws As Excel.worksheet
        Dim rng As Excel.range
        Dim check As String
        Dim needUpdate As Boolean
        Dim IsChecked As Boolean
        Dim tmpDictLeft As New Scripting.Dictionary
        Dim tmpDictTop As New Scripting.Dictionary
        With oExcel
            .DisplayAlerts = False
            .Visible = False
            'Create new workbook from the template file
            mWorkingFile = mmd.TemplateFilePath & Constants.FILE_EXTENSION_TEMPLATE
            Logger.LogDebug "MappingHelper.ParseMapping", "Open excel template: " & mWorkingFile
            Set WB = .Workbooks.Open(mWorkingFile)
            With WB
                Logger.LogDebug "MappingHelper.ParseMapping", "Select worksheet: " & mmd.worksheet
                Set ws = WB.workSheets(mmd.worksheet)
                
                With ws
                    If .FilterMode Then
                        .ShowAllData
                    End If
                    l = mmd.StartColTop
                    For Each v In dictTop.keys
                        tmpDictTop.Add l, CStr(v)
                        l = l + 1
                    Next v
                    For k = mmd.StartRowLeft To (mmd.StartRowLeft + dictLeft.count - 1)
                        Set rng = .Cells(k, mmd.StartColLeft)
                        tmpValue = rng.value
                        tmpDictLeft.Add k, StringHelper.GetDictKey(dictLeft, tmpValue)
                    Next
                    For l = mmd.StartColTop To (mmd.StartColTop + dictTop.count - 1)
                        For k = mmd.StartRowLeft To (mmd.StartRowLeft + dictLeft.count - 1)
                            Set rng = .Cells(k, l)
                            tmpValue = Trim(rng.value)
                            tmpIdLeft = tmpDictLeft.Item(k)
                            tmpIdTop = tmpDictTop.Item(l)
                            ' Check database if is contain old mapping
                            Set tmpData = New Scripting.Dictionary
                            tmpData.Add Q_KEY_FUNCTION_REGION_ID, ss.RegionFunctionId
                            tmpData.Add Q_KEY_ID_LEFT, tmpIdLeft
                            tmpData.Add Q_KEY_ID_TOP, tmpIdTop
                            
                            dbm.Init
                            dbm.OpenRecordSet mmd.query(Q_CHECK, tmpData)
                            If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
                                needUpdate = True
                            Else
                                needUpdate = False
                            End If
                            dbm.Recycle
                            ' Check the mapping
                            If Len(tmpValue) <> 0 Then
                                tmpData.Add Q_KEY_VALUE, tmpValue
                                Logger.LogDebug "MappingHelper.ParseMapping", "Found check mark. Row: " & CStr(k) & ". Col: " & CStr(l)
                                Logger.LogDebug "MappingHelper.ParseMapping", "Left: " & dictLeft.Item(tmpIdLeft) & ". Top: " & dictTop.Item(tmpIdTop)
                                IsChecked = True
                            Else
                                tmpData.Add Q_KEY_VALUE, ""
                                IsChecked = False
                            End If
                            If needUpdate Then
                                If IsChecked Then
                                    tmpData.Add Q_KEY_CHECK, "0"
                                Else
                                    tmpData.Add Q_KEY_CHECK, "-1"
                                End If
                                
                                dbm.Init
                                Logger.LogDebug "MappingHelper.ParseMapping", "Update local database"
                                dbm.ExecuteQuery mmd.query(Q_UPDATE, tmpData)
                                dbm.Recycle
                            ElseIf IsChecked Then
                                tmpData.Add Q_KEY_ID, StringHelper.GetGUID
                                dbm.Init
                                Logger.LogDebug "MappingHelper.ParseMapping", "Create local database"
                                dbm.ExecuteQuery mmd.query(Q_INSERT, tmpData)
                                dbm.Recycle
                            End If
                        Next k
                    Next l
                End With
                Logger.LogDebug "MappingHelper.ParseMapping", "Close excel file " & mmd.TemplateFilePath
            End With
            .Quit
        End With
        TimerHelper.Sleep 1000
        mmd.RefreshLastModified
    End If
End Function

Public Function PrepareMappingActivitesBBJobRoles()
    Dim str1 As String
    Dim str2 As String
    Dim str3 As String
    Dim query As String
    Dim filter As String
    Dim rps As ReportSection
    Dim rm As ReportMetaData
    Set rm = Session.ReportMetaData(Constants.RP_END_USER_TO_BB_JOB_ROLE)
    For Each rps In rm.ReportSheets.Item("Role Mapping Report")
            filter = StringHelper.GenerateFilter(rps.PivotHeader)
            Exit For
    Next
    If Len(filter) = 0 Then
        filter = "'" & StringHelper.GetGUID & "'"
    End If
    Dim tmpRst As DAO.RecordSet
    Dim tmpQdf As DAO.QueryDef
    query = "select MA.idActivity, MA.idBpRoleStandard, MA.Description from (MappingActivityBpStandardRole as MA inner join BpRoleStandard AS BR on BR.id = MA.idBpRolestandard)" _
            & " where BR.BpRoleStandardName in (" & filter & ")" _
            & " and MA.deleted = 0"
    dbm.Init
    dbm.OpenRecordSet query
    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
    Else
        mmd.SetComplete (False)
        query = "select MA.idActivity, MA.idBpRoleStandard, MA.Description from (MappingActivityBpStandardRole as MA inner join BpRoleStandard AS BR on BR.id = MA.idBpRolestandard)" _
            & " where BR.BpRoleStandardName in (" & filter & ")" _
            & " and MA.deleted = 0"
        Set tmpQdf = dbm.Database.CreateQueryDef("", query)
        Set tmpRst = tmpQdf.OpenRecordSet
        If Not (tmpRst.EOF And tmpRst.BOF) Then
            tmpRst.MoveFirst
            Do Until tmpRst.EOF = True
                str1 = dbm.GetFieldValue(tmpRst, "idActivity")
                str2 = dbm.GetFieldValue(tmpRst, "idBpRoleStandard")
                str3 = dbm.GetFieldValue(tmpRst, "Description")
                dbm.ExecuteQuery "insert into MappingActivityBpStandardRole(id, idActivity, idBpRoleStandard,Description, deleted)" _
                    & " values('" _
                    & StringHelper.EscapeQueryString(StringHelper.GetGUID) & "', '" _
                    & StringHelper.EscapeQueryString(str1) & "', '" _
                    & StringHelper.EscapeQueryString(str2) & "','" _
                    & StringHelper.EscapeQueryString(str3) & "', '0')"
                tmpRst.MoveNext
            Loop
        End If
    End If
    dbm.Recycle
End Function

Public Property Get WorkingFile() As String
    WorkingFile = mWorkingFile
End Property