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
Private mmd As MappingMetadata
Private ss As SystemSettings
Private valid As Boolean
Private mWorkingFile As String

Public Function Init(md As MappingMetadata, Optional mss As SystemSettings)
    Set mmd = md
    Set ss = mss
    If mmd.valid Then
        Set dbm = New DbManager
        Set mmd = md
        If ss Is Nothing Then
            Set ss = New SystemSettings
            ss.Init
        End If
        PrepareData
        If dictTop.Count <> 0 And dictTopComment.Count <> 0 _
            And dictLeft.Count <> 0 And dictLeftComment.Count <> 0 Then
            valid = True
        Else
            ' Add warning here
            valid = False
        End If
        
    Else
        Logger.LogError "MappingHelper.Init", "Mapping meta data " & mmd.name & " is not valid", Nothing
        ' Add warning here
        valid = False
    End If
End Function

Private Function PrepareData()
    Dim tmpId As String, tmpIdLeft As String
    Dim tmpComment As String
    Dim tmpValue As String
    '========= TOP FIELD DATA BLOCK ========
    Set dictTop = New Scripting.Dictionary
    Set dictTopComment = New Scripting.Dictionary
    dbm.Init
    dbm.OpenRecordSet mmd.Query(Q_TOP)
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
    dbm.OpenRecordSet mmd.Query(Q_LEFT)
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
    Dim i As Long, j As Long, l As Long, k As Long
    Dim tmpId As String, tmpComment As String, tmpValue As String
    Dim tmpData As Scripting.Dictionary
    Dim v As Variant, y As Variant
    Dim oExcel As New Excel.Application
    Dim WB As New Excel.Workbook
    Dim WS As Excel.WorkSheet
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
            Logger.LogDebug "MappingHelper.GenerateMapping", "Select worksheet: " & mmd.WorkSheet
            Set WS = WB.workSheets(mmd.WorkSheet)
            With WS
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
                        dbm.OpenRecordSet mmd.Query(Q_CHECK, tmpData)
                        If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
                            dbm.RecordSet.MoveFirst
                            check = dbm.GetFieldValue(dbm.RecordSet, Constants.FIELD_DELETED)
                            Logger.LogDebug "MappingHelper.GenerateMapping", "Check: " & check
                            If StringHelper.IsEqual(check, "false", True) Then
                                Set rng = .Cells(k, l)
                                rng.value = Chr(251)
                                'rng.Borders.LineStyle = xlContinuous
                                rng.Interior.Color = RGB(255, 255, 0)
                            End If
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
End Function

Public Function OpenMapping()
    Dim oExcel As New Excel.Application
    Dim WB As New Excel.Workbook
    Dim WS As Excel.WorkSheet
    Dim rng As Excel.range
    With oExcel
        .Visible = True
        If .CommandBars("Ribbon").Height >= 150 Then
            oExcel.SendKeys "^{F1}"
        End If
        'Create new workbook from the template file
        mWorkingFile = mmd.TemplateFilePath & Constants.FILE_EXTENSION_TEMPLATE
        Logger.LogDebug "MappingHelper.ParseMapping", "Open excel template: " & mWorkingFile
        Set WB = .Workbooks.Open(mWorkingFile)
    End With
End Function

Public Function Wait()
    ' Waiting forever :)
    WaitForFileClose WorkingFile, 0, 0
End Function

Public Function ParseMapping()
    Dim i As Integer, j As Integer, l As Long, k As Long
    Dim tmpId As String, tmpComment As String, tmpValue As String
    Dim tmpIdLeft As String, tmpIdTop As String
    Dim tmpData As Scripting.Dictionary
    Dim tmpCols As Collection
    Dim v As Variant, y As Variant
    Dim oExcel As New Excel.Application
    Dim WB As New Excel.Workbook
    Dim WS As Excel.WorkSheet
    Dim rng As Excel.range
    Dim check As String
    Dim NeedUpdate As Boolean
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
            Logger.LogDebug "MappingHelper.ParseMapping", "Select worksheet: " & mmd.WorkSheet
            Set WS = WB.workSheets(mmd.WorkSheet)
            With WS
                l = mmd.StartColTop
                For Each v In dictTop.keys
                    tmpDictTop.Add l, CStr(v)
                    l = l + 1
                Next v
                For k = mmd.StartRowLeft To (mmd.StartRowLeft + dictLeft.Count - 1)
                    Set rng = .Cells(k, mmd.StartColLeft)
                    tmpValue = rng.value
                    tmpDictLeft.Add k, StringHelper.GetDictKey(dictLeft, tmpValue)
                Next
                For l = mmd.StartColTop To (mmd.StartColTop + dictTop.Count - 1)
                    For k = mmd.StartRowLeft To (mmd.StartRowLeft + dictLeft.Count - 1)
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
                        dbm.OpenRecordSet mmd.Query(Q_CHECK, tmpData)
                        If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
                            NeedUpdate = True
                        Else
                            NeedUpdate = False
                        End If
                        dbm.Recycle
                        ' Check the mapping
                        If Len(tmpValue) <> 0 Then
                            Logger.LogDebug "MappingHelper.ParseMapping", "Found check mark. Row: " & CStr(k) & ". Col: " & CStr(l)
                            Logger.LogDebug "MappingHelper.ParseMapping", "Specialism: " & dictLeft.Item(tmpIdLeft) & ". Activity: " & dictTop.Item(tmpIdTop)
                            IsChecked = True
                        Else
                            IsChecked = False
                        End If
                        If NeedUpdate Then
                            If IsChecked Then
                                tmpData.Add Q_KEY_CHECK, "0"
                            Else
                                tmpData.Add Q_KEY_CHECK, "-1"
                            End If
                            dbm.Init
                            Logger.LogDebug "MappingHelper.ParseMapping", "Update local database"
                            dbm.ExecuteQuery mmd.Query(Q_UPDATE, tmpData)
                            dbm.Recycle
                        ElseIf IsChecked Then
                            tmpData.Add Q_KEY_ID, StringHelper.GetGUID
                            dbm.Init
                            Logger.LogDebug "MappingHelper.ParseMapping", "Create local database"
                            dbm.ExecuteQuery mmd.Query(Q_INSERT, tmpData)
                            dbm.Recycle
                        End If
                    Next k
                Next l
            End With
            Logger.LogDebug "MappingHelper.ParseMapping", "Close excel file " & mmd.TemplateFilePath
        End With
        .Quit
    End With
End Function

Public Property Get WorkingFile() As String
    WorkingFile = mWorkingFile
End Property