'@author Hai Lu
'Requied:
' - Set reference to Microsoft Excel Object library
' - Set reference to Microsoft ActiveX DataObject 2.x
Option Explicit

Public Sub ExportExcelReport(sSQL As String, sFileNameTemplate As String, output As String, workSheets As String, range As String)
    
    Dim oExcel As New Excel.Application
    
    Dim WB As New Excel.Workbook
    Dim WS As Excel.worksheet
    Dim rng As Excel.range
    Dim objConn As New ADODB.Connection
    Dim objRs As New ADODB.RecordSet
    Set objConn = CurrentProject.Connection
    With oExcel
        .Visible = False
                   'Create new workbook from the template file
                    Set WB = .Workbooks.Add(sFileNameTemplate)
                            With WB
                                 Set WS = WB.workSheets(workSheets) 'Replace with the name of actual sheet
                                 With WS
                                        
                                          objRs.Open sSQL, objConn, adOpenStatic, adLockReadOnly
                                          Set rng = .range(range) 'Starting point of the data range
                                          rng.CopyFromRecordset objRs
                                          objRs.Close
                                 End With
                                 WS.SaveAs (output)
                            End With
        .Quit
    End With
     
    Set objConn = Nothing
    Set objRs = Nothing
End Sub

Public Sub CheckReport(Name As String)
    Dim sh As SyncHelper
    If StringHelper.IsEqual(Name, Constants.RP_AUDIT_LOG, True) Then
        Set sh = New SyncHelper
        sh.Init Constants.TABLE_AUDIT_LOG
        sh.sync
        sh.Recycle
    End If
    If StringHelper.IsEqual(Name, Constants.RP_USER_DATA_CHANGE_LOG, True) Then
        Set sh = New SyncHelper
        sh.Init Constants.TABLE_USER_CHANGE_LOG
        sh.sync
        sh.Recycle
    End If
End Sub

Public Sub GenerateReport(rpm As ReportMetaData)
    Dim ss As SystemSetting
    Set ss = Session.Settings()
    If FileHelper.IsExistFile(rpm.OutputPath) Then
        Logger.LogDebug "Reporting.GenerateReport", "Delete file " & rpm.OutputPath
        FileHelper.DeleteFile rpm.OutputPath
    End If
    If rpm.Valid Then
        Logger.LogDebug "Reporting.GenerateReport", "Report metadata is valid "
        Dim i As Integer, j As Integer, k As Integer
        Dim reportSects As Collection
        Dim reportSheet As Variant
        Dim oExcel As New Excel.Application
        Dim WB As New Excel.Workbook
        Dim WS As Excel.worksheet
        Dim rng As Excel.range
        Dim Pivot As Excel.PivotTable
        Dim c As Long
        Dim tmpValue As String
        Dim headerCol As Collection
        Dim objConn As New ADODB.Connection
        Dim objRs As New ADODB.RecordSet
        Set objConn = CurrentProject.Connection
        Dim recordCount As Double
        Dim colHeadCount As Long
        
        ' Save all state
        Logger.LogDebug "Reporting.GenerateReport", "Save all excel state"
        Dim screenUpdateState, statusBarState, calcState, eventsState, displayPageBreakState As Boolean
        screenUpdateState = oExcel.ScreenUpdating
        Logger.LogDebug "Reporting.GenerateReport", "Save state ScreenUpdating"
        statusBarState = oExcel.DisplayStatusBar
        Logger.LogDebug "Reporting.GenerateReport", "Save state DisplayStatusBar"
        eventsState = oExcel.EnableEvents
        Logger.LogDebug "Reporting.GenerateReport", "Save state EnableEvents"
        
        With oExcel
            .DisplayAlerts = False
            .Visible = False
            'Create new workbook from the template file
            Logger.LogDebug "Reporting.GenerateReport", "Open excel template: " & rpm.TemplateFilePath
            Set WB = .Workbooks.Add(rpm.TemplateFilePath)
                With WB
                    For Each reportSheet In rpm.ReportSheets.keys
                        Set reportSects = rpm.ReportSheets.Item(CStr(reportSheet))
                        Logger.LogDebug "Reporting.GenerateReport", "Select worksheet: " & CStr(reportSheet)
                        Set WS = WB.workSheets(CStr(reportSheet)) 'Replace with the name of actual sheet
                        'Save sheet state
                        Logger.LogDebug "Reporting.GenerateReport", "Save state DisplayPageBreaks"
                        displayPageBreakState = WS.DisplayPageBreaks
                        'Turn off some Excel functionality so the code runs faster
                        Logger.LogDebug "Reporting.GenerateReport", "Turn off ScreenUpdating"
                        oExcel.ScreenUpdating = False
                        Logger.LogDebug "Reporting.GenerateReport", "Turn off DisplayStatusBar"
                        oExcel.DisplayStatusBar = False
                        'oExcel.Calculation = xlCalculationManual
                        Logger.LogDebug "Reporting.GenerateReport", "Turn off EnableEvents"
                        oExcel.EnableEvents = False
                        Logger.LogDebug "Reporting.GenerateReport", "Turn off DisplayPageBreaks"
                        WS.DisplayPageBreaks = False
                        
                        
                        With WS
                            If .FilterMode Then
                                .ShowAllData
                            End If
                            Logger.LogDebug "Reporting.GenerateReport", "Detect query type: Section"
                            Dim rSect As ReportSection
                            Dim colCount As Long
                            Dim category As String
                            Dim lastCategory As String
                            Dim colCategory As Long
                            colCount = rpm.StartCol
                            colHeadCount = rpm.StartHeaderCol
                            colCategory = rpm.StartHeaderCol
                            For Each rSect In reportSects
                                Dim headers() As String
                                headers = rSect.header
                                If rpm.FillHeader Then
                                    For i = LBound(headers) To UBound(headers)
                                        Set rng = .Cells(rpm.StartHeaderRow, colHeadCount)
                                        tmpValue = headers(i)
                                        rng.value = tmpValue
                                        If rpm.FillCategory And rSect.Categories.count > 0 Then
                                            Set rng = .Cells(rpm.StartCategoryRow, colHeadCount)
                                            If rSect.Categories.Exists(tmpValue) Then
                                                category = rSect.Categories.Item(tmpValue)
                                                rng.value = category
                                            End If
                                            If Not StringHelper.IsEqual(category, lastCategory, True) _
                                                Or i = UBound(headers) Then
                                                If i = UBound(headers) Then
                                                    .range(.Cells(rpm.StartCategoryRow, colCategory), .Cells(rpm.StartCategoryRow, colHeadCount)).Merge
                                                    If Len(category) > 0 Then
                                                        .range(.Cells(rpm.StartCategoryRow, colCategory), .Cells(rpm.StartHeaderRow + 1, colHeadCount)).Interior.Color = rSect.CatBcolor(category)
                                                        .range(.Cells(rpm.StartCategoryRow, colCategory), .Cells(rpm.StartHeaderRow + 1, colHeadCount)).Characters.Font.Color = rSect.CatFcolor(category)
                                                        .range(.Cells(rpm.StartCategoryRow, colCategory), .Cells(rpm.StartHeaderRow + 1, colHeadCount)).Cells.Borders.Color = RGB(0, 0, 0)
                                                    End If
                                                Else
                                                    .range(.Cells(rpm.StartCategoryRow, colCategory), .Cells(rpm.StartCategoryRow, colHeadCount - 1)).Merge
                                                    If Len(lastCategory) > 0 Then
                                                        .range(.Cells(rpm.StartCategoryRow, colCategory), .Cells(rpm.StartHeaderRow + 1, colHeadCount - 1)).Interior.Color = rSect.CatBcolor(lastCategory)
                                                        .range(.Cells(rpm.StartCategoryRow, colCategory), .Cells(rpm.StartHeaderRow + 1, colHeadCount - 1)).Characters.Font.Color = rSect.CatFcolor(lastCategory)
                                                        .range(.Cells(rpm.StartCategoryRow, colCategory), .Cells(rpm.StartHeaderRow + 1, colHeadCount - 1)).Cells.Borders.Color = RGB(0, 0, 0)
                                                    End If
                                                End If
                                                colCategory = colHeadCount
                                            End If
                                            lastCategory = category
                                        End If
                                        colHeadCount = colHeadCount + 1
                                    Next i
                                    If rpm.FillCategory And rSect.Categories.count > 0 Then
                                        .Cells(rpm.StartCategoryRow, rpm.StartHeaderCol).EntireRow.AutoFit
                                    End If
                                End If
                                Select Case rSect.SectionType
                                    Case Constants.RP_SECTION_TYPE_AUTO:
                                        Dim query As String
                                        c = 0
                                        Logger.LogDebug "Reporting.GenerateReport", "Generate section type: Auto"
                                        
                                        For i = LBound(headers) To UBound(headers)
                                            If headerCol Is Nothing Then
                                                Set headerCol = New Collection
                                            End If
                                            headerCol.Add headers(i)
                                            c = c + 1
                                            If c = rpm.BulkSize Then
                                                query = rSect.MakeQuery(headerCol, ss)
                                                Logger.LogDebug "Reporting.GenerateReport", "Prepare query: " & query
                                                objRs.Open query, objConn, adOpenStatic, adLockReadOnly
                                                recordCount = objRs.recordCount
                                                Set rng = .Cells(rpm.startRow, colCount)
                                                rng.CopyFromRecordset objRs
                                                objRs.Close
                                                colCount = colCount + c
                                                c = 0
                                                Set headerCol = New Collection
                                            End If
                                        Next i
                                        If Not headerCol Is Nothing And headerCol.count > 0 Then
                                            query = rSect.MakeQuery(headerCol, ss)
                                            Logger.LogDebug "Reporting.GenerateReport", "Prepare query: " & query
                                            objRs.Open query, objConn, adOpenStatic, adLockReadOnly
                                            recordCount = objRs.recordCount
                                            Set rng = .Cells(rpm.startRow, colCount)
                                            rng.CopyFromRecordset objRs
                                            objRs.Close
                                            colCount = colCount + headerCol.count
                                        End If
                                        
                                        Logger.LogDebug "Reporting.GenerateReport", "Complete generate section type: Auto"
                                    Case Constants.RP_SECTION_TYPE_FIXED, Constants.RP_SECTION_TYPE_TMP_TABLE, Constants.RP_SECTION_TYPE_TMP_PILOT_REPORT:
                                        Logger.LogDebug "Reporting.GenerateReport", "Generate section type: " & rSect.SectionType
                                        objRs.Open rSect.query, objConn, adOpenStatic, adLockReadOnly
                                        recordCount = objRs.recordCount
                                        'Logger.LogDebug "Reporting.GenerateReport", "Prepare Cells(" & CStr(rpm.StartRow) & "," & CStr(colCount) & ")"
                                        Set rng = .Cells(rpm.startRow, colCount) 'Starting point of the data range
                                        rng.CopyFromRecordset objRs
                                        colCount = colCount + objRs.fields.count
                                        objRs.Close
                                        Logger.LogDebug "Reporting.GenerateReport", "Complete generate section type: Fixed"
                                    Case Else
                                End Select
                            Next
                            
                             If rpm.CustomMode Then
                               k = rpm.startRow
                               Dim ntid1 As String
                               Dim ntid2 As String
                               Dim courseId1 As String
                               Dim courseId2 As String
                               Dim psValue As String
                               Dim countValue As String
                               Dim startRow As Long
                               startRow = rpm.startRow
                               Dim v As Variant
                               Do While k < 65536
                                    psValue = .Cells(k, 12).value
                                    countValue = .Cells(k, 14).value
                                    ntid1 = .Cells(k, 1).value
                                    courseId1 = .Cells(k, 8).value
                                    If Len(Trim(psValue)) = 0 Then
                                        If rpm.MergeEnable Then
                                            For Each v In rpm.MergeColumes
                                                     .range(.Cells(startRow, CInt(v)), .Cells(k - 1, CInt(v))).Merge
                                            Next v
                                        End If
                                        Exit Do
                                    End If
                                    If StringHelper.IsEqual(psValue, "s", True) _
                                        And StringHelper.IsEqual(ntid1, ntid2, True) _
                                        And StringHelper.IsEqual(courseId1, courseId2, True) Then
                                        .Cells(k, 12).EntireRow.Delete
                                    Else
                                        k = k + 1
                                    End If
                                    If rpm.MergeEnable Then
                                        If Not StringHelper.IsEqual(ntid1, ntid2, True) And Len(ntid2) > 0 Then
                                            For Each v In rpm.MergeColumes
                                                 .range(.Cells(startRow, CInt(v)), .Cells(k - 1, CInt(v))).Merge
                                            Next v
                                            startRow = k
                                        End If
                                    End If
                                    courseId2 = courseId1
                                    ntid2 = ntid1
                                 Loop
                                 .Cells(rpm.startRow, 14).EntireColumn.Delete
                            End If
                            
                            If rpm.DateColumes.count > 0 Then
                                On Error Resume Next
                                For Each v In rpm.DateColumes
                                    .range(.Cells(2, CInt(v)), .Cells(65000, CInt(v))).NumberFormat = rpm.DateFormat
                                Next v
                            End If
                            '.PrintOut Copies:=1, Preview:=False, Collate:=True
                        End With
                    Next reportSheet
                    If rpm.PivotTable Then
                            Logger.LogDebug "Reporting.GenerateReport", "Select pivot worksheet: " & rpm.PivotTableWorksheet
                            Set WS = WB.workSheets(rpm.PivotTableWorksheet)
                            Logger.LogDebug "Reporting.GenerateReport", "Select pivot table: " & rpm.PivotTableName
                            Set Pivot = WS.PivotTables(rpm.PivotTableName)
                            Pivot.RefreshTable
                            Pivot.Update
                            If rpm.PivotWordWrapCols.count > 0 Then
                                For Each v In rpm.PivotWordWrapCols
                                    WS.range(WS.Cells(1, CInt(v)), WS.Cells(rpm.startRow + recordCount, CInt(v))).WrapText = True
                                Next v
                            End If
                            WS.Rows.AutoFit
                            Dim pi As PivotItem
                            On Error Resume Next
                             For Each pi In Pivot.PivotFields("NTID").PivotItems
                                 If pi.value = "(blank)" Then
                                     pi.Visible = False
                                End If
                             Next pi
                        End If
                        
                        ' Restore state
                        oExcel.ScreenUpdating = screenUpdateState
                        oExcel.DisplayStatusBar = statusBarState
                        'oExcel.Calculation = calcState
                        oExcel.EnableEvents = eventsState
                        WS.DisplayPageBreaks = displayPageBreakState

                        Logger.LogDebug "Reporting.GenerateReport", "Save report as : " & rpm.OutputPath
                        .SaveAs (rpm.OutputPath)
                    End With
                
            Logger.LogDebug "Reporting.GenerateReport", "Close excel"
            .Quit
        End With
         
        Set objConn = Nothing
        Set objRs = Nothing
        rpm.SetComplete (True)
        rpm.Recyle
    Else
        Logger.LogError "Reporting.GenerateReport", "The reporting meta data format is not valid", Nothing
    End If
End Sub