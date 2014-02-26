'@author Hai Lu
'Requied:
' - Set reference to Microsoft Excel Object library
' - Set reference to Microsoft ActiveX DataObject 2.x
Option Compare Database

Public Sub ExportExcelReport(sSQL As String, sFileNameTemplate As String, output As String, workSheets As String, range As String)
    
    Dim oExcel As New Excel.Application
    Dim WB As New Excel.Workbook
    Dim WS As Excel.WorkSheet
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

Public Sub GenerateReport(rpm As ReportMetaData)
    Dim ss As SystemSetting
    Set ss = Session.Settings()
    If FileHelper.IsExistFile(rpm.OutputPath) Then
        Logger.LogDebug "Reporting.GenerateReport", "Delete file " & rpm.OutputPath
        FileHelper.DeleteFile rpm.OutputPath
    End If
    If rpm.Valid Then
        Logger.LogDebug "Reporting.GenerateReport", "Report metadata is valid "
        Dim i As Integer, j As Integer
        Dim oExcel As New Excel.Application
        Dim WB As New Excel.Workbook
        Dim WS As Excel.WorkSheet
        Dim rng As Excel.range
        Dim c As Long
        Dim headerCol As Collection
        Dim objConn As New ADODB.Connection
        Dim objRs As New ADODB.RecordSet
        Set objConn = CurrentProject.Connection
        With oExcel
            .DisplayAlerts = False
            .Visible = False
            'Create new workbook from the template file
            Logger.LogDebug "Reporting.GenerateReport", "Open excel template: " & rpm.TemplateFilePath
            Set WB = .Workbooks.Add(rpm.TemplateFilePath)
                With WB
                    Logger.LogDebug "Reporting.GenerateReport", "Select worksheet: " & rpm.WorkSheet
                    Set WS = WB.workSheets(rpm.WorkSheet) 'Replace with the name of actual sheet
                    With WS
                        Logger.LogDebug "Reporting.GenerateReport", "Detect query type: Section"
                        Dim rSect As ReportSection
                        Dim colCount As Long
                        colCount = rpm.StartCol
                        colHeadCount = rpm.StartHeaderCol
                        For Each rSect In rpm.ReportSections
                            Dim headers() As String
                            headers = rSect.header
                            If rpm.FillHeader Then
                                For i = LBound(headers) To UBound(headers)
                                    Set rng = .Cells(rpm.StartHeaderRow, colHeadCount)
                                    rng.value = headers(i)
                                    colHeadCount = colHeadCount + 1
                                Next i
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
                                            Set rng = .Cells(rpm.startRow, colCount)
                                            rng.CopyFromRecordset objRs
                                            objRs.Close
                                            colCount = colCount + c
                                            c = 0
                                            Set headerCol = New Collection
                                        End If
                                    Next i
                                    If Not headerCol Is Nothing And headerCol.Count > 0 Then
                                        query = rSect.MakeQuery(headerCol, ss)
                                        Logger.LogDebug "Reporting.GenerateReport", "Prepare query: " & query
                                        objRs.Open query, objConn, adOpenStatic, adLockReadOnly
                                        Set rng = .Cells(rpm.startRow, colCount)
                                        rng.CopyFromRecordset objRs
                                        objRs.Close
                                        colCount = colCount + headerCol.Count
                                    End If
                                    
                                    Logger.LogDebug "Reporting.GenerateReport", "Complete generate section type: Auto"
                                Case Constants.RP_SECTION_TYPE_FIXED, Constants.RP_SECTION_TYPE_TMP_TABLE, Constants.RP_SECTION_TYPE_TMP_PILOT_REPORT:
                                    Logger.LogDebug "Reporting.GenerateReport", "Generate section type: " & rSect.SectionType
                                    objRs.Open rSect.query, objConn, adOpenStatic, adLockReadOnly
                                    
                                    'Logger.LogDebug "Reporting.GenerateReport", "Prepare Cells(" & CStr(rpm.StartRow) & "," & CStr(colCount) & ")"
                                    Set rng = .Cells(rpm.startRow, colCount) 'Starting point of the data range
                                    rng.CopyFromRecordset objRs
                                    colCount = colCount + objRs.fields.Count
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
                                    For Each v In rpm.MergeColumes
                                             .range(.Cells(startRow, CInt(v)), .Cells(k - 1, CInt(v))).Merge
                                        Next v
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
                        
                        
                        '.PrintOut Copies:=1, Preview:=False, Collate:=True
                    End With
                    Logger.LogDebug "Reporting.GenerateReport", "Save report as : " & rpm.OutputPath
                    WS.SaveAs (rpm.OutputPath)
                End With
            Logger.LogDebug "Reporting.GenerateReport", "Close excel"
            .Quit
        End With
         
        Set objConn = Nothing
        Set objRs = Nothing
        rpm.SetComplete (True)
    Else
        Logger.LogError "Reporting.GenerateReport", "The reporting meta data format is not valid", Nothing
    End If
End Sub

Public Function CountUncompleteReport() As Integer
    Dim v As Variant
    Dim i As Integer
    Dim rpm As ReportMetaData
    i = 0
    For Each v In Session.ReportMDCols.keys
        Set rpm = Session.ReportMetaData(CStr(v))
        If Not rpm.Complete Then
            i = i + 1
        End If
    Next v
    CountUncompleteReport = i
End Function