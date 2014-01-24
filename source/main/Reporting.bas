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

Public Sub GenerateReport(Name As String)
    Dim ss As New SystemSettings
    ss.Init
    Dim rpm As New ReportMetaData
    Logger.LogDebug "Reporting.GenerateReport", "Start init report metadata " & Name
    rpm.Init Name
    If FileHelper.IsExist(rpm.OutputPath) Then
        Logger.LogDebug "Reporting.GenerateReport", "Delete file " & rpm.OutputPath
        FileHelper.Delete rpm.OutputPath
    End If
    If rpm.Valid Then
        Logger.LogDebug "Reporting.GenerateReport", "Report metadata is valid "
        Dim i As Integer, j As Integer
        Dim oExcel As New Excel.Application
        Dim WB As New Excel.Workbook
        Dim WS As Excel.WorkSheet
        Dim rng As Excel.range
        
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
                                    headers = rSect.Header
                                    If rpm.FillHeader Then
                                        For i = LBound(headers) To UBound(headers)
                                            Set rng = .Cells(rpm.StartHeaderRow, colHeadCount)
                                            rng.value = headers(i)
                                            colHeadCount = colHeadCount + 1
                                         Next i
                                    End If
                                    Select Case rSect.SectionType
                                        Case Constants.RP_SECTION_TYPE_AUTO:
                                            Logger.LogDebug "Reporting.GenerateReport", "Generate section type: Auto"
                                            
                                            For i = LBound(headers) To UBound(headers)
                                                Dim Query As String
                                                Query = rSect.MakeQuery(headers(i), ss)
                                                Logger.LogDebug "Reporting.GenerateReport", "Prepare query: " & Query
                                                objRs.Open Query, objConn, adOpenStatic, adLockReadOnly
                                                'Logger.LogDebug "Reporting.GenerateReport", "Prepare Cells(" & CStr(rpm.StartRow) & "," & CStr(colCount) & ")"
                                                
                                                Set rng = .Cells(rpm.StartRow, colCount) 'Starting point of the data range
                                                rng.CopyFromRecordset objRs
                                                objRs.Close
                                                colCount = colCount + 1
                                            Next i
                                            
                                            Logger.LogDebug "Reporting.GenerateReport", "Complete generate section type: Auto"
                                        Case Constants.RP_SECTION_TYPE_FIXED:
                                            Logger.LogDebug "Reporting.GenerateReport", "Generate section type: Fixed"
                                            objRs.Open rSect.Query, objConn, adOpenStatic, adLockReadOnly
                                            'Logger.LogDebug "Reporting.GenerateReport", "Prepare Cells(" & CStr(rpm.StartRow) & "," & CStr(colCount) & ")"
                                            Set rng = .Cells(rpm.StartRow, colCount) 'Starting point of the data range
                                            rng.CopyFromRecordset objRs
                                            colCount = colCount + rSect.HeaderCount
                                            objRs.Close
                                            Logger.LogDebug "Reporting.GenerateReport", "Complete generate section type: Fixed"
                                        Case Else
                                    End Select
                                Next
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
    Else
        Logger.LogError "Reporting.GenerateReport", "The reporting meta data format is not valid", Nothing
    End If
End Sub