'@author Hai Lu
'Requied:
' - Set reference to Microsoft Excel Object library
' - Set reference to Microsoft ActiveX DataObject 2.x
Option Compare Database

Public Sub ExportExcelReport(sSQL As String, sFileNameTemplate As String, output As String, workSheets As String, range As String)
    
    Dim oExcel As New Excel.Application
    Dim WB As New Excel.Workbook
    Dim WS As Excel.Worksheet
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

Public Sub GenerateReport(name As String)
    Dim dbm As New DbManager
    Dim rmd As New ReportMetaData, l As Long, r As Long, q As String, length As Long, strTemp As String
    Dim tmp As String, cQuery, tmpSplit() As String, qOut As String, qIn As String, tmpVal As String, tmpQuery As String
    Logger.LogDebug "Reporting.GenerateReport", "Generate report name: " & name
    rmd.Init (name)
    q = rmd.query
    length = 0
    
    Do While Not InStr(q, "{%") = 0
        length = length + 1
        'Logger.LogDebug "Reporting.GenerateReport", "Raw query: " & q
        l = InStr(q, "{%")
        Logger.LogDebug "Reporting.GenerateReport", "Found start pos: " & CStr(l)
        r = InStr(l, q, "%}")
        Logger.LogDebug "Reporting.GenerateReport", "Found end pos: " & CStr(r)
        cQuery = Mid(q, l, r - l + 2)
        'Logger.LogDebug "Reporting.GenerateReport", "Custom query: " & cQuery
        tmp = Trim(Mid(cQuery, 3, Len(cQuery) - 4))
        'Logger.LogDebug "Reporting.GenerateReport", "Prefix removed query: " & tmp
        tmpSplit = Split(tmp, "|")
        qOut = Trim(tmpSplit(0))
        qIn = Trim(tmpSplit(1))
        'Logger.LogDebug "Reporting.GenerateReport", "Generate query: " & qIn
        'Logger.LogDebug "Reporting.GenerateReport", "Get value query: " & qOut
        dbm.Init
        dbm.OpenRecordSet (qOut)
        
        If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
            tmpQuery = ""
            dbm.RecordSet.MoveFirst
            Do Until dbm.RecordSet.EOF = True
                tmpVal = dbm.RecordSet("VAL_OUT")
                strTemp = Replace(qIn, "(%VAL_IN%)", StringHelper.EscapeQueryString(tmpVal))
                strTemp = Replace(strTemp, "(%VAL_COL%)", StringHelper.EscapeQueryString(tmpVal) & " " & CStr(length))
                tmpQuery = tmpQuery & strTemp & ","
                'Logger.LogDebug "Reporting.GenerateReport", "Found value: " & tmpVal
                dbm.RecordSet.MoveNext
            Loop
        Else
            
        End If
        If StringHelper.EndsWith(tmpQuery, ",", True) Then
            tmpQuery = Left(tmpQuery, Len(tmpQuery) - 1)
        End If
        'Logger.LogDebug "Reporting.GenerateReport", "tmpQuery: " & tmpQuery
        q = Replace(q, cQuery, tmpQuery)
        dbm.Recycle
    Loop
    
    Logger.LogDebug "Reporting.GenerateReport", "Output query: " & q
    Dim testTempXlsx As String: testTempXlsx = FileHelper.CurrentDbPath & Constants.END_USER_DATA_REPORTING_TEMPLATE
        Dim output As String: output = FileHelper.CurrentDbPath & Constants.END_USER_DATA_REPORTING_OUTPUT_DIR & "/test"
        FileHelper.CheckDir output
        output = output & "/" & Constants.END_USER_DATA_REPORTING_OUTPUT_FILE
        FileHelper.Delete (output)
        Reporting.ExportExcelReport q, testTempXlsx, output, "Role Mapping Template", "A5"
        check = FileHelper.IsExist(output)
End Sub