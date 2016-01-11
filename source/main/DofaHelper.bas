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

Public Function Init(Optional path As String, Optional worksheet As String)
    If Len(worksheet) > 0 Then
        mWorksheet = worksheet
    Else
        mWorksheet = "DOFA Interface"
    End If
    If Len(path) > 0 Then
        mPath = path
    Else
        mPath = FileHelper.GetCSVFile("Select Dofa data")
    End If
    mPath = FileHelper.DuplicateAsTemporary(mPath)
    Set dbm = New DbManager
    Set mCols = New Collection
    mCols.Add "id"
    mCols.Add "sno"
    mCols.Add "username1"
    mCols.Add "DOA_SRM_Au"
    mCols.Add "Employee_G"
    mCols.Add "username2"
    mCols.Add "DOA_Spend_Limit"
    mCols.Add "Crcy"
    mCols.Add "changeOn"
    mCols.Add "timechange"
    mCols.Add "changeby"
    'mCols.Add "region"
End Function


Public Function ImportDofa()
    DoCmd.Echo False
    dbm.Init
    If Ultilities.IfTableExists("dofa") Then
        Session.UpdateDbFlag (False)
        dbm.ExecuteQuery "update dofa set deleted=-1 where deleted=0"
        'dbm.ExecuteQuery "delete from dofa"
        Session.UpdateDbFlag (True)
    Else
        dbm.ExecuteQuery FileHelper.ReadQuery("dofa", Constants.Q_CREATE)
    End If

    dbm.Init
    Dim oExcel As New Excel.Application
    Dim WB As New Excel.Workbook
    Dim WS As Excel.worksheet
    Dim rng As Excel.range
    Dim l, k As Long
    Dim tmpValue As String
    Dim query As String
    Dim c As Long
    c = 0
    
    ' Save all state
    Logger.LogDebug "DofaHelper.ImportDofa", "Save all excel state"
    Dim screenUpdateState, statusBarState, calcState, eventsState, displayPageBreakState As Boolean
    screenUpdateState = oExcel.ScreenUpdating
    Logger.LogDebug "DofaHelper.ImportDofa", "Save state ScreenUpdating"
    statusBarState = oExcel.DisplayStatusBar
    Logger.LogDebug "DofaHelper.ImportDofa", "Save state DisplayStatusBar"
    eventsState = oExcel.EnableEvents
    Logger.LogDebug "DofaHelper.ImportDofa", "Save state EnableEvents"
    
    With oExcel
            .DisplayAlerts = False
            .Visible = False
            Logger.LogDebug "DofaHelper.ImportDofa", "Open excel template: " & mPath
            Set WB = .Workbooks.Open(mPath)
            
            With WB
                Logger.LogDebug "DofaHelper.ImportDofa", "Select worksheet: " & mWorksheet
                Set WS = WB.workSheets(mWorksheet)
                'Turn off some Excel functionality so the code runs faster
                Logger.LogDebug "DofaHelper.ImportDofa", "Turn off ScreenUpdating"
                oExcel.ScreenUpdating = False
                Logger.LogDebug "DofaHelper.ImportDofa", "Turn off DisplayStatusBar"
                oExcel.DisplayStatusBar = False
                Logger.LogDebug "DofaHelper.ImportDofa", "Turn off EnableEvents"
                oExcel.EnableEvents = False
                Logger.LogDebug "DofaHelper.ImportDofa", "Turn off DisplayPageBreaks"
                WS.DisplayPageBreaks = False
                
                With WS
                    If .FilterMode Then
                        .ShowAllData
                    End If
                    l = 2
                    DoCmd.SetWarnings False
                    Application.Echo False
                    Do While l < 65000
                        Set rng = .Cells(l, 1)
                        tmpValue = Trim(rng.value)
                        If Len(tmpValue) = 0 Then
                            Exit Do
                        End If
                        Set mData = New Scripting.Dictionary
                        mData.Add "id", StringHelper.GetGUID
                        mData.Add "sno", tmpValue
                        mData.Add "username1", Trim(.Cells(l, 2).value)
                        mData.Add "DOA_SRM_Au", Trim(.Cells(l, 3).value)
                        mData.Add "Employee_G", Trim(.Cells(l, 4).value)
                        mData.Add "username2", Trim(.Cells(l, 5).value)
                        mData.Add "DOA_Spend_Limit", Trim(.Cells(l, 6).value)
                        mData.Add "Crcy", Trim(.Cells(l, 7).value)
                        mData.Add "changeOn", Trim(.Cells(l, 8).Text)
                        mData.Add "timechange", Trim(.Cells(l, 9).Text)
                        mData.Add "changeby", Trim(.Cells(l, 10).value)
                        'mData.Add "region", Session.CurrentUser.FuncRegion.Region
                        query = dbm.CreateRecordQuery(mData, mCols, "dofa")
                        On Error Resume Next
                        Forms("frm_dofa_upload").Painting = False
                        Forms("frm_Mainboard").Painting = False
                        dbm.ExecuteQuery query
                        
                        l = l + 1
                    Loop
                    DoCmd.SetWarnings True
                    Application.Echo True
                End With
                ' Restore state
                oExcel.ScreenUpdating = screenUpdateState
                oExcel.DisplayStatusBar = statusBarState
                oExcel.EnableEvents = eventsState
                WS.DisplayPageBreaks = displayPageBreakState
                
                Logger.LogDebug "DofaHelper.ImportDofa", "Close excel file " & mPath
            End With
            .Quit
        End With

    dbm.Recycle
    DoCmd.Echo True
End Function