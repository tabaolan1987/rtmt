Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

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
    mCols.Add "region"
End Function


Public Function ImportDofa()
    dbm.Init
    If Ultilities.IfTableExists("dofa") Then
        Session.UpdateDbFlag (False)
        dbm.ExecuteQuery "update dofa set deleted=-1 where region='" & StringHelper.EscapeQueryString(Session.CurrentUser.FuncRegion.Region) & "' and deleted=0"
        Session.UpdateDbFlag (True)
    Else
        dbm.ExecuteQuery FileHelper.ReadQuery("dofa", Constants.Q_CREATE)
    End If

    dbm.Init
    Dim oExcel As New Excel.Application
    Dim WB As New Excel.Workbook
    Dim ws As Excel.worksheet
    Dim rng As Excel.range
    Dim l, k As Long
    Dim tmpValue As String
    Dim query As String
    With oExcel
            .DisplayAlerts = False
            .Visible = False
            Logger.LogDebug "DofaHelper.ImportDofa", "Open excel template: " & mPath
            Set WB = .Workbooks.Open(mPath)
            With WB
                Logger.LogDebug "DofaHelper.ImportDofa", "Select worksheet: " & mWorksheet
                Set ws = WB.workSheets(mWorksheet)
                
                With ws
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
                        mData.Add "id", StringHelper.GetGUID
                        mData.Add "sno", tmpValue
                        mData.Add "username1", Trim(.Cells(l, 2).value)
                        mData.Add "DOA_SRM_Au", Trim(.Cells(l, 3).value)
                        mData.Add "Employee_G", Trim(.Cells(l, 4).value)
                        mData.Add "username2", Trim(.Cells(l, 5).value)
                        mData.Add "DOA_Spend_Limit", Trim(.Cells(l, 6).value)
                        mData.Add "Crcy", Trim(.Cells(l, 7).value)
                        mData.Add "changeOn", Trim(.Cells(l, 8).value)
                        mData.Add "timechange", Trim(.Cells(l, 9).value)
                        mData.Add "changeby", Trim(.Cells(l, 10).value)
                        mData.Add "region", Session.CurrentUser.FuncRegion.Region
                        query = dbm.CreateRecordQuery(mData, mCols, "dofa")
                        dbm.ExecuteQuery query
                        l = l + 1
                    Loop
                End With
                Logger.LogDebug "DofaHelper.ImportDofa", "Close excel file " & mPath
            End With
            .Quit
        End With
    dbm.TableDefsRefresh
    dbm.Recycle
End Function