Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @author Hai Lu
' Report meta data object
Option Explicit
Private mSectionType As String
Private mHeader() As String
Private mQuery As String
Private mValid As Boolean
Private ss As SystemSetting
Private mCount As Long

Private Property Get DataQuery() As Scripting.Dictionary
    Dim data As New Scripting.Dictionary
    Set ss = Session.Settings()
    data.Add Constants.Q_KEY_FUNCTION_REGION_ID, ss.RegionFunctionId
    data.Add Constants.Q_KEY_REGION_NAME, ss.RegionName
    data.Add Constants.Q_KEY_FUNCTION_REGION_NAME, Session.CurrentUser.FuncRegion.name
    Set DataQuery = data
End Property

Public Function Init(raw As String, Optional mss As SystemSetting)
    mValid = False
    Dim dbm As New DbManager
    Logger.LogDebug "ReportSection.Init", "Prepare raw: " & raw
    Dim i As Integer, tmpStr As String, tmpList() As String
    Dim arraySize As Integer
    mQuery = raw
    Set ss = mss
    Dim queryCache() As String
        
    If InStr(mQuery, "{%") <> 0 And InStr(mQuery, "%}") <> 0 Then
        mSectionType = Constants.RP_SECTION_TYPE_AUTO
    ElseIf InStr(mQuery, Constants.SPLIT_LEVEL_2) <> 0 Then
        mSectionType = Constants.RP_SECTION_TYPE_TMP_TABLE
    Else
        mSectionType = Constants.RP_SECTION_TYPE_FIXED
    End If
    
    Logger.LogDebug "ReportSection.Init", "Section type: " & mSectionType
            
    Select Case mSectionType
        Case Constants.RP_SECTION_TYPE_AUTO:
             Logger.LogDebug "ReportSection.Init", "RP_SECTION_TYPE_AUTO"
            ' Start generate query
            mQuery = PrepareQuery(mQuery, ss)
            mQuery = StringHelper.GenerateQuery(mQuery, DataQuery)
        Case Constants.RP_SECTION_TYPE_FIXED:
            dbm.Init
            Logger.LogDebug "ReportSection.Init", "RP_SECTION_TYPE_FIXED"
            mQuery = StringHelper.GenerateQuery(mQuery, DataQuery)
           
            dbm.OpenRecordSet mQuery
            mCount = dbm.RecordSet.RecordCount
            ' Execute query and get all header name
            Logger.LogDebug "ReportSection.Init", "Fields count: " & CStr(dbm.RecordSet.fields.Count)
            For i = 0 To dbm.RecordSet.fields.Count - 1
                ReDim Preserve mHeader(arraySize)
                mHeader(arraySize) = dbm.RecordSet.fields(i).name
                arraySize = arraySize + 1
            Next i
            Logger.LogDebug "ReportSection.Init", "Complete RP_SECTION_TYPE_FIXED"
            dbm.Recycle
        Case Constants.RP_SECTION_TYPE_TMP_TABLE:
            Logger.LogDebug "ReportSection.Init", "RP_SECTION_TYPE_TMP_TABLE"
            Dim valueCache As New Scripting.Dictionary
            Dim v As Variant
            Dim tmpRst As DAO.RecordSet
            Dim tmpQdf As DAO.QueryDef
            Dim tmpQuery As String
            Dim tmpData As Scripting.Dictionary
            Dim tmpKey As String
            Dim tmpValue As String
            Dim tableName As String
            queryCache = Split(mQuery, Constants.SPLIT_LEVEL_2)
            
            If UBound(queryCache) > 0 Then
                tmpQuery = StringHelper.GenerateQuery(StringHelper.TrimNewLine(queryCache(3)), DataQuery)
                mQuery = tmpQuery
                Logger.LogDebug "ReportSection.Init", "Primary query: " & mQuery
                tableName = StringHelper.TrimNewLine(queryCache(0))
                dbm.Init
                If Ultilities.IfTableExists(tableName) Then
                    Logger.LogDebug "ReportSection.Init", "Delete all records table " & tableName
                    dbm.ExecuteQuery "DELETE * FROM [" & tableName & "]"
                Else
                    Logger.LogDebug "ReportSection.Init", "Create new table " & tableName
                    dbm.ExecuteQuery FileHelper.ReadQuery(tableName, Constants.Q_CREATE)
                End If

                tmpQuery = StringHelper.GenerateQuery(StringHelper.TrimNewLine(queryCache(1)), DataQuery)
                Logger.LogDebug "ReportSection.Init", "Get cache value query: " & tmpQuery
                dbm.OpenRecordSet tmpQuery
                If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
                    dbm.RecordSet.MoveFirst
                    Do Until dbm.RecordSet.EOF = True
                        tmpKey = dbm.RecordSet(0)
                        Logger.LogDebug "ReportSection.Init", " tmpKey: " & tmpKey
                        tmpValue = ""
                        Set tmpData = DataQuery
                        tmpData.Add Constants.Q_KEY_VALUE, tmpKey
                        tmpQuery = StringHelper.GenerateQuery(StringHelper.TrimNewLine(queryCache(2)), tmpData)
                        Logger.LogDebug "ReportSection.Init", "Get value data query: " & tmpQuery
                        Set tmpQdf = dbm.Database.CreateQueryDef("", tmpQuery)
                        Set tmpRst = tmpQdf.OpenRecordSet
                        If Not (tmpRst.EOF And tmpRst.BOF) Then
                            tmpRst.MoveFirst
                            Do Until tmpRst.EOF = True
                                tmpValue = tmpValue & tmpRst(0) & Chr(13) & Chr(10)
                                tmpRst.MoveNext
                            Loop
                        End If
                        tmpRst.Close
                        Set tmpRst = Nothing
                        Logger.LogDebug "ReportSection.Init", "tmpValue: " & tmpValue
                        dbm.ExecuteQuery "INSERT INTO [" & tableName & "]([key],[value]) VALUES('" & tmpKey & "','" & tmpValue & "')"
                        dbm.RecordSet.MoveNext
                    Loop
                End If
                dbm.Recycle
            End If
            
            dbm.Init
            mQuery = StringHelper.GenerateQuery(mQuery, DataQuery)
            dbm.OpenRecordSet mQuery
            mCount = dbm.RecordSet.RecordCount
            ' Execute query and get all header name
            Logger.LogDebug "ReportSection.Init", "Fields count: " & CStr(dbm.RecordSet.fields.Count)
            For i = 0 To dbm.RecordSet.fields.Count - 1
                ReDim Preserve mHeader(arraySize)
                mHeader(arraySize) = dbm.RecordSet.fields(i).name
                arraySize = arraySize + 1
            Next i

            dbm.Recycle
        Case Else
    End Select
            
    Logger.LogDebug "ReportSection.Init", "Query: " & mQuery
    If HeaderCount > 0 And Len(mQuery) > 0 Then
        Logger.LogDebug "ReportSection.Init", "Found " & CStr(HeaderCount) & " header: "
        mValid = True
        For i = LBound(mHeader) To UBound(mHeader)
            Logger.LogDebug "ReportSection.Init", "- " & mHeader(i)
        Next i
    End If
End Function

Private Function PrepareQuery(query As String, Optional ss As SystemSetting) As String
    Dim arraySize As Integer
    Dim dbm As New DbManager
    Dim i As Integer
    Dim l As Long, r As Long, q As String, length As Long, strTemp As String
    Dim tmp As String, cQuery, tmpSplit() As String, qOut As String, qIn As String, tmpVal As String, tmpQuery As String

    q = query
    length = 0
    
    Do While Not InStr(q, "{%") = 0
        length = length + 1
        'Logger.LogDebug "ReportSection.PrepareQuery", "Raw query: " & q
        l = InStr(q, "{%")
        Logger.LogDebug "ReportSection.PrepareQuery", "Found start pos: " & CStr(l)
        r = InStr(l, q, "%}")
        Logger.LogDebug "ReportSection.PrepareQuery", "Found end pos: " & CStr(r)
        cQuery = Mid(q, l, r - l + 2)
        Logger.LogDebug "ReportSection.PrepareQuery", "Custom query: " & cQuery
        tmp = Trim(Mid(cQuery, 3, Len(cQuery) - 4))
        'Logger.LogDebug "ReportSection.PrepareQuery", "Prefix removed query: " & tmp
        tmpSplit = Split(tmp, "|")
        qOut = Trim(tmpSplit(0))
        qIn = Trim(tmpSplit(1))
        'Logger.LogDebug "ReportSection.PrepareQuery", "Generate query: " & qIn
        'Logger.LogDebug "ReportSection.PrepareQuery", "Get value query: " & qOut
        If StringHelper.IsContain(qOut, "select", True) And StringHelper.IsContain(qOut, "from", True) Then
            dbm.Init
            dbm.OpenRecordSet (qOut)
            If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
                tmpQuery = ""
                dbm.RecordSet.MoveFirst
                Do Until dbm.RecordSet.EOF = True
                    tmpVal = dbm.RecordSet(0)
                    ReDim Preserve mHeader(arraySize)
                    mHeader(arraySize) = tmpVal
                    arraySize = arraySize + 1
                    
                    'Logger.LogDebug "ReportSection.PrepareQuery", "Found value: " & tmpVal
                    dbm.RecordSet.MoveNext
                Loop
            Else
            End If
        Else
            Dim tmpListStr() As String
            tmpListStr = Split(qOut, ",")
            For i = LBound(tmpListStr) To UBound(tmpListStr)
                ReDim Preserve mHeader(arraySize)
                mHeader(arraySize) = Trim(Replace(Replace(Replace(tmpListStr(i), Chr(10), " "), Chr(13), " "), Chr(9), " "))
                arraySize = arraySize + 1
            Next i
        End If
        Dim tmpStr As String
        For i = LBound(mHeader) To UBound(mHeader)
            strTemp = Replace(qIn, "(%VALUE%)", StringHelper.EscapeQueryString(mHeader(i)))
            If Not ss Is Nothing Then
                strTemp = Replace(strTemp, "(%RG_F_ID%)", ss.RegionName)
            End If
            tmpQuery = tmpQuery & qIn & ","
        Next i
        'If StringHelper.EndsWith(tmpQuery, ",", True) Then
        '    tmpQuery = Left(tmpQuery, Len(tmpQuery) - 1)
        'End If
        'Logger.LogDebug "ReportSection.PrepareQuery", "tmpQuery: " & tmpQuery
        q = Replace(q, cQuery, qIn)
        dbm.Recycle
    Loop
    PrepareQuery = q
End Function

Private Function GenerateQuery(query As String, Optional ss As SystemSetting) As String
    Dim dbm As New DbManager
    Dim l As Long, r As Long, q As String, length As Long, strTemp As String
    Dim tmp As String, cQuery, tmpSplit() As String, qOut As String, qIn As String, tmpVal As String, tmpQuery As String

    q = query
    length = 0
    
    Do While Not InStr(q, "{%") = 0
        length = length + 1
        'Logger.LogDebug "ReportSection.GenerateQuery", "Raw query: " & q
        l = InStr(q, "{%")
        Logger.LogDebug "ReportSection.GenerateQuery", "Found start pos: " & CStr(l)
        r = InStr(l, q, "%}")
        Logger.LogDebug "ReportSection.GenerateQuery", "Found end pos: " & CStr(r)
        cQuery = Mid(q, l, r - l + 2)
        'Logger.LogDebug "ReportSection.GenerateQuery", "Custom query: " & cQuery
        tmp = Trim(Mid(cQuery, 3, Len(cQuery) - 4))
        'Logger.LogDebug "ReportSection.GenerateQuery", "Prefix removed query: " & tmp
        tmpSplit = Split(tmp, "|")
        qOut = Trim(tmpSplit(0))
        qIn = Trim(tmpSplit(1))
        'Logger.LogDebug "ReportSection.GenerateQuery", "Generate query: " & qIn
        'Logger.LogDebug "ReportSection.GenerateQuery", "Get value query: " & qOut
        dbm.Init
        dbm.OpenRecordSet (qOut)
        
        If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
            tmpQuery = ""
            dbm.RecordSet.MoveFirst
            Do Until dbm.RecordSet.EOF = True
                tmpVal = dbm.RecordSet(0)
                strTemp = Replace(qIn, "(%VALUE%)", StringHelper.EscapeQueryString(tmpVal))
                If Not ss Is Nothing Then
                    strTemp = Replace(strTemp, "(%RG_F_ID%)", ss.RegionName)
                End If
                tmpQuery = tmpQuery & strTemp & ","
                'Logger.LogDebug "ReportSection.GenerateQuery", "Found value: " & tmpVal
                dbm.RecordSet.MoveNext
            Loop
        Else
            
        End If
        If StringHelper.EndsWith(tmpQuery, ",", True) Then
            tmpQuery = Left(tmpQuery, Len(tmpQuery) - 1)
        End If
        Logger.LogDebug "ReportSection.GenerateQuery", "tmpQuery: " & tmpQuery
        q = Replace(q, cQuery, tmpQuery)
        dbm.Recycle
    Loop
    GenerateQuery = q
End Function


Public Property Get query() As String
    query = mQuery
End Property

Public Property Get header() As String()
    header = mHeader
End Property

Public Property Get HeaderCount() As Integer
    If Not Ultilities.IsVarArrayEmpty(mHeader) Then
        HeaderCount = UBound(mHeader) + 1
    Else
        HeaderCount = 0
    End If
End Property

Public Property Get SectionType() As String
    SectionType = mSectionType
End Property

Public Property Get Valid() As Boolean
    Valid = mValid
End Property

Public Property Get Count() As Long
    Count = mCount
End Property

Public Function MakeQuery(colName As String, Optional mss As SystemSetting) As String
    If StringHelper.IsEqual(mSectionType, Constants.RP_SECTION_TYPE_AUTO, True) Then
        Dim data As Scripting.Dictionary
        Set data = DataQuery
        data.Add Constants.Q_KEY_VALUE, StringHelper.EscapeQueryString(colName)
        MakeQuery = StringHelper.GenerateQuery(mQuery, data)
    End If
End Function