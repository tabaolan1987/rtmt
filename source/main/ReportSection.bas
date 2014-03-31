Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @author Hai Lu
' Report meta data object
Option Explicit
Private mSectionType As String
Private mHeader() As String
Private mPivotHeader() As String
Private mQuery As String
Private mValid As Boolean
Private ss As SystemSetting
Private mCount As Long
Private mCachedTable As String
Private mCategory As Scripting.Dictionary
Private mCategoryBColor As Scripting.Dictionary
Private mCategoryFColor As Scripting.Dictionary

Private Property Get DataQuery() As Scripting.Dictionary
    Dim data As New Scripting.Dictionary
    
    Set ss = Session.Settings()
    data.Add Constants.Q_KEY_FUNCTION_REGION_ID, ss.RegionFunctionId
    data.Add Constants.Q_KEY_REGION_NAME, ss.regionName
    data.Add Constants.Q_KEY_FUNCTION_REGION_NAME, Session.CurrentUser.FuncRegion.Name
    Set DataQuery = data
End Property

Public Function Init(raw As String, Optional mss As SystemSetting, Optional SkipCheckHeader As Boolean)
    Set mCategory = New Scripting.Dictionary
    Set mCategoryBColor = New Scripting.Dictionary
    Set mCategoryFColor = New Scripting.Dictionary
    mValid = False
    Dim dbm As New DbManager
    Logger.LogDebug "ReportSection.Init", "SkipCheckHeader: " & SkipCheckHeader
    Logger.LogDebug "ReportSection.Init", "Prepare raw: " & raw
    Dim i As Integer, tmpStr As String, tmpList() As String
    Dim arraySize As Integer
    mQuery = raw
    Set ss = mss
    Dim queryCache() As String
    Dim tmpRst As DAO.RecordSet
    Dim tmpQdf As DAO.QueryDef
    Dim valueCache As New Scripting.Dictionary
    Dim v As Variant
    Dim m As Variant
    Dim check As Boolean
            
    Dim tmpQuery As String
    Dim tmpData As Scripting.Dictionary
    Dim tmpKey As String
    Dim tmpValue As String
    Dim tableName As String
    
    If InStr(mQuery, "{%") <> 0 And InStr(mQuery, "%}") <> 0 Then
        mSectionType = Constants.RP_SECTION_TYPE_AUTO
    ElseIf InStr(mQuery, Constants.SPLIT_LEVEL_2) <> 0 Then
        If InStr(mQuery, Constants.RP_SECTION_TYPE_TMP_PILOT_REPORT) <> 0 Then
            mSectionType = Constants.RP_SECTION_TYPE_TMP_PILOT_REPORT
        Else
            mSectionType = Constants.RP_SECTION_TYPE_TMP_TABLE
        End If
    Else
        mSectionType = Constants.RP_SECTION_TYPE_FIXED
    End If
    
    Logger.LogDebug "ReportSection.Init", "Section type: " & mSectionType
            
    Select Case mSectionType
        Case Constants.RP_SECTION_TYPE_AUTO:
             Logger.LogDebug "ReportSection.Init", "RP_SECTION_TYPE_AUTO"
            ' Start generate query
            PrepareQuery mQuery, ss
            mQuery = StringHelper.GenerateQuery(mQuery, DataQuery)
        Case Constants.RP_SECTION_TYPE_FIXED:
            dbm.Init
            Logger.LogDebug "ReportSection.Init", "RP_SECTION_TYPE_FIXED"
            mQuery = StringHelper.GenerateQuery(mQuery, DataQuery)
           
            dbm.OpenRecordSet mQuery
            mCount = dbm.RecordSet.recordCount
            ' Execute query and get all header name
            Logger.LogDebug "ReportSection.Init", "Fields count: " & CStr(dbm.RecordSet.fields.count)
            For i = 0 To dbm.RecordSet.fields.count - 1
                ReDim Preserve mHeader(arraySize)
                mHeader(arraySize) = dbm.RecordSet.fields(i).Name
                arraySize = arraySize + 1
            Next i
            Logger.LogDebug "ReportSection.Init", "Complete RP_SECTION_TYPE_FIXED"
            dbm.Recycle
        Case Constants.RP_SECTION_TYPE_TMP_TABLE:
            Logger.LogDebug "ReportSection.Init", "RP_SECTION_TYPE_TMP_TABLE"
            
            queryCache = Split(mQuery, Constants.SPLIT_LEVEL_2)
            
            If UBound(queryCache) > 0 Then
                tmpQuery = StringHelper.GenerateQuery(StringHelper.TrimNewLine(queryCache(3)), DataQuery)
                mQuery = tmpQuery
                Logger.LogDebug "ReportSection.Init", "Primary query: " & mQuery
                tableName = StringHelper.TrimNewLine(queryCache(0))
                mCachedTable = tableName
                dbm.Init
                If Ultilities.IfTableExists(tableName) Then
                    Logger.LogDebug "ReportSection.Init", "Delete all records table " & tableName
                    dbm.ExecuteQuery "DELETE * FROM [" & tableName & "]"
                Else
                    Logger.LogDebug "ReportSection.Init", "Create new table " & tableName
                    dbm.ExecuteQuery FileHelper.ReadQuery(tableName, Constants.Q_CREATE)
                End If

                tmpQuery = StringHelper.GenerateQuery(StringHelper.TrimNewLine(queryCache(1)), DataQuery)
                'Logger.LogDebug "ReportSection.Init", "Get cache value query: " & tmpQuery
                dbm.OpenRecordSet tmpQuery
                If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
                    dbm.RecordSet.MoveFirst
                    Do Until dbm.RecordSet.EOF = True
                        tmpKey = dbm.RecordSet(0)
                 '       Logger.LogDebug "ReportSection.Init", " tmpKey: " & tmpKey
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
                  '      Logger.LogDebug "ReportSection.Init", "tmpValue: " & tmpValue
                        dbm.ExecuteQuery "INSERT INTO [" & tableName & "]([key],[value]) VALUES('" & tmpKey & "','" & tmpValue & "')"
                        dbm.RecordSet.MoveNext
                    Loop
                End If
                dbm.Recycle
            End If
            
            dbm.Init
            mQuery = StringHelper.GenerateQuery(mQuery, DataQuery)
            If Not SkipCheckHeader Then
                dbm.OpenRecordSet mQuery
                mCount = dbm.RecordSet.recordCount
                ' Execute query and get all header name
                Logger.LogDebug "ReportSection.Init", "Fields count: " & CStr(dbm.RecordSet.fields.count)
                For i = 0 To dbm.RecordSet.fields.count - 1
                    ReDim Preserve mHeader(arraySize)
                    mHeader(arraySize) = dbm.RecordSet.fields(i).Name
                    arraySize = arraySize + 1
                Next i
            End If
            dbm.Recycle
        Case RP_SECTION_TYPE_TMP_PILOT_REPORT:
            dbm.Init
            queryCache = Split(mQuery, Constants.SPLIT_LEVEL_2)
            Dim tmpCol As New Collection
            Dim tmpValueCol As Collection
            Dim tmpHeader As String
            Dim tmpQueryPilot As String
            Dim tmpColor As String
            Dim tmpNtid As String
            Dim tmpTableName As String
            Dim tmpDataIn As Scripting.Dictionary
            Dim tmpDataPara As Scripting.Dictionary
            Dim tmpCache As String
            Dim tmpMappingFields As String
            Dim pilotHeader As New Scripting.Dictionary
            tmpCol.Add "key"
            If UBound(queryCache) > 0 Then
                tmpTableName = StringHelper.TrimNewLine(queryCache(0))
                mCachedTable = tmpTableName
                dbm.DeleteTable tmpTableName
                tmpQuery = "create table [" & tmpTableName & "] ( [key] varchar(255), "
                tmpCache = StringHelper.TrimNewLine(queryCache(1))
                If StringHelper.IsContain(tmpCache, "select", True) _
                    And StringHelper.IsContain(tmpCache, "from", True) _
                    And StringHelper.IsContain(tmpCache, "where", True) _
                    Then
                    arraySize = 0
                    dbm.OpenRecordSet StringHelper.GenerateQuery(tmpCache, DataQuery)
                    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
                        dbm.RecordSet.MoveFirst
                        Do While Not dbm.RecordSet.EOF
                            tmpKey = dbm.GetFieldValue(dbm.RecordSet, "header")
                            tmpValue = dbm.GetFieldValue(dbm.RecordSet, "Category")
                            tmpColor = dbm.GetFieldValue(dbm.RecordSet, "fColor")
                            If Not mCategoryFColor.Exists(tmpValue) Then
                                mCategoryFColor.Add tmpValue, tmpColor
                            End If
                            tmpColor = dbm.GetFieldValue(dbm.RecordSet, "bColor")
                            If Not mCategoryBColor.Exists(tmpValue) Then
                                mCategoryBColor.Add tmpValue, tmpColor
                            End If
                            ReDim Preserve tmpList(arraySize)
                            tmpList(arraySize) = tmpKey
                            If Not mCategory.Exists(tmpKey) Then
                                mCategory.Add tmpKey, tmpValue
                            End If
                            
                            arraySize = arraySize + 1
                            dbm.RecordSet.MoveNext
                        Loop
                    End If
                Else
                    tmpList = Split(tmpCache, ",")
                End If
                arraySize = 0
                For i = LBound(tmpList) To UBound(tmpList)
                    tmpStr = Trim(Replace(Replace(Replace(tmpList(i), Chr(10), " "), Chr(13), " "), Chr(9), " "))
                    tmpHeader = "f" & CStr(i + 1)
                    pilotHeader.Add tmpHeader, tmpStr
                    tmpCol.Add tmpHeader
                    tmpQuery = tmpQuery & "[" & tmpHeader & "]" & " varchar(255)" & ","
                    tmpMappingFields = tmpMappingFields & "[" & tmpHeader & "] AS [" & tmpStr & "],"
                    ReDim Preserve mPivotHeader(arraySize)
                    mPivotHeader(arraySize) = tmpStr
                    arraySize = arraySize + 1
                    
                Next i
                If StringHelper.EndsWith(tmpQuery, ",", True) Then
                    tmpQuery = Left(tmpQuery, Len(tmpQuery) - 1)
                End If
                If StringHelper.EndsWith(tmpMappingFields, ",", True) Then
                    tmpMappingFields = Left(tmpMappingFields, Len(tmpMappingFields) - 1)
                End If
                tmpQuery = tmpQuery & ")"
                Logger.LogDebug "ReportSection.Init", "Create new table cache query: " & tmpQuery
                dbm.ExecuteQuery tmpQuery
                
                tmpQuery = StringHelper.GenerateQuery(StringHelper.TrimNewLine(queryCache(2)), DataQuery)
                 Logger.LogDebug "ReportSection.Init", "List all key query: " & tmpQuery
                dbm.OpenRecordSet tmpQuery
                If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
                    dbm.RecordSet.MoveFirst
                    tmpQuery = StringHelper.TrimNewLine(queryCache(3))
                    Logger.LogDebug "ReportSection.Init", "Get pilot data query: " & tmpQuery
                    Do Until dbm.RecordSet.EOF = True
                        Set tmpDataPara = DataQuery
                        Set tmpDataIn = New Scripting.Dictionary
                        tmpNtid = dbm.GetFieldValue(dbm.RecordSet, Constants.Q_KEY_VALUE)
                        tmpDataIn.Add "key", tmpNtid
                        tmpDataPara.Add Constants.Q_KEY_VALUE, tmpNtid
                        tmpQueryPilot = StringHelper.GenerateQuery(tmpQuery, tmpDataPara)
                        Logger.LogDebug "ReportSection.Init", tmpNtid & " | " & tmpQueryPilot
                        Set tmpQdf = dbm.Database.CreateQueryDef("", tmpQueryPilot)
                        Set tmpRst = tmpQdf.OpenRecordSet
                        Set tmpValueCol = New Collection
                        If Not (tmpRst.EOF And tmpRst.BOF) Then
                            tmpRst.MoveFirst
                            Do Until tmpRst.EOF = True
                                tmpValue = dbm.GetFieldValue(tmpRst, Constants.Q_KEY_VALUE)
                                tmpValueCol.Add tmpValue
                                tmpRst.MoveNext
                            Loop
                        End If
                        For Each v In pilotHeader.keys
                            tmpStr = pilotHeader.Item(CStr(v))
                            check = False
                            For Each m In tmpValueCol
                                If StringHelper.IsEqual(CStr(m), tmpStr, True) Then
                                    check = True
                                    Exit For
                                End If
                            Next m
                            If check Then
                                tmpDataIn.Add CStr(v), "Y"
                            Else
                                tmpDataIn.Add CStr(v), ""
                            End If
                        Next v
                        dbm.CreateLocalRecord tmpDataIn, tmpCol, tmpTableName
                        dbm.RecordSet.MoveNext
                    Loop
                End If
                Set tmpData = DataQuery
                tmpData.Add Constants.Q_KEY_MAPPING_FIELDS, tmpMappingFields
                Logger.LogDebug "ReportSection.Init", "tmpMappingFields: " & tmpMappingFields
                mQuery = StringHelper.GenerateQuery(StringHelper.TrimNewLine(queryCache(4)), tmpData)
                'Logger.LogDebug "ReportSection.Init", "Query: " & mQuery
                If Not SkipCheckHeader Then
                    dbm.OpenRecordSet mQuery
                    mCount = dbm.RecordSet.recordCount
                    ' Execute query and get all header name
                    Logger.LogDebug "ReportSection.Init", "Fields count: " & CStr(dbm.RecordSet.fields.count)
                    arraySize = 0
                    For i = 0 To dbm.RecordSet.fields.count - 1
                        ReDim Preserve mHeader(arraySize)
                        mHeader(arraySize) = dbm.RecordSet.fields(i).Name
                        arraySize = arraySize + 1
                    Next i
                End If
                dbm.Recycle
            End If
        Case Else
    End Select
            
    Logger.LogDebug "ReportSection.Init", "Query: " & mQuery
    If HeaderCount > 0 And Len(mQuery) > 0 Then
        mValid = True
        Logger.LogDebug "ReportSection.Init", "Found " & CStr(HeaderCount) & " header: "
        For i = LBound(mHeader) To UBound(mHeader)
            Logger.LogDebug "ReportSection.Init", "- " & mHeader(i)
        Next i
    End If
    If SkipCheckHeader Then
        mValid = True
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
            dbm.OpenRecordSet StringHelper.GenerateQuery(qOut, DataQuery)
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
                strTemp = Replace(strTemp, "(%RG_F_ID%)", ss.regionName)
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
        dbm.OpenRecordSet StringHelper.GenerateQuery(qOut, DataQuery)
        
        If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
            tmpQuery = ""
            dbm.RecordSet.MoveFirst
            Do Until dbm.RecordSet.EOF = True
                tmpVal = dbm.RecordSet(0)
                strTemp = Replace(qIn, "(%VALUE%)", StringHelper.EscapeQueryString(tmpVal))
                If Not ss Is Nothing Then
                    strTemp = Replace(strTemp, "(%RG_F_ID%)", ss.regionName)
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

Public Property Get count() As Long
    count = mCount
End Property

Public Function MakeQuery(mHeaderCol As Collection, Optional mss As SystemSetting) As String
    If StringHelper.IsEqual(mSectionType, Constants.RP_SECTION_TYPE_AUTO, True) Then
        Dim arraySize As Integer
        Dim i As Integer
        Dim v As Variant
        Dim l As Long, r As Long, q As String, length As Long, strTemp As String
        Dim tmp As String, cQuery, tmpSplit() As String, qOut As String, qIn As String, tmpVal As String, tmpQuery As String
        q = mQuery
        length = 0
        Dim tmpStr As String
        Do While Not InStr(q, "{%") = 0
            length = length + 1
            l = InStr(q, "{%")
            Logger.LogDebug "ReportSection.MakeQuery", "Found start pos: " & CStr(l)
            r = InStr(l, q, "%}")
            Logger.LogDebug "ReportSection.MakeQuery", "Found end pos: " & CStr(r)
            cQuery = Mid(q, l, r - l + 2)
            Logger.LogDebug "ReportSection.MakeQuery", "Custom query: " & cQuery
            tmp = Trim(Mid(cQuery, 3, Len(cQuery) - 4))
            tmpSplit = Split(tmp, "|")
            qOut = Trim(tmpSplit(0))
            qIn = Trim(tmpSplit(1))
            
            For Each v In mHeaderCol
                tmpStr = CStr(v)
                Logger.LogDebug "ReportSection.MakeQuery", "make query for header: " & tmpStr
                strTemp = Replace(qIn, "(%VALUE%)", StringHelper.EscapeQueryString(tmpStr))
                tmpQuery = tmpQuery & strTemp & ","
            Next v
            If StringHelper.EndsWith(tmpQuery, ",", True) Then
                tmpQuery = Left(tmpQuery, Len(tmpQuery) - 1)
            End If
            
            q = Replace(q, cQuery, tmpQuery)
        Loop
        
        MakeQuery = q
    End If
End Function

Public Property Get CachedTable() As String
    CachedTable = mCachedTable
End Property

Public Property Get PivotHeader() As String()
    PivotHeader = mPivotHeader
End Property

Public Property Get Categories() As Scripting.Dictionary
    Set Categories = mCategory
End Property

Public Property Get CategoryBColor() As Scripting.Dictionary
    Set CategoryBColor = mCategoryBColor
End Property

Public Property Get CategoryFColor() As Scripting.Dictionary
    Set CategoryFColor = mCategoryFColor
End Property

Public Function CatBcolor(CategoryName As String) As Long
    Dim colors() As String
    CatBcolor = RGB(255, 255, 255)
    If mCategoryBColor.Exists(CategoryName) Then
        On Error Resume Next
        colors = Split(mCategoryBColor.Item(CategoryName), ",")
        If UBound(colors) = 2 Then
            CatBcolor = RGB(CInt(colors(0)), CInt(colors(1)), CInt(colors(2)))
        End If
    End If
End Function

Public Function CatFcolor(CategoryName As String) As Long
    Dim colors() As String
    CatFcolor = RGB(0, 0, 0)
    If mCategoryFColor.Exists(CategoryName) Then
        On Error Resume Next
        colors = Split(mCategoryFColor.Item(CategoryName), ",")
        If UBound(colors) = 2 Then
            CatFcolor = RGB(CInt(colors(0)), CInt(colors(1)), CInt(colors(2)))
        End If
    End If
End Function
    
Public Property Get PivotHeaderCount() As Integer
    If Not Ultilities.IsVarArrayEmpty(mPivotHeader) Then
        PivotHeaderCount = UBound(mPivotHeader) + 1
    Else
        PivotHeaderCount = 0
    End If
End Property