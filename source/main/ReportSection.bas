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

Public Function Init(raw As String)
    mValid = False
    Dim dbm As New DbManager
    Logger.LogDebug "ReportSection.Init", "Prepare raw: " & raw
    Dim i As Integer, tmpStr As String, tmpList() As String
    Dim arraySize As Integer
    mQuery = raw
    If InStr(mQuery, "{%") <> 0 And InStr(mQuery, "%}") <> 0 Then
        mSectionType = Constants.RP_SECTION_TYPE_AUTO
    Else
        mSectionType = Constants.RP_SECTION_TYPE_FIXED
    End If
    
    
            Logger.LogDebug "ReportSection.Init", "Section type: " & mSectionType
            
            Select Case mSectionType
                Case Constants.RP_SECTION_TYPE_AUTO:
                    ' Start generate query
                    mQuery = PrepareQuery(mQuery)
                Case Constants.RP_SECTION_TYPE_FIXED:
                    dbm.Init
                    dbm.OpenRecordSet mQuery
                    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
                        ' Execute query and get all header name
                        Logger.LogDebug "ReportSection.Init", "Fields count: " & CStr(dbm.RecordSet.fields.count)
                        For i = 0 To dbm.RecordSet.fields.count - 1
                            ReDim Preserve mHeader(arraySize)
                            mHeader(arraySize) = dbm.RecordSet.fields(i).name
                            arraySize = arraySize + 1
                        Next i
                    End If
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

Private Function PrepareQuery(query As String) As String
    Dim arraySize As Integer
    Dim dbm As New DbManager
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
                strTemp = Replace(qIn, "(%VALUE%)", StringHelper.EscapeQueryString(tmpVal))
                tmpQuery = tmpQuery & qIn & ","
                'Logger.LogDebug "ReportSection.PrepareQuery", "Found value: " & tmpVal
                dbm.RecordSet.MoveNext
            Loop
        Else
            
        End If
        'If StringHelper.EndsWith(tmpQuery, ",", True) Then
        '    tmpQuery = Left(tmpQuery, Len(tmpQuery) - 1)
        'End If
        'Logger.LogDebug "ReportSection.PrepareQuery", "tmpQuery: " & tmpQuery
        q = Replace(q, cQuery, qIn)
        dbm.Recycle
    Loop
    PrepareQuery = q
End Function

Private Function GenerateQuery(query As String) As String
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
                tmpQuery = tmpQuery & strTemp & ","
                'Logger.LogDebug "ReportSection.GenerateQuery", "Found value: " & tmpVal
                dbm.RecordSet.MoveNext
            Loop
        Else
            
        End If
        If StringHelper.EndsWith(tmpQuery, ",", True) Then
            tmpQuery = Left(tmpQuery, Len(tmpQuery) - 1)
        End If
        'Logger.LogDebug "ReportSection.GenerateQuery", "tmpQuery: " & tmpQuery
        q = Replace(q, cQuery, tmpQuery)
        dbm.Recycle
    Loop
    GenerateQuery = q
End Function


Public Property Get query() As String
    query = mQuery
End Property

Public Property Get Header() As String()
    Header = mHeader
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

Public Function MakeQuery(colName As String) As String
    If StringHelper.IsEqual(mSectionType, Constants.RP_SECTION_TYPE_AUTO, True) Then
        MakeQuery = Replace(mQuery, "(%VALUE%)", StringHelper.EscapeQueryString(colName))
    End If
End Function