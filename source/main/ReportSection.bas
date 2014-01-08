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
    Logger.LogDebug "ReportSection.Init", "Prepare raw: " & raw
    Dim tmpSplit() As String, i As Integer, tmpStr As String, tmpList() As String
    Dim arraySize As Integer
    tmpSplit = Split(raw, Constants.RP_SPLIT_LEVEL_2)
    If UBound(tmpSplit) = 2 Then
            mSectionType = Replace(tmpSplit(0), vbCrLf, " ")
            mSectionType = Trim(mSectionType)
            mQuery = Trim(tmpSplit(2))
            Logger.LogDebug "ReportSection.Init", "Section type: " & mSectionType
            
            Select Case mSectionType
                Case Constants.RP_SECTION_TYPE_AUTO:
                    tmpStr = tmpSplit(1)
                    Dim dbm As New DbManager
                    dbm.Init
                    dbm.OpenRecordSet tmpStr
                    If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
                        dbm.RecordSet.MoveFirst
                        Do Until dbm.RecordSet.EOF = True
                            ReDim Preserve mHeader(arraySize)
                            mHeader(arraySize) = dbm.RecordSet(0)
                            arraySize = arraySize + 1
                            dbm.RecordSet.MoveNext
                        Loop
                    End If
                    ' Start generate query
                    mQuery = GenerateQuery(mQuery)
                Case Constants.RP_SECTION_TYPE_FIXED:
                    tmpList = Split(tmpSplit(1), vbCrLf)
                    Logger.LogDebug "ReportSection.Init", "fixed size: " & CStr(UBound(tmpList))
                    For i = LBound(tmpList) To UBound(tmpList)
                        tmpStr = Trim(tmpList(i))
                        If Len(tmpStr) <> 0 Then
                            ReDim Preserve mHeader(arraySize)
                            mHeader(arraySize) = tmpStr
                            arraySize = arraySize + 1
                        End If
                    Next i
                    
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
    End If
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
                tmpVal = dbm.RecordSet("VAL_OUT")
                strTemp = Replace(qIn, "(%VAL_IN%)", StringHelper.EscapeQueryString(tmpVal))
                strTemp = Replace(strTemp, "(%VAL_COL%)", StringHelper.EscapeQueryString(tmpVal) & " " & CStr(length))
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
        HeaderCount = UBound(mHeader)
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