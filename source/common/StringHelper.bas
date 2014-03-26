' @author Hai Lu
' Custom function with String
Option Compare Database
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
#If VBA7 Then
    Private Declare PtrSafe Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long
#Else
    Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long
#End If

Public Function EncodeXml(entry As String) As String
    Dim returnVal As String
    returnVal = entry
    returnVal = Replace(returnVal, "&", "&amp;")
    returnVal = Replace(returnVal, """", "&quot;")
    returnVal = Replace(returnVal, "'", "&apos;")
    returnVal = Replace(returnVal, "<", "&lt;")
    returnVal = Replace(returnVal, ">", "&gt;")
    EncodeXml = returnVal
End Function

Public Function IsContain(source As String, find As String, ignoreCase As Boolean) As Boolean
    Dim pos As Integer
    If ignoreCase = True Then
        pos = InStr(1, source, find, vbTextCompare)
    Else
        pos = InStr(1, source, find, vbBinaryCompare)
    End If
    
    If pos = 0 Then
        IsContain = False
    Else
        IsContain = True
    End If
End Function

Public Function IsEqual(str1 As String, str2 As String, ignoreCase As Boolean) As Boolean
    Dim pos As Integer
    If ignoreCase = True Then
        pos = StrComp(str1, str2, vbTextCompare)
    Else
        pos = StrComp(str1, str2, vbBinaryCompare)
    End If
    
    If pos = 0 Then
        IsEqual = True
    Else
        IsEqual = False
    End If
End Function

Public Function StartsWith(ByVal strValue As String, _
  CheckFor As String, ignoreCase As Boolean) As Boolean
  Dim sCompare As String
  Dim lLen As Long
  lLen = Len(CheckFor)
  If lLen > Len(strValue) Then Exit Function
  sCompare = Left(strValue, lLen)
  StartsWith = IsEqual(sCompare, CheckFor, ignoreCase)
End Function

Public Function EndsWith(ByVal strValue As String, _
   CheckFor As String, ignoreCase As Boolean) As Boolean
  Dim sCompare As String
  Dim lLen As Long
  lLen = Len(CheckFor)
  If lLen > Len(strValue) Then Exit Function
  sCompare = Right(strValue, lLen)
  EndsWith = IsEqual(sCompare, CheckFor, ignoreCase)
End Function

Function FixQuote(FQText As String) As String
    On Error GoTo OnError
    FixQuote = Replace(FQText, "'", "''")
    FixQuote = Replace(FixQuote, """", """""")
OnExit:
    Exit Function
OnError:
    Logger.LogError "StringHelper.FixQuote", "Could not fix quote of string: " & FQText, Err
    Resume OnExit
    Resume 0
End Function

Function EscapeQueryString(str As String) As String
    If IsContain(str, "'", True) Then
        str = Replace(str, "'", "''")
    End If
    str = Replace(str, Chr(13) & Chr(10), "")
    EscapeQueryString = str
End Function


Public Function EncodeURL( _
   StringToEncode As String, _
   Optional UsePlusRatherThanHexForSpace As Boolean = False _
) As String

  Dim TempAns As String
  Dim CurChr As Integer
  CurChr = 1

  Do Until CurChr - 1 = Len(StringToEncode)
    Select Case Asc(Mid(StringToEncode, CurChr, 1))
      Case 48 To 57, 65 To 90, 97 To 122
        TempAns = TempAns & Mid(StringToEncode, CurChr, 1)
      Case 32
        If UsePlusRatherThanHexForSpace = True Then
          TempAns = TempAns & "+"
        Else
          TempAns = TempAns & "%" & Hex(32)
        End If
      Case Else
        TempAns = TempAns & "%" & _
          Right("0" & Hex(Asc(Mid(StringToEncode, _
          CurChr, 1))), 2)
    End Select

    CurChr = CurChr + 1
  Loop

  EncodeURL = TempAns
End Function

Public Function GetGUID() As String
    '(c) 2000 Gus Molina
    Dim udtGUID As GUID
    If (CoCreateGuid(udtGUID) = 0) Then
        GetGUID = _
            String(8 - Len(Hex$(udtGUID.Data1)), "0") & Hex$(udtGUID.Data1) & _
            String(4 - Len(Hex$(udtGUID.Data2)), "0") & Hex$(udtGUID.Data2) & _
            String(4 - Len(Hex$(udtGUID.Data3)), "0") & Hex$(udtGUID.Data3) & _
            IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex$(udtGUID.Data4(0)) & _
            IIf((udtGUID.Data4(1) < &H10), "0", "") & Hex$(udtGUID.Data4(1)) & _
            IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex$(udtGUID.Data4(2)) & _
            IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex$(udtGUID.Data4(3)) & _
            IIf((udtGUID.Data4(4) < &H10), "0", "") & Hex$(udtGUID.Data4(4)) & _
            IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex$(udtGUID.Data4(5)) & _
            IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex$(udtGUID.Data4(6)) & _
            IIf((udtGUID.Data4(7) < &H10), "0", "") & Hex$(udtGUID.Data4(7))
    End If
End Function

Public Function GetDictKey(dict As Scripting.Dictionary, value As String) As String
    Dim i As Integer
    Dim key As String
    Dim v As Variant
    key = ""
    For Each v In dict.keys
        If StringHelper.IsEqual(value, dict.Item(CStr(v)), True) Then
            key = CStr(v)
            Exit For
        End If
    Next v
    GetDictKey = key
End Function

Public Function GenerateQuery(query As String, Optional data As Scripting.Dictionary) As String
    Dim mQuery As String
    mQuery = query
    If Not data Is Nothing Then
        For Each v In data.keys
            If StringHelper.IsEqual(CStr(v), Constants.Q_KEY_FILTER, True) Then
                mQuery = Replace(mQuery, "(%" & CStr(v) & "%)", data.Item(CStr(v)))
            Else
                mQuery = Replace(mQuery, "(%" & CStr(v) & "%)", EscapeQueryString(data.Item(CStr(v))))
            End If
        Next v
    End If
    'Logger.LogDebug "StringHelper.GenerateQuery", mQuery
    GenerateQuery = mQuery
End Function

Public Function TrimNewLine(str As String) As String
    Dim tmp As String
    tmp = Replace(str, Chr(13) & Chr(10), " ")
    tmp = Replace(tmp, Chr(10) & Chr(13), " ")
    TrimNewLine = Trim(tmp)
End Function

Public Function GenerateFilter(source() As String) As String
    Dim i As Integer
    Dim filter As String
    filter = ""
    For i = LBound(source) To UBound(source)
        filter = filter & "'" & StringHelper.EscapeQueryString(source(i)) & "',"
    Next i
    If StringHelper.EndsWith(filter, ",", True) Then
        filter = Left(filter, Len(filter) - 1)
    End If
    If Len(filter) > 0 Then
        GenerateFilter = filter
    Else
        GenerateFilter = "'" & StringHelper.EscapeQueryString(StringHelper.GetGUID) & "'"
    End If
End Function

Public Function GenerateFilterDict(source As Scripting.Dictionary, Optional UseKey As Boolean) As String
    Dim v As Variant
    Dim filter As String
    Dim mKey As String
    Dim mValue As String
    filter = ""
    For Each v In source.keys
        mKey = CStr(v)
        mValue = source.Item(mKey)
        Logger.LogDebug "StringHelper.GenerateFilterDict", "key: " & mKey
        If UserKey Then
            filter = filter & "'" & EscapeQueryString(mKey) & "',"
        Else
            filter = filter & "'" & EscapeQueryString(mValue) & "',"
        End If
    Next v
    If StringHelper.EndsWith(filter, ",", True) Then
        filter = Left(filter, Len(filter) - 1)
    End If
    Logger.LogDebug "StringHelper.GenerateFilterDict", "filer: " & filter
    If Len(filter) > 0 Then
        GenerateFilterDict = filter
    Else
        GenerateFilterDict = "'" & StringHelper.EscapeQueryString(StringHelper.GetGUID) & "'"
    End If
End Function