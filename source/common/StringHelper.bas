' @author Hai Lu
' Custom function with String
Option Compare Database

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
    str = Replace(str, "'", "''")
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